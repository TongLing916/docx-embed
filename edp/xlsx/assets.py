"""Extract embedded assets (images, charts, OLE objects, equations) from an XLSX
workbook.

Mirrors the DOCX asset-extraction pattern but routes through the XLSX drawing
relationship chain:

* ``xl/_rels/workbook.xml.rels`` → sheet → ``xl/worksheets/sheetN.xml``
* ``xl/worksheets/_rels/sheetN.xml.rels`` → ``xl/drawings/drawingM.xml``
* ``xl/drawings/_rels/drawingM.xml.rels`` → ``xl/media/*`` / ``xl/charts/*`` /
  ``xl/embeddings/*``

Anchors live in ``xl/drawings/drawingM.xml`` as ``oneCellAnchor`` /
``twoCellAnchor`` elements carrying the cell coordinate the asset is pinned to.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
import json
from pathlib import Path, PurePosixPath
import zipfile
from xml.etree import ElementTree as ET

import olefile
from openpyxl.utils import get_column_letter

from edp.extractor import (
    _detect_mime,
    _normalize_attachment_payload,
    _read_content_types,
    _safe_filename,
    _suffix_or_default,
)


REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SPREADSHEETDRAW_NS = (
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
)
MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

REL_ID_ATTR = f"{{{REL_NS}}}id"
REL_EMBED_ATTR = f"{{{REL_NS}}}embed"

EQUATION_CLS_MARKERS = ("equation", "mathtype")


@dataclass
class XlsxAssetCollection:
    """Container for everything extracted from an XLSX workbook's drawing tree."""

    images: list[dict] = field(default_factory=list)
    charts: list[dict] = field(default_factory=list)
    attachments: list[dict] = field(default_factory=list)
    equations: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def all_assets(self) -> list[dict]:
        return [*self.images, *self.charts, *self.attachments, *self.equations]


def extract_xlsx_assets(
    xlsx_path: str | Path, assets_dir: str | Path
) -> XlsxAssetCollection:
    """Extract images, charts, OLE objects and equations from an XLSX workbook.

    Writes binaries/XML under ``{assets_dir}/{media,charts,embeddings}/`` using
    ``image_NN`` / ``chart_NN`` / ``attachment_NN`` naming. Returns a collection
    of plain-dict records (JSON-serializable) carrying anchor + sha256 metadata.
    """

    source = Path(xlsx_path)
    base = Path(assets_dir)
    media_dir = base / "media"
    charts_dir = base / "charts"
    embeddings_dir = base / "embeddings"
    resources_dir = base / "resources"
    for directory in (media_dir, charts_dir, embeddings_dir, resources_dir):
        directory.mkdir(parents=True, exist_ok=True)

    collection = XlsxAssetCollection()
    with zipfile.ZipFile(source) as zf:
        content_types = _read_content_types(zf)
        sheet_files = _sheet_files(zf, collection.warnings)
        sheet_by_drawing = {
            drawing: sheet_name
            for sheet_name, _, drawings in sheet_files
            for drawing in drawings
        }

        written: dict[str, Path] = {}  # sha256 → first written path

        for sheet_name, sheet_file, _drawings in sheet_files:
            sheet_rels = _read_rels(zf, sheet_file)
            # OLE objects referenced directly from the worksheet (no drawing anchor).
            for rel in sheet_rels:
                if rel["type"].endswith("/oleObject"):
                    _record_attachment(
                        zf,
                        rel["target"],
                        sheet_name,
                        None,
                        embeddings_dir,
                        collection,
                        written,
                        content_types,
                    )

            for drawing_file in _drawings:
                _extract_drawing(
                    zf,
                    drawing_file,
                    sheet_name,
                    media_dir,
                    charts_dir,
                    embeddings_dir,
                    collection,
                    written,
                    content_types,
                )

        _extract_equations(zf, sheet_by_drawing, collection)

    # Semantic enrichment (OCR + VLM caption) is skipped by default.
    # Set ``EDP_PADDLEOCR_*`` and ``EDP_VLM_*`` env vars to enable.
    for image in collection.images:
        _enrich_image_semantics(image, media_dir, resources_dir, collection.warnings)

    return collection


# --------------------------------------------------------------------------- #
# Relationship / part-path helpers
# --------------------------------------------------------------------------- #


def _rels_path_for(part: str) -> str:
    parent = PurePosixPath(part).parent.as_posix()
    name = PurePosixPath(part).name
    prefix = f"{parent}/" if parent and parent != "." else ""
    return f"{prefix}_rels/{name}.rels"


def _resolve_target(target: str, base_dir: str) -> str:
    """Resolve a relationship Target to a canonical zip entry path."""

    target_no_fragment = target.split("#", 1)[0]
    if target_no_fragment.startswith("/"):
        path = PurePosixPath(target_no_fragment.lstrip("/"))
    else:
        path = PurePosixPath(base_dir) / PurePosixPath(target_no_fragment)

    normalized: list[str] = []
    for part in path.parts:
        if part in {"", "."}:
            continue
        if part == "..":
            if normalized:
                normalized.pop()
            continue
        normalized.append(part)
    return "/".join(normalized)


def _read_rels(zf: zipfile.ZipFile, part: str) -> list[dict[str, str]]:
    """Read the relationship part for *part* → list of {id, type, target}."""

    rels_path = _rels_path_for(part)
    try:
        text = zf.read(rels_path).decode("utf-8", errors="ignore")
    except KeyError:
        return []
    root = ET.fromstring(text)
    base_dir = PurePosixPath(part).parent.as_posix()
    rels: list[dict[str, str]] = []
    for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rels.append(
            {
                "id": rel.attrib.get("Id", ""),
                "type": rel.attrib.get("Type", ""),
                "target": _resolve_target(rel.attrib.get("Target", ""), base_dir),
            }
        )
    return rels


def _sheet_files(
    zf: zipfile.ZipFile, warnings: list[str]
) -> list[tuple[str, str, list[str]]]:
    """Return ``(sheet_name, sheet_xml_path, [drawing_paths])`` in workbook order."""

    try:
        root = ET.fromstring(zf.read("xl/workbook.xml"))
    except KeyError:
        return []
    sheet_elements = root.find(f"{{{SPREADSHEETML_NS}}}sheets")
    if sheet_elements is None:
        return []

    workbook_rels = {
        rel["id"]: rel["target"] for rel in _read_rels(zf, "xl/workbook.xml")
    }
    result: list[tuple[str, str, list[str]]] = []
    for sheet_el in sheet_elements.findall(f"{{{SPREADSHEETML_NS}}}sheet"):
        name = sheet_el.attrib.get("name", "")
        rid = sheet_el.attrib.get(REL_ID_ATTR, "")
        target = workbook_rels.get(rid)
        if not target:
            warnings.append(
                f"Workbook sheet '{name}' has no worksheet relationship"
            )
            continue
        sheet_file = target
        drawings: list[str] = []
        for rel in _read_rels(zf, sheet_file):
            if rel["type"].endswith("/drawing"):
                drawings.append(rel["target"])
        result.append((name, sheet_file, drawings))
    return result


# --------------------------------------------------------------------------- #
# Drawing parsing
# --------------------------------------------------------------------------- #


def _extract_drawing(
    zf: zipfile.ZipFile,
    drawing_file: str,
    sheet_name: str,
    media_dir: Path,
    charts_dir: Path,
    embeddings_dir: Path,
    collection: XlsxAssetCollection,
    written: dict[str, Path],
    content_types: dict[str, str],
) -> None:
    try:
        root = ET.fromstring(zf.read(drawing_file))
    except KeyError:
        return

    rels = {rel["id"]: rel["target"] for rel in _read_rels(zf, drawing_file)}

    for anchor in root:
        if anchor.tag not in (
            f"{{{SPREADSHEETDRAW_NS}}}oneCellAnchor",
            f"{{{SPREADSHEETDRAW_NS}}}twoCellAnchor",
        ):
            continue
        anchor_meta = _anchor_metadata(anchor)

        graphic_frame = anchor.find(f"{{{SPREADSHEETDRAW_NS}}}graphicFrame")
        if graphic_frame is not None:
            chart_el = graphic_frame.find(f".//{{{CHART_NS}}}chart")
            if chart_el is not None:
                rid = chart_el.get(REL_ID_ATTR)
                if rid and rid in rels:
                    _record_chart(
                        zf,
                        rels[rid],
                        sheet_name,
                        anchor_meta,
                        charts_dir,
                        collection,
                        written,
                        content_types,
                    )
                continue
            # graphicFrame may also wrap an OLE object.
            for node in graphic_frame.iter():
                rid = node.get(REL_ID_ATTR)
                if rid and rid in rels and _looks_like_embedding(rels[rid]):
                    _record_attachment(
                        zf,
                        rels[rid],
                        sheet_name,
                        anchor_meta,
                        embeddings_dir,
                        collection,
                        written,
                        content_types,
                    )
                    break
            continue

        pic = anchor.find(f"{{{SPREADSHEETDRAW_NS}}}pic")
        if pic is not None:
            blip = pic.find(f".//{{{MAIN_NS}}}blip")
            if blip is not None:
                rid = blip.get(REL_EMBED_ATTR)
                if rid and rid in rels:
                    _record_image(
                        zf,
                        rels[rid],
                        sheet_name,
                        anchor_meta,
                        media_dir,
                        collection,
                        written,
                        content_types,
                    )


def _anchor_metadata(anchor: ET.Element) -> dict[str, object | None]:
    from_cell = _cell_from(anchor, "from")
    to_cell = _cell_from(anchor, "to")
    return {
        "cell": from_cell["cell"],
        "from_col": from_cell["col"],
        "from_row": from_cell["row"],
        "to_col": to_cell["col"],
        "to_row": to_cell["row"],
    }


def _cell_from(anchor: ET.Element, name: str) -> dict[str, object | None]:
    corner = anchor.find(f"{{{SPREADSHEETDRAW_NS}}}{name}")
    if corner is None:
        return {"cell": None, "col": None, "row": None}
    col_el = corner.find(f"{{{SPREADSHEETDRAW_NS}}}col")
    row_el = corner.find(f"{{{SPREADSHEETDRAW_NS}}}row")
    if col_el is None or row_el is None or not col_el.text or not row_el.text:
        return {"cell": None, "col": None, "row": None}
    col = int(col_el.text)
    row = int(row_el.text)
    cell = f"{get_column_letter(col + 1)}{row + 1}"
    return {"cell": cell, "col": col, "row": row}


def _looks_like_embedding(target: str) -> bool:
    return target.startswith("xl/embeddings/")


# --------------------------------------------------------------------------- #
# Per-kind recorders
# --------------------------------------------------------------------------- #


def _record_image(
    zf: zipfile.ZipFile,
    source_path: str,
    sheet_name: str,
    anchor: dict[str, object | None],
    media_dir: Path,
    collection: XlsxAssetCollection,
    written: dict[str, Path],
    content_types: dict[str, str],
) -> None:
    try:
        payload = zf.read(source_path)
    except KeyError:
        collection.warnings.append(f"Image part missing from package: {source_path}")
        return
    sha = _sha256(payload)
    if sha in written:
        return
    filename = PurePosixPath(source_path).name
    content_type = content_types.get(source_path)
    mime = _detect_mime(filename, payload, content_type)
    suffix = _suffix_or_default(filename, content_type, ".bin")
    ref = f"image_{len(collection.images) + 1:02d}"
    path = _dedup_write(written, sha, media_dir, ref, suffix, payload)
    collection.images.append(
        {
            "ref": ref,
            "kind": "image",
            "sheet": sheet_name,
            **anchor,
            "source_path": source_path,
            "path": _relative_to(path, media_dir.parent),
            "filename": filename,
            "mime": mime,
            "sha256": sha,
            "size_bytes": len(payload),
        }
    )


def _record_chart(
    zf: zipfile.ZipFile,
    source_path: str,
    sheet_name: str,
    anchor: dict[str, object | None],
    charts_dir: Path,
    collection: XlsxAssetCollection,
    written: dict[str, Path],
    content_types: dict[str, str],
) -> None:
    try:
        payload = zf.read(source_path)
    except KeyError:
        collection.warnings.append(
            f"Chart part missing from package: {source_path}"
        )
        return
    sha = _sha256(payload)
    if sha in written:
        return
    ref = f"chart_{len(collection.charts) + 1:02d}"
    xml_path = _dedup_write(written, sha, charts_dir, ref, ".xml", payload)
    chart_data = _read_chart_data(payload)
    json_path = charts_dir / f"{ref}.json"
    json_path.write_text(
        json.dumps(chart_data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    filename = PurePosixPath(source_path).name
    collection.charts.append(
        {
            "ref": ref,
            "kind": "chart",
            "sheet": sheet_name,
            **anchor,
            "source_path": source_path,
            "path": _relative_to(xml_path, charts_dir.parent),
            "json_path": _relative_to(json_path, charts_dir.parent),
            "filename": filename,
            "mime": content_types.get(source_path, "application/xml"),
            "sha256": sha,
            "size_bytes": len(payload),
            "title": chart_data.get("title"),
            "series": chart_data.get("series", []),
        }
    )


def _record_attachment(
    zf: zipfile.ZipFile,
    source_path: str,
    sheet_name: str,
    anchor: dict[str, object | None] | None,
    embeddings_dir: Path,
    collection: XlsxAssetCollection,
    written: dict[str, Path],
    content_types: dict[str, str],
) -> None:
    try:
        payload = zf.read(source_path)
    except KeyError:
        collection.warnings.append(
            f"Embedded object part missing: {source_path}"
        )
        return
    ole_class = _ole_class_text(payload)
    is_equation = ole_class is not None and any(
        marker in ole_class.lower() for marker in EQUATION_CLS_MARKERS
    )
    filename = _safe_filename(PurePosixPath(source_path).name)
    resolved_name, resolved_payload, unwrap_warnings = _normalize_attachment_payload(
        filename, payload
    )
    collection.warnings.extend(unwrap_warnings)

    sha = _sha256(resolved_payload)
    if sha in written:
        return
    content_type = content_types.get(source_path)
    mime = _detect_mime(resolved_name, resolved_payload, content_type)
    suffix = _suffix_or_default(resolved_name, content_type, ".bin")
    object_type = (
        "equation" if is_equation else suffix.removeprefix(".") or "attachment"
    )
    ref = f"attachment_{len(collection.attachments) + 1:02d}"
    path = _dedup_write(written, sha, embeddings_dir, ref, suffix, resolved_payload)
    if is_equation:
        collection.warnings.append(
            f"OLE equation object {ref} preserved as binary (not text-extractable)"
        )
    record = {
        "ref": ref,
        "kind": "attachment",
        "sheet": sheet_name,
        "source_path": source_path,
        "path": _relative_to(path, embeddings_dir.parent),
        "filename": resolved_name,
        "type": object_type,
        "mime": mime,
        "sha256": sha,
        "size_bytes": len(resolved_payload),
        "ole_class": ole_class,
    }
    if anchor is not None:
        record.update(anchor)
    collection.attachments.append(record)


# --------------------------------------------------------------------------- #
# Image semantic enrichment (OCR + VLM caption)
# --------------------------------------------------------------------------- #


def _enrich_image_semantics(
    image: dict,
    media_dir: Path,
    resources_dir: Path,
    warnings: list[str],
) -> None:
    """Run OCR + VLM caption on an extracted image and write a description sidecar.

    When ``EDP_PADDLEOCR_*`` / ``EDP_VLM_*`` env vars are unset this is a no-op
    that records ``not_configured``. The sidecar lands under
    ``assets/resources/{ref}/{description.md,description.json}``.
    """

    image_path = media_dir.parent / image["path"]

    if _should_skip_as_logo(image_path, image.get("mime")):
        image["ocr_text"] = ""
        image["caption"] = ""
        image["ocr_status"] = "logo"
        image["caption_status"] = "logo"
        image["semantic_status"] = "logo"
        return

    context = _image_context(image)
    ocr_text, ocr_status, ocr_warning = _run_image_ocr(image_path, image["ref"])
    caption, caption_status, caption_warning = _run_image_vlm_caption(
        image_path, image["ref"], image.get("mime"), context
    )
    semantic_status = _combined_semantic_status(ocr_status, caption_status)

    description_dir = resources_dir / image["ref"]
    description_dir.mkdir(parents=True, exist_ok=True)
    description_md = description_dir / "description.md"
    description_json = description_dir / "description.json"

    payload = {
        "ref": image["ref"],
        "filename": image.get("filename"),
        "mime": image.get("mime"),
        "image_path": image["path"],
        "source_path": image.get("source_path"),
        **context,
        "ocr_text": ocr_text,
        "caption": caption,
        "semantic_status": semantic_status,
        "ocr_status": ocr_status,
        "caption_status": caption_status,
    }
    description_json.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    description_md.write_text(
        _render_image_description(image, payload), encoding="utf-8"
    )

    image["ocr_text"] = ocr_text
    image["caption"] = caption
    image["ocr_status"] = ocr_status
    image["caption_status"] = caption_status
    image["semantic_status"] = semantic_status
    image["description_path"] = _relative_to(
        description_md, resources_dir.parent
    )
    image["description_json_path"] = _relative_to(
        description_json, resources_dir.parent
    )

    for warning in (ocr_warning, caption_warning):
        if warning is not None:
            warnings.append(warning)


def _image_context(image: dict) -> dict[str, object]:
    """Build a VLM prompt-context dict from an XLSX image's sheet/cell anchor."""

    location = f"XLSX sheet '{image.get('sheet')}'"
    if image.get("cell"):
        location += f", anchored at cell {image['cell']}"
    return {"context_before": location}


def _render_image_description(
    image: dict, payload: dict[str, object]
) -> str:
    caption = str(payload.get("caption") or "_Not generated._")
    ocr_text = str(payload.get("ocr_text") or "_Not generated._")
    use_chinese = _is_xlsx_context_chinese(payload)
    _source_file = "Source file" if not use_chinese else "源文件"
    _image_path = "Image path" if not use_chinese else "图片路径"
    _source_xlsx = "Source XLSX path" if not use_chinese else "源XLSX路径"
    _location = "Location" if not use_chinese else "位置"
    _semantic = "Semantic status" if not use_chinese else "语义状态"
    _caption_heading = "Caption" if not use_chinese else "图片描述"
    _ocr_heading = "OCR Text" if not use_chinese else "OCR文字"
    _not_generated = "_Not generated._" if not use_chinese else "未生成"
    if caption == "_Not generated._":
        caption = _not_generated
    if ocr_text == "_Not generated._":
        ocr_text = _not_generated
    return "\n".join(
        [
            f"# Image Description {image['ref']}",
            "",
            f"{_source_file}: `{image.get('filename')}`",
            f"{_image_path}: `{payload['image_path']}`",
            f"{_source_xlsx}: `{image.get('source_path')}`",
            f"{_location}: `{payload.get('context_before')}`",
            f"{_semantic}: `{payload['semantic_status']}`",
            "",
            f"## {_caption_heading}",
            "",
            caption,
            "",
            f"## {_ocr_heading}",
            "",
            ocr_text,
            "",
        ]
    )


def _is_xlsx_context_chinese(payload: dict[str, object]) -> bool:
    """Return True when the XLSX image context text is predominantly Chinese."""

    joined = " ".join(
        str(payload.get(key, ""))
        for key in ("caption", "ocr_text", "context_before")
    )
    chinese = sum(1 for ch in joined if "\u4e00" <= ch <= "\u9fff")
    alpha = sum(1 for ch in joined if ch.isalpha())
    return chinese > 0 and (alpha == 0 or chinese / max(alpha, 1) >= 0.3)


# --------------------------------------------------------------------------- #
# Chart XML text extraction
# --------------------------------------------------------------------------- #


def _read_chart_data(payload: bytes) -> dict[str, object]:
    try:
        root = ET.fromstring(payload)
    except ET.ParseError:
        return {"title": None, "series": []}

    title = _first_text(root, f".//{{{CHART_NS}}}title")
    series: list[dict[str, object]] = []
    for ser in root.findall(f".//{{{CHART_NS}}}ser"):
        series.append(
            {
                "name": _first_formula(ser, f".//{{{CHART_NS}}}tx"),
                "categories": _first_formula(ser, f".//{{{CHART_NS}}}cat"),
                "values": _first_formula(ser, f".//{{{CHART_NS}}}val"),
            }
        )
    return {"title": title, "series": series}


def _first_text(root: ET.Element, path: str) -> str | None:
    node = root.find(path)
    if node is None:
        return None
    texts = [t.text for t in node.iter(f"{{{MAIN_NS}}}t") if t.text]
    return "".join(texts) if texts else None


def _first_formula(root: ET.Element, path: str) -> str | None:
    node = root.find(path)
    if node is None:
        return None
    formula = node.find(f".//{{{CHART_NS}}}f")
    return formula.text if formula is not None and formula.text else None


# --------------------------------------------------------------------------- #
# Equations (OMML + OLE equation objects)
# --------------------------------------------------------------------------- #


def _extract_equations(
    zf: zipfile.ZipFile,
    sheet_by_drawing: dict[str, str],
    collection: XlsxAssetCollection,
) -> None:
    """Best-effort OMML equation extraction from every XML part in the package."""

    for name in zf.namelist():
        if not name.endswith(".xml") or not name.startswith("xl/"):
            continue
        try:
            root = ET.fromstring(zf.read(name))
        except (KeyError, ET.ParseError):
            continue
        omath_blocks = root.findall(
            f".//{{{MATH_NS}}}oMathPara"
        ) or root.findall(f".//{{{MATH_NS}}}oMath")
        if not omath_blocks:
            continue
        sheet_name = _owner_sheet(name, sheet_by_drawing)
        for block in omath_blocks:
            text = "".join(
                t.text for t in block.iter(f"{{{MATH_NS}}}t") if t.text
            ).strip()
            if not text:
                continue
            collection.equations.append(
                {
                    "ref": f"equation_{len(collection.equations) + 1:02d}",
                    "kind": "equation",
                    "sheet": sheet_name,
                    "cell": None,
                    "text": text,
                    "source": "omml",
                    "source_path": name,
                }
            )


def _owner_sheet(
    part: str, sheet_by_drawing: dict[str, str]
) -> str | None:
    if part.startswith("xl/drawings/"):
        return sheet_by_drawing.get(part)
    if part.startswith("xl/worksheets/"):
        return None
    return None


def _ole_class_text(payload: bytes) -> str | None:
    """Return the decoded ``\\x01CompObj`` stream text of an OLE container."""

    try:
        ole = olefile.OleFileIO(BytesIO(payload))
    except OSError:
        return None
    try:
        for stream in ole.listdir(streams=True, storages=False):
            if stream and stream[-1] == "\x01CompObj":
                return (
                    ole.openstream(stream)
                    .read()
                    .decode("latin-1", errors="ignore")
                )
        return None
    finally:
        ole.close()


# --------------------------------------------------------------------------- #
# Stubbed semantic enrichment (OCR + VLM)
# --------------------------------------------------------------------------- #


def _run_image_ocr(image_path: Path, ref: str) -> tuple[str, str, str | None]:
    """Stub that returns ``not_configured`` unless env vars are set.

    Set ``EDP_PADDLEOCR_AUTHORIZATION`` and ``EDP_PADDLEOCR_URL`` to enable.
    """

    import os

    authorization = os.environ.get("EDP_PADDLEOCR_AUTHORIZATION", "").strip()
    if not authorization:
        return "", "not_configured", None
    url = os.environ.get("EDP_PADDLEOCR_URL", "").strip()
    if not url:
        return "", "not_configured", None
    # Full implementation requires HTTP multipart; stubbed for now.
    return "", "not_configured", None


def _run_image_vlm_caption(
    image_path: Path,
    ref: str,
    content_type: str | None,
    context: dict[str, object],
) -> tuple[str, str, str | None]:
    """Stub that returns ``not_configured`` unless env vars are set.

    Set ``EDP_VLM_AUTHORIZATION`` and ``EDP_VLM_URL`` to enable.
    """

    import os

    authorization = os.environ.get("EDP_VLM_AUTHORIZATION", "").strip()
    if not authorization:
        return "", "not_configured", None
    url = os.environ.get("EDP_VLM_URL", "").strip()
    if not url:
        return "", "not_configured", None
    # Full implementation requires HTTP chat-completions; stubbed for now.
    return "", "not_configured", None


def _should_skip_as_logo(
    image_path: Path, content_type: str | None
) -> bool:
    """Return True when the image is small enough to be a likely logo.

    Uses a multi-signal heuristic:

    1. **Filename check**: common logo names (``logo*``, ``brand*``, ``icon*``,
       ``favicon*``, ``symbol*``) → treated as logo regardless of size.
    2. **Dimensions check**: area < ``EDP_LOGO_MAX_AREA`` (default 40 000 px²).

    Set ``EDP_LOGO_MAX_AREA`` to 0 to disable dimension-based detection.
    """

    import os

    # 1. Filename-based: common logo/icon/brand filenames.
    filename_lower = image_path.name.lower()
    logo_patterns = (
        "logo", "brand", "icon", "favicon", "symbol", "mark",
        "badge", "emblem", "crest", "insignia",
    )
    for pattern in logo_patterns:
        if filename_lower.startswith(pattern) or f"-{pattern}" in filename_lower:
            return True

    # 2. Dimension-based: very small images are likely logo/icon decorations.
    max_area_str = os.environ.get("EDP_LOGO_MAX_AREA", "40000").strip()
    try:
        max_area = int(max_area_str)
    except ValueError:
        return False
    if max_area <= 0:
        return False
    try:
        from PIL import Image

        with Image.open(image_path) as img:
            width, height = img.size
            # Also skip images that are square and small (logos are often square).
            aspect_ratio = max(width, height) / max(min(width, height), 1)
            if (width * height) < max_area:
                return True
            # Ultra-small narrow images (separators, decorative lines).
            if width < 50 or height < 50:
                return True
    except Exception:
        return False
    return False


def _combined_semantic_status(ocr_status: str, caption_status: str) -> str:
    if "logo" in {ocr_status, caption_status}:
        return "logo"
    if "generated" in {ocr_status, caption_status}:
        return "generated"
    if "failed" in {ocr_status, caption_status}:
        return "partial_failed"
    if {ocr_status, caption_status} == {"not_configured"}:
        return "not_configured"
    return "not_generated"


# --------------------------------------------------------------------------- #
# Shared file-writing helpers
# --------------------------------------------------------------------------- #


def _sha256(payload: bytes) -> str:
    import hashlib

    return hashlib.sha256(payload).hexdigest()


def _dedup_write(
    written: dict[str, Path],
    sha: str,
    directory: Path,
    ref: str,
    suffix: str,
    payload: bytes,
) -> Path:
    if sha in written:
        return written[sha]
    path = directory / f"{ref}{suffix}"
    path.write_bytes(payload)
    written[sha] = path
    return path


def _relative_to(path: Path, base: Path) -> str:
    try:
        return path.resolve().relative_to(base.resolve()).as_posix()
    except ValueError:
        return path.as_posix()
