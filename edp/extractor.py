from __future__ import annotations

import hashlib
from io import BytesIO
import mimetypes
from pathlib import Path, PurePosixPath
import re
import shutil
import struct
import zipfile
from xml.etree import ElementTree as ET

import olefile

from edp.models import EmbeddedObject, ExtractionResult


REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_ID_ATTR = f"{{{REL_NS}}}id"
REL_EMBED_ATTR = f"{{{REL_NS}}}embed"
REL_LINK_ATTR = f"{{{REL_NS}}}link"
REL_DIAGRAM_DATA_ATTR = f"{{{REL_NS}}}dm"
REL_DIAGRAM_LAYOUT_ATTR = f"{{{REL_NS}}}lo"
REL_DIAGRAM_QUICK_STYLE_ATTR = f"{{{REL_NS}}}qs"
REL_DIAGRAM_COLOR_STYLE_ATTR = f"{{{REL_NS}}}cs"
REL_ATTRS = (
    REL_ID_ATTR,
    REL_EMBED_ATTR,
    REL_LINK_ATTR,
    REL_DIAGRAM_DATA_ATTR,
    REL_DIAGRAM_LAYOUT_ATTR,
    REL_DIAGRAM_QUICK_STYLE_ATTR,
    REL_DIAGRAM_COLOR_STYLE_ATTR,
)
SAFE_FILENAME_RE = re.compile(r"[^A-Za-z0-9_. -]+")
RAGFLOW_EMBED_DIRS = (
    "word/embeddings/",
    "word/objects/",
    "word/activex/",
    "xl/embeddings/",
    "ppt/embeddings/",
)
RECOVERABLE_FILENAME_RE = re.compile(
    rb"([A-Za-z0-9][A-Za-z0-9_. -]{0,180}\."
    rb"(?:xlsx|xlsm|xls|docx|docm|pptx|pptm|pdf|png|jpe?g|csv|txt|log|zip))",
    re.IGNORECASE,
)
OOXML_EXTENSIONS = {
    "xl/workbook.xml": ".xlsx",
    "word/document.xml": ".docx",
    "ppt/presentation.xml": ".pptx",
}


def extract_document_assets(
    input_docx: str | Path, work_dir: str | Path, *, unsafe_unwrap: bool = False
) -> ExtractionResult:
    """Extract DOCX embedded attachments and images into a stable work directory."""

    input_path = Path(input_docx)
    output_dir = Path(work_dir)
    raw_dir = output_dir / "raw"
    embedded_dir = raw_dir / "embedded"
    media_dir = raw_dir / "media"
    charts_dir = raw_dir / "charts"
    diagrams_dir = raw_dir / "diagrams"
    raw_dir.mkdir(parents=True, exist_ok=True)
    embedded_dir.mkdir(parents=True, exist_ok=True)
    media_dir.mkdir(parents=True, exist_ok=True)
    charts_dir.mkdir(parents=True, exist_ok=True)
    diagrams_dir.mkdir(parents=True, exist_ok=True)

    raw_original = raw_dir / "original.docx"
    if input_path.resolve() != raw_original.resolve():
        shutil.copy2(input_path, raw_original)

    objects: list[EmbeddedObject] = []
    images: list[EmbeddedObject] = []
    charts: list[EmbeddedObject] = []
    diagrams: list[EmbeddedObject] = []
    warnings: list[str] = []

    try:
        with zipfile.ZipFile(input_path) as docx:
            relationships = _read_asset_relationships(docx, warnings)
            preview_image_rel_ids = _read_ole_preview_image_rel_ids(docx)
            relationships = _without_ole_preview_images(relationships, preview_image_rel_ids)
            relationships = _add_ragflow_directory_assets(docx, relationships)
            relationship_order = _read_document_relationship_order(docx, warnings)
            anchors = _read_document_relationship_anchors(docx, warnings)
            content_types = _read_content_types(docx)
            seen_payloads: set[str] = set()

            for rel_id in _ordered_asset_relationships(relationships, relationship_order):
                relationship = relationships[rel_id]
                source_path = relationship["source_path"]
                filename = PurePosixPath(source_path).name
                try:
                    payload = docx.read(source_path)
                except KeyError:
                    warnings.append(f"Embedded relationship {rel_id} target is missing: {source_path}")
                    continue
                payload_hash = hashlib.sha256(payload).hexdigest()
                if rel_id.startswith("scan:") and payload_hash in seen_payloads:
                    continue
                seen_payloads.add(payload_hash)

                content_type = (
                    content_types.get(source_path)
                    or mimetypes.guess_type(filename)[0]
                    or "application/octet-stream"
                )
                position = len(objects) + len(images) + len(charts) + len(diagrams) + 1
                if relationship["kind"] == "chart":
                    chart_index = len(charts) + 1
                    suffix = _suffix_or_default(filename, content_type, ".xml")
                    ref = f"chart_{chart_index:02d}"
                    asset_path = charts_dir / f"{ref}{suffix}"
                    asset_path.write_bytes(payload)
                    detected_mime = _detect_mime(filename, payload, content_type)
                    charts.append(
                        EmbeddedObject(
                            ref=ref,
                            filename=filename,
                            type="chart",
                            path=asset_path,
                            rel_id=rel_id,
                            source_path=source_path,
                            position=position,
                            content_type=content_type,
                            kind="chart",
                            resource_id=ref,
                            relationship_type=relationship.get("relationship_type"),
                            original_filename=filename,
                            detected_mime=detected_mime,
                            extension=suffix,
                            size_bytes=len(payload),
                            sha256=hashlib.sha256(payload).hexdigest(),
                            anchor=anchors.get(rel_id, {}),
                        )
                    )
                    continue

                if relationship["kind"] == "diagram":
                    diagram_index = len(diagrams) + 1
                    suffix = _suffix_or_default(filename, content_type, ".xml")
                    ref = f"diagram_{diagram_index:02d}"
                    asset_path = diagrams_dir / f"{ref}{suffix}"
                    asset_path.write_bytes(payload)
                    detected_mime = _detect_mime(filename, payload, content_type)
                    related_parts = [
                        part
                        for part in relationship.get("related_parts", [])
                        if isinstance(part, str)
                    ]
                    diagrams.append(
                        EmbeddedObject(
                            ref=ref,
                            filename=filename,
                            type="diagram",
                            path=asset_path,
                            rel_id=rel_id,
                            source_path=source_path,
                            position=position,
                            content_type=content_type,
                            kind="diagram",
                            resource_id=ref,
                            relationship_type=relationship.get("relationship_type"),
                            original_filename=filename,
                            detected_mime=detected_mime,
                            extension=suffix,
                            size_bytes=len(payload),
                            sha256=hashlib.sha256(payload).hexdigest(),
                            anchor=anchors.get(rel_id, {}),
                            related_parts=related_parts,
                        )
                    )
                    continue

                if relationship["kind"] == "image":
                    image_index = len(images) + 1
                    suffix = _suffix_or_default(filename, content_type, ".bin")
                    ref = f"image_{image_index:02d}"
                    asset_path = media_dir / f"{ref}{suffix}"
                    asset_path.write_bytes(payload)
                    detected_mime = _detect_mime(filename, payload, content_type)
                    images.append(
                        EmbeddedObject(
                            ref=ref,
                            filename=filename,
                            type=suffix.removeprefix(".") or "image",
                            path=asset_path,
                            rel_id=rel_id,
                            source_path=source_path,
                            position=position,
                            content_type=content_type,
                            kind="image",
                            resource_id=ref,
                            relationship_type=relationship.get("relationship_type"),
                            original_filename=filename,
                            detected_mime=detected_mime,
                            extension=suffix,
                            size_bytes=len(payload),
                            sha256=hashlib.sha256(payload).hexdigest(),
                            anchor=anchors.get(rel_id, {}),
                        )
                    )
                    continue

                attachment_filename, attachment_payload, attachment_warnings = (
                    _normalize_attachment_payload(filename, payload, unsafe_unwrap=unsafe_unwrap)
                )
                warnings.extend(attachment_warnings)
                attachment_index = len(objects) + 1
                suffix = _suffix_or_default(attachment_filename, content_type, ".bin")
                ref = f"attachment_{attachment_index:02d}"
                asset_path = embedded_dir / f"{ref}{suffix}"
                asset_path.write_bytes(attachment_payload)
                detected_mime = _detect_mime(attachment_filename, attachment_payload, content_type)
                objects.append(
                    EmbeddedObject(
                        ref=ref,
                        filename=attachment_filename,
                        type=suffix.removeprefix(".") or "attachment",
                        path=asset_path,
                        rel_id=rel_id,
                        source_path=source_path,
                        position=position,
                        content_type=content_type,
                        kind="attachment",
                        resource_id=ref,
                        relationship_type=relationship.get("relationship_type"),
                        original_filename=attachment_filename,
                        detected_mime=detected_mime,
                        extension=suffix,
                        size_bytes=len(attachment_payload),
                        sha256=hashlib.sha256(attachment_payload).hexdigest(),
                        anchor=anchors.get(rel_id, {}),
                    )
                )
    except zipfile.BadZipFile as exc:
        warnings.append(f"Input is not a readable DOCX zip package: {exc}")

    return ExtractionResult(
        input_docx=input_path,
        work_dir=output_dir,
        raw_original_path=raw_original,
        embedded_dir=embedded_dir,
        media_dir=media_dir,
        objects=objects,
        images=images,
        charts=charts,
        diagrams=diagrams,
        warnings=warnings,
    )


def extract_embedded_xlsx(input_docx: str | Path, work_dir: str | Path) -> ExtractionResult:
    """Extract embedded XLSX files from a DOCX package into a stable work directory."""

    input_path = Path(input_docx)
    output_dir = Path(work_dir)
    raw_dir = output_dir / "raw"
    embedded_dir = raw_dir / "embedded"
    raw_dir.mkdir(parents=True, exist_ok=True)
    embedded_dir.mkdir(parents=True, exist_ok=True)

    raw_original = raw_dir / "original.docx"
    if input_path.resolve() != raw_original.resolve():
        shutil.copy2(input_path, raw_original)

    objects: list[EmbeddedObject] = []
    warnings: list[str] = []

    try:
        with zipfile.ZipFile(input_path) as docx:
            relationships = _read_embedding_relationships(docx, warnings)
            relationship_order = _read_document_relationship_order(docx, warnings)
            ordered_rel_ids = _ordered_embedding_relationships(relationships, relationship_order)

            for rel_id in ordered_rel_ids:
                target = relationships[rel_id]
                source_path = _docx_zip_path(target)
                filename = PurePosixPath(source_path).name
                suffix = PurePosixPath(filename).suffix.lower()

                try:
                    payload = docx.read(source_path)
                except KeyError:
                    warnings.append(f"Embedded relationship {rel_id} target is missing: {source_path}")
                    continue

                if suffix == ".xlsx":
                    if not _is_xlsx_payload(payload):
                        warnings.append(f"Unsupported or corrupt XLSX embedded object skipped: {filename}")
                        continue
                    xlsx_payload = payload
                elif suffix == ".bin":
                    xlsx_payload = _extract_xlsx_from_ole_payload(payload)
                    if xlsx_payload is None:
                        warnings.append(f"Unsupported or corrupt OLE embedded object skipped: {filename}")
                        continue
                else:
                    continue

                ref = f"attachment_{len(objects) + 1:02d}"
                embedded_path = embedded_dir / f"{ref}.xlsx"
                embedded_path.write_bytes(xlsx_payload)
                objects.append(
                    EmbeddedObject(
                        ref=ref,
                        filename=filename,
                        type="xlsx",
                        path=embedded_path,
                        rel_id=rel_id,
                        source_path=source_path,
                        position=len(objects) + 1,
                    )
                )
    except zipfile.BadZipFile as exc:
        warnings.append(f"Input is not a readable DOCX zip package: {exc}")

    return ExtractionResult(
        input_docx=input_path,
        work_dir=output_dir,
        raw_original_path=raw_original,
        embedded_dir=embedded_dir,
        objects=objects,
        warnings=warnings,
    )


def _read_embedding_relationships(
    docx: zipfile.ZipFile, warnings: list[str]
) -> dict[str, str]:
    try:
        rels_xml = docx.read("word/_rels/document.xml.rels")
    except KeyError:
        warnings.append("DOCX relationship file is missing: word/_rels/document.xml.rels")
        return {}

    try:
        root = ET.fromstring(rels_xml)
    except ET.ParseError as exc:
        warnings.append(f"DOCX relationship file is not valid XML: {exc}")
        return {}

    relationships: dict[str, str] = {}
    for rel in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        target_mode = rel.attrib.get("TargetMode")
        if not rel_id or not target or target_mode == "External":
            continue

        path = _docx_zip_path(target)
        suffix = PurePosixPath(path).suffix.lower()
        if path.startswith("word/embeddings/") and suffix in {".xlsx", ".bin"}:
            relationships[rel_id] = target

    return relationships


def _read_asset_relationships(
    docx: zipfile.ZipFile, warnings: list[str]
) -> dict[str, dict[str, str]]:
    try:
        rels_xml = docx.read("word/_rels/document.xml.rels")
    except KeyError:
        warnings.append("DOCX relationship file is missing: word/_rels/document.xml.rels")
        return {}

    try:
        root = ET.fromstring(rels_xml)
    except ET.ParseError as exc:
        warnings.append(f"DOCX relationship file is not valid XML: {exc}")
        return {}

    relationships: dict[str, dict[str, str]] = {}
    for rel in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        target_mode = rel.attrib.get("TargetMode")
        if not rel_id or not target:
            continue
        if target_mode == "External":
            warnings.append(f"External relationship recorded but not fetched: {target}")
            continue

        source_path = _docx_zip_path(target)
        if _is_attachment_part(source_path):
            relationships[rel_id] = {
                "kind": "attachment",
                "source_path": source_path,
                "relationship_type": rel.attrib.get("Type", ""),
            }
        elif source_path.startswith("word/media/"):
            relationships[rel_id] = {
                "kind": "image",
                "source_path": source_path,
                "relationship_type": rel.attrib.get("Type", ""),
            }
        elif source_path.startswith("word/charts/") and source_path.endswith(".xml"):
            relationships[rel_id] = {
                "kind": "chart",
                "source_path": source_path,
                "relationship_type": rel.attrib.get("Type", ""),
            }
        elif _is_diagram_data_relationship(source_path, rel.attrib.get("Type", "")):
            relationships[rel_id] = {
                "kind": "diagram",
                "source_path": source_path,
                "relationship_type": rel.attrib.get("Type", ""),
                "related_parts": _diagram_related_parts(source_path, docx),
            }

    return relationships


def _is_diagram_data_relationship(source_path: str, relationship_type: str) -> bool:
    return (
        source_path.startswith("word/diagrams/")
        and PurePosixPath(source_path).name.startswith("data")
        and source_path.endswith(".xml")
        and relationship_type.endswith("/diagramData")
    )


def _diagram_related_parts(source_path: str, docx: zipfile.ZipFile) -> list[str]:
    stem = PurePosixPath(source_path).stem.removeprefix("data")
    candidates = [
        f"word/diagrams/drawing{stem}.xml",
        f"word/diagrams/layout{stem}.xml",
        f"word/diagrams/quickStyle{stem}.xml",
        f"word/diagrams/colors{stem}.xml",
    ]
    names = set(docx.namelist())
    return [candidate for candidate in candidates if candidate in names]


def _is_attachment_part(source_path: str) -> bool:
    return source_path.startswith(RAGFLOW_EMBED_DIRS)


def _add_ragflow_directory_assets(
    docx: zipfile.ZipFile, relationships: dict[str, dict[str, str]]
) -> dict[str, dict[str, str]]:
    discovered = dict(relationships)
    known_paths = {relationship["source_path"] for relationship in relationships.values()}
    for name in sorted(docx.namelist()):
        if not name.lower().startswith(RAGFLOW_EMBED_DIRS):
            continue
        if name in known_paths:
            continue
        rel_id = f"scan:{name}"
        discovered[rel_id] = {
            "kind": "attachment",
            "source_path": name,
            "relationship_type": "directory-scan",
        }
        known_paths.add(name)
    return discovered


def _read_document_relationship_order(
    docx: zipfile.ZipFile, warnings: list[str]
) -> list[str]:
    try:
        document_xml = docx.read("word/document.xml")
    except KeyError:
        warnings.append("DOCX document body is missing: word/document.xml")
        return []

    try:
        root = ET.fromstring(document_xml)
    except ET.ParseError as exc:
        warnings.append(f"DOCX document body is not valid XML: {exc}")
        return []

    ordered: list[str] = []
    seen: set[str] = set()
    for element in root.iter():
        for rel_id in _relationship_ids(element):
            if rel_id not in seen:
                ordered.append(rel_id)
                seen.add(rel_id)
    return ordered


def _read_ole_preview_image_rel_ids(docx: zipfile.ZipFile) -> set[str]:
    try:
        document_xml = docx.read("word/document.xml")
    except KeyError:
        return set()

    try:
        root = ET.fromstring(document_xml)
    except ET.ParseError:
        return set()

    preview_rel_ids: set[str] = set()
    for run in root.iter(f"{{{WORD_NS}}}r"):
        if not any(_local_name(element.tag) == "OLEObject" for element in run.iter()):
            continue
        for element in run.iter():
            if _local_name(element.tag) not in {"blip", "imagedata", "binData"}:
                continue
            preview_rel_ids.update(_relationship_ids(element))
    return preview_rel_ids


def _without_ole_preview_images(
    relationships: dict[str, dict[str, str]], preview_image_rel_ids: set[str]
) -> dict[str, dict[str, str]]:
    if not preview_image_rel_ids:
        return relationships
    return {
        rel_id: relationship
        for rel_id, relationship in relationships.items()
        if not (relationship["kind"] == "image" and rel_id in preview_image_rel_ids)
    }


def _relationship_ids(element: ET.Element) -> list[str]:
    return [element.attrib[attr] for attr in REL_ATTRS if element.attrib.get(attr)]


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _read_document_relationship_anchors(
    docx: zipfile.ZipFile, warnings: list[str]
) -> dict[str, dict[str, object]]:
    try:
        document_xml = docx.read("word/document.xml")
    except KeyError:
        return {}

    try:
        root = ET.fromstring(document_xml)
    except ET.ParseError:
        return {}

    anchors: dict[str, dict[str, object]] = {}
    paragraphs = list(root.iter(f"{{{WORD_NS}}}p"))
    for index, paragraph in enumerate(paragraphs, start=1):
        paragraph_text = "".join(
            text.text or "" for text in paragraph.iter(f"{{{WORD_NS}}}t")
        ).strip()
        rel_ids: list[str] = []
        for element in paragraph.iter():
            rel_ids.extend(_relationship_ids(element))
        for rel_id in rel_ids:
            anchors.setdefault(
                rel_id,
                {
                    "block_id": f"p_{index:04d}",
                    "paragraph_text": paragraph_text,
                    "page": None,
                },
            )
    return anchors


def _ordered_embedding_relationships(
    relationships: dict[str, str], relationship_order: list[str]
) -> list[str]:
    ordered = [rel_id for rel_id in relationship_order if rel_id in relationships]
    remaining = sorted(
        (rel_id for rel_id in relationships if rel_id not in ordered),
        key=lambda rel_id: relationships[rel_id],
    )
    return ordered + remaining


def _ordered_asset_relationships(
    relationships: dict[str, dict[str, str]], relationship_order: list[str]
) -> list[str]:
    ordered = [rel_id for rel_id in relationship_order if rel_id in relationships]
    remaining = sorted(
        (rel_id for rel_id in relationships if rel_id not in ordered),
        key=lambda rel_id: relationships[rel_id]["source_path"],
    )
    return ordered + remaining


def _read_content_types(docx: zipfile.ZipFile) -> dict[str, str]:
    try:
        text = docx.read("[Content_Types].xml").decode("utf-8", errors="ignore")
    except KeyError:
        return {}

    content_types: dict[str, str] = {}
    for part, content_type in re.findall(
        r'<Override\s+[^>]*PartName="([^"]+)"[^>]*ContentType="([^"]+)"', text
    ):
        content_types[part.lstrip("/")] = content_type

    defaults = {
        extension.lower(): content_type
        for extension, content_type in re.findall(
            r'<Default\s+[^>]*Extension="([^"]+)"[^>]*ContentType="([^"]+)"', text
        )
    }
    for name in docx.namelist():
        suffix = PurePosixPath(name).suffix.lower().removeprefix(".")
        if name not in content_types and suffix in defaults:
            content_types[name] = defaults[suffix]
    return content_types


def _docx_zip_path(target: str) -> str:
    target_without_fragment = target.split("#", 1)[0]
    if target_without_fragment.startswith("/"):
        path = PurePosixPath(target_without_fragment.lstrip("/"))
    else:
        path = PurePosixPath("word") / PurePosixPath(target_without_fragment)

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


def _extract_xlsx_from_ole_payload(payload: bytes) -> bytes | None:
    if _is_xlsx_payload(payload):
        return payload

    try:
        ole = olefile.OleFileIO(BytesIO(payload))
    except OSError:
        return _find_zip_payload(payload)

    try:
        stream_names = ole.listdir(streams=True, storages=False)
        preferred_streams = [["Package"]]
        preferred_streams.extend(
            stream for stream in stream_names if stream and stream[-1] == "\x01Ole10Native"
        )
        preferred_streams.extend(stream for stream in stream_names if stream not in preferred_streams)

        for stream_name in preferred_streams:
            if not ole.exists(stream_name):
                continue
            stream_payload = ole.openstream(stream_name).read()
            xlsx_payload = _find_zip_payload(stream_payload)
            if xlsx_payload is not None:
                return xlsx_payload
    finally:
        ole.close()

    return _find_zip_payload(payload)


def _normalize_attachment_payload(
    filename: str, payload: bytes, *, unsafe_unwrap: bool = False
) -> tuple[str, bytes, list[str]]:
    suffix = PurePosixPath(filename).suffix.lower()
    if suffix != ".bin":
        return filename, payload, []

    ole_payload = _extract_ole10_native_payload(payload)
    if ole_payload is not None:
        payload_filename, payload_bytes = ole_payload
        return _safe_filename(payload_filename), payload_bytes, []

    if unsafe_unwrap:
        unwrapped = _unsafe_unwrap_attachment_payload(filename, payload)
        if unwrapped is not None:
            return unwrapped[0], unwrapped[1], []

    xlsx_payload = _extract_xlsx_from_ole_payload(payload)
    if xlsx_payload is not None:
        return f"{PurePosixPath(filename).stem}.xlsx", xlsx_payload, []

    return filename, payload, [f"Unknown OLE embedded object preserved as binary: {filename}"]


def _extract_ole10_native_payload(payload: bytes) -> tuple[str, bytes] | None:
    try:
        ole = olefile.OleFileIO(BytesIO(payload))
    except OSError:
        return None

    try:
        stream_names = ole.listdir(streams=True, storages=False)
        native_stream = next(
            (stream for stream in stream_names if stream and stream[-1] == "\x01Ole10Native"),
            None,
        )
        if native_stream is None or not ole.exists(native_stream):
            return None
        native = ole.openstream(native_stream).read()
    finally:
        ole.close()

    return _parse_ole10_native(native)


def _parse_ole10_native(data: bytes) -> tuple[str, bytes] | None:
    if len(data) < 8:
        return None
    pos = 4
    if len(data) >= 6 and data[4:6] in {b"\x01\x00", b"\x02\x00"}:
        pos = 6
    try:
        filename, pos = _read_c_string(data, pos)
        _, pos = _read_c_string(data, pos)
    except ValueError:
        return None
    if not filename:
        filename = "ole_payload.bin"

    for skip in (8, 4, 0):
        candidate_pos = pos + skip
        if candidate_pos >= len(data):
            continue
        try:
            _, size_pos = _read_c_string(data, candidate_pos)
        except ValueError:
            continue
        if size_pos + 4 > len(data):
            continue
        declared_size = struct.unpack_from("<I", data, size_pos)[0]
        payload_start = size_pos + 4
        payload_end = payload_start + declared_size
        if declared_size <= 0 or payload_end > len(data):
            continue
        return filename, data[payload_start:payload_end]
    return None


def _unsafe_unwrap_attachment_payload(filename: str, payload: bytes) -> tuple[str, bytes] | None:
    native_payload = _scan_ole10_native_payload(payload)
    if native_payload is not None:
        payload_filename, payload_bytes = native_payload
        return _safe_filename(payload_filename), payload_bytes

    candidates = [payload, *_ole_stream_payloads(payload)]
    recovered_filename = _recover_embedded_filename(candidates)
    for candidate in candidates:
        extracted = _extract_known_payload(candidate, include_text=False)
        if extracted is None:
            continue
        extension, payload_bytes = extracted
        payload_filename = recovered_filename or f"{PurePosixPath(filename).stem}{extension}"
        return _safe_filename(_filename_with_extension(payload_filename, extension)), payload_bytes

    for candidate in candidates:
        extracted = _extract_known_payload(candidate, include_text=True)
        if extracted is None:
            continue
        extension, payload_bytes = extracted
        payload_filename = recovered_filename or f"{PurePosixPath(filename).stem}{extension}"
        return _safe_filename(_filename_with_extension(payload_filename, extension)), payload_bytes

    return None


def _scan_ole10_native_payload(payload: bytes) -> tuple[str, bytes] | None:
    for match in RECOVERABLE_FILENAME_RE.finditer(payload):
        offsets = [match.start() - skip for skip in (4, 6)]
        for offset in offsets:
            if offset < 0:
                continue
            native_payload = _parse_ole10_native(payload[offset:])
            if native_payload is not None:
                return native_payload
    return None


def _ole_stream_payloads(payload: bytes) -> list[bytes]:
    try:
        ole = olefile.OleFileIO(BytesIO(payload))
    except OSError:
        return []

    try:
        stream_payloads = []
        for stream_name in ole.listdir(streams=True, storages=False):
            if ole.exists(stream_name):
                stream_payloads.append(ole.openstream(stream_name).read())
        return stream_payloads
    finally:
        ole.close()


def _recover_embedded_filename(candidates: list[bytes]) -> str | None:
    for candidate in candidates:
        match = RECOVERABLE_FILENAME_RE.search(candidate)
        if match is not None:
            return _safe_filename(match.group(1).decode("utf-8", errors="replace"))
    return None


def _extract_known_payload(payload: bytes, *, include_text: bool) -> tuple[str, bytes] | None:
    zip_payload = _find_zip_like_payload(payload)
    if zip_payload is not None:
        return zip_payload

    signature_payloads = (
        (".pdf", b"%PDF-"),
        (".png", b"\x89PNG\r\n\x1a\n"),
        (".jpg", b"\xff\xd8\xff"),
    )
    for extension, signature in signature_payloads:
        start = payload.find(signature)
        if start >= 0:
            return extension, payload[start:]

    if include_text and _looks_like_text_payload(payload):
        return ".txt", payload
    return None


def _find_zip_like_payload(payload: bytes) -> tuple[str, bytes] | None:
    for candidate in _zip_candidates(payload):
        try:
            with zipfile.ZipFile(BytesIO(candidate)) as package:
                if package.testzip() is not None:
                    continue
                names = set(package.namelist())
        except zipfile.BadZipFile:
            continue

        for marker, extension in OOXML_EXTENSIONS.items():
            if marker in names:
                return extension, candidate
        return ".zip", candidate
    return None


def _zip_candidates(payload: bytes) -> list[bytes]:
    candidates: list[bytes] = []
    start = 0
    while True:
        zip_start = payload.find(b"PK\x03\x04", start)
        if zip_start < 0:
            break
        candidates.append(payload[zip_start:])
        start = zip_start + 1
    if payload not in candidates:
        candidates.append(payload)
    return candidates


def _looks_like_text_payload(payload: bytes) -> bool:
    if not payload or b"\x00" in payload[:4096]:
        return False
    sample = payload[:4096]
    printable = sum(byte in b"\t\n\r" or 32 <= byte <= 126 for byte in sample)
    return printable / len(sample) > 0.85


def _filename_with_extension(filename: str, extension: str) -> str:
    current = PurePosixPath(filename).suffix.lower()
    if current == extension:
        return filename
    return f"{PurePosixPath(filename).stem}{extension}"


def _read_c_string(data: bytes, pos: int) -> tuple[str, int]:
    end = data.find(b"\x00", pos)
    if end < 0:
        raise ValueError("NUL-terminated string not found")
    raw = data[pos:end]
    return raw.decode("utf-8", errors="replace"), end + 1


def _safe_filename(filename: str) -> str:
    name = PurePosixPath(filename.replace("\\", "/")).name.strip() or "ole_payload.bin"
    return SAFE_FILENAME_RE.sub("_", name)


def _suffix_or_default(filename: str, content_type: str | None, default: str) -> str:
    suffix = PurePosixPath(filename).suffix.lower()
    if suffix:
        return suffix
    guessed = mimetypes.guess_extension(content_type or "")
    return guessed or default


def _detect_mime(filename: str, payload: bytes, content_type: str | None) -> str:
    if _is_xlsx_payload(payload):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if payload.startswith(b"%PDF-"):
        return "application/pdf"
    if payload.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if payload.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if payload.startswith(b"PK\x03\x04"):
        return "application/zip"

    guessed = mimetypes.guess_type(filename)[0]
    if guessed == "text/xml" and PurePosixPath(filename).suffix.lower() == ".xml":
        return "application/xml"
    if guessed:
        return guessed
    return content_type or "application/octet-stream"


def _find_zip_payload(payload: bytes) -> bytes | None:
    candidates = [payload]
    zip_start = payload.find(b"PK\x03\x04")
    if zip_start > 0:
        candidates.append(payload[zip_start:])

    for candidate in candidates:
        if _is_xlsx_payload(candidate):
            return candidate
    return None


def _is_xlsx_payload(payload: bytes) -> bool:
    try:
        with zipfile.ZipFile(BytesIO(payload)) as package:
            names = set(package.namelist())
            return (
                package.testzip() is None
                and "[Content_Types].xml" in names
                and "xl/workbook.xml" in names
            )
    except zipfile.BadZipFile:
        return False
