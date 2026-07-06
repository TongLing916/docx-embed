from __future__ import annotations

import csv
import json
import re
from pathlib import Path
import shutil
import tempfile
import zipfile
from xml.etree import ElementTree as ET

from edp.docx_notes import normalize_docx_notes_markdown
from edp.main_parser import parse_main_document
from edp.manifest_builder import build_manifest
from edp.models import DocumentPackage, EmbeddedObject, ExtractionResult, ParsedPackage
from edp.resource_preview import build_resource_previews


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
OOXML_NAMESPACES = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": REL_NS,
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w": WORD_NS,
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}
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
POSITION_MAP_COLUMNS = [
    "position",
    "ref",
    "kind",
    "filename",
    "type",
    "source_path",
    "markdown_reference",
    "target_path",
    "entry_point",
    "tables",
]


def merge_parent_with_attachments(
    parent_docx: str | Path,
    extraction: ExtractionResult,
    parsed_attachments: list[ParsedPackage],
    output_dir: str | Path,
    main_parser: str = "markitdown",
) -> DocumentPackage:
    """Build the final document package and insert placeholders for parsed attachments."""

    parent_path = Path(parent_docx)
    package_dir = Path(output_dir)
    raw_dir = package_dir / "raw"
    raw_embedded_dir = raw_dir / "embedded"
    raw_charts_dir = raw_dir / "charts"
    raw_diagrams_dir = raw_dir / "diagrams"
    structured_dir = package_dir / "structured"
    images_dir = structured_dir / "assets" / "images"
    attachments_dir = structured_dir / "attachments"
    meta_dir = structured_dir / "_meta"

    raw_embedded_dir.mkdir(parents=True, exist_ok=True)
    raw_charts_dir.mkdir(parents=True, exist_ok=True)
    raw_diagrams_dir.mkdir(parents=True, exist_ok=True)
    images_dir.mkdir(parents=True, exist_ok=True)
    structured_dir.mkdir(parents=True, exist_ok=True)
    meta_dir.mkdir(parents=True, exist_ok=True)

    extraction = build_resource_previews(extraction, package_dir)

    _copy_file(parent_path, raw_dir / "original.docx")
    for embedded in extraction.objects:
        _copy_file(embedded.path, raw_embedded_dir / embedded.path.name)
    for image in extraction.images:
        _copy_file(image.path, images_dir / image.path.name)
    for chart in extraction.charts:
        _copy_file(chart.path, raw_charts_dir / chart.path.name)
    for diagram in extraction.diagrams:
        _copy_file(diagram.path, raw_diagrams_dir / diagram.path.name)

    parsed_by_ref = {parsed.ref: parsed for parsed in parsed_attachments}
    if parsed_attachments:
        attachments_dir.mkdir(parents=True, exist_ok=True)
    for parsed in parsed_attachments:
        destination = attachments_dir / parsed.ref
        if parsed.package_dir.resolve() == destination.resolve():
            continue
        if destination.exists():
            shutil.rmtree(destination)
        shutil.copytree(parsed.package_dir, destination)

    warnings = [*extraction.warnings]
    for parsed in parsed_attachments:
        warnings.extend(parsed.warnings)

    content_path = structured_dir / "content.md"
    clean_main_path = meta_dir / "work" / "clean" / "main.docx"
    parser_artifact_dir = meta_dir / "parsers" / main_parser
    parser_artifact_dir.mkdir(parents=True, exist_ok=True)
    content, render_warnings = _render_parent_content(
        parent_path,
        extraction,
        parsed_by_ref,
        main_parser=main_parser,
        clean_main_path=clean_main_path,
        parser_artifact_dir=parser_artifact_dir,
    )
    warnings.extend(render_warnings)
    content_path.write_text(content, encoding="utf-8")
    asset_entries = _asset_entries(extraction, parsed_by_ref)
    child_files_path = structured_dir / "child_files.md"
    child_files_path.write_text(_render_child_files(asset_entries), encoding="utf-8")
    _write_position_map(structured_dir / "position_map.csv", asset_entries)
    embedded_resources_path = structured_dir / "embedded_resources.jsonl"
    _write_resource_index(embedded_resources_path, extraction)

    content_map = {
        "main": "structured/content.md",
        "clean_main": "structured/_meta/work/clean/main.docx",
        "main_parser": main_parser,
        "main_parser_artifacts": f"structured/_meta/parsers/{main_parser}/",
        "child_files": "structured/child_files.md",
        "position_map": "structured/position_map.csv",
        "embedded_resources": "structured/embedded_resources.jsonl",
        "embedded_objects": [
            _manifest_entry(embedded, parsed_by_ref.get(embedded.ref)) for embedded in extraction.objects
        ],
        "images": [_image_manifest_entry(image) for image in extraction.images],
        "charts": [_chart_manifest_entry(chart) for chart in extraction.charts],
        "diagrams": [_diagram_manifest_entry(diagram) for diagram in extraction.diagrams],
    }
    status = "partial" if warnings else "success"
    manifest = build_manifest(package_dir, status, warnings, content_map)

    manifest_path = package_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (meta_dir / "parse_log.txt").write_text(_render_parse_log(status, warnings), encoding="utf-8")

    return DocumentPackage(
        package_dir=package_dir,
        content_path=content_path,
        manifest_path=manifest_path,
        warnings=warnings,
    )


def _render_parent_content(
    parent_docx: Path,
    extraction: ExtractionResult,
    parsed_by_ref: dict[str, ParsedPackage],
    *,
    main_parser: str = "markitdown",
    clean_main_path: Path | None = None,
    parser_artifact_dir: Path | None = None,
) -> tuple[str, list[str]]:
    rel_to_asset = {
        **{embedded.rel_id: embedded for embedded in extraction.objects},
        **{image.rel_id: image for image in extraction.images},
        **{chart.rel_id: chart for chart in extraction.charts},
        **{diagram.rel_id: diagram for diagram in extraction.diagrams},
    }
    if extraction.media_dir is None:
        return _render_parent_content_from_xml(parent_docx, rel_to_asset, parsed_by_ref), []

    placeholders = _asset_placeholders(extraction, parsed_by_ref)
    warnings: list[str] = []

    try:
        if clean_main_path is None:
            with tempfile.TemporaryDirectory(prefix="edp_clean_main_") as temp_dir:
                clean_docx = Path(temp_dir) / parent_docx.name
                _write_docx_with_asset_sentinels(parent_docx, clean_docx, rel_to_asset)
                markdown = _parse_clean_main_docx(clean_docx, main_parser, parser_artifact_dir)
        else:
            clean_main_path.parent.mkdir(parents=True, exist_ok=True)
            _write_docx_with_asset_sentinels(parent_docx, clean_main_path, rel_to_asset)
            markdown = _parse_clean_main_docx(clean_main_path, main_parser, parser_artifact_dir)
        return _replace_asset_sentinels(markdown, placeholders), warnings
    except Exception as exc:
        parser_label = "MarkItDown" if main_parser == "markitdown" else main_parser
        warnings.append(f"{parser_label} failed to convert parent document: {exc}")
        return _render_parent_content_from_xml(parent_docx, rel_to_asset, parsed_by_ref), warnings


def _render_parent_content_from_xml(
    parent_docx: Path,
    rel_to_asset: dict[str, EmbeddedObject],
    parsed_by_ref: dict[str, ParsedPackage],
) -> str:

    lines: list[str] = []
    try:
        with zipfile.ZipFile(parent_docx) as docx:
            document_xml = docx.read("word/document.xml")
        root = ET.fromstring(document_xml)
    except (KeyError, ET.ParseError, zipfile.BadZipFile):
        root = None

    if root is not None:
        body = root.find(f".//{{{WORD_NS}}}body")
        blocks = list(body) if body is not None else list(root)
        for block in blocks:
            block_lines = _block_lines(block, rel_to_asset, parsed_by_ref)
            if block_lines:
                lines.extend(block_lines)
                lines.append("")

    if not lines:
        lines.extend(["# Document", "", "No extractable parent text was found.", ""])

    return "\n".join(lines).rstrip() + "\n"


def _block_lines(
    block: ET.Element,
    rel_to_asset: dict[str, EmbeddedObject],
    parsed_by_ref: dict[str, ParsedPackage],
) -> list[str]:
    lines: list[str] = []
    text_buffer: list[str] = []
    seen: set[str] = set()

    def flush_text() -> None:
        text = "".join(text_buffer).strip()
        text_buffer.clear()
        if text:
            lines.append(text)

    def walk(element: ET.Element) -> None:
        rel_id = _element_rel_id(element)
        asset = rel_to_asset.get(rel_id) if rel_id else None
        if asset is not None and rel_id not in seen:
            flush_text()
            if asset.kind == "image":
                lines.extend(_image_placeholder_lines(asset))
            else:
                lines.extend(_placeholder_lines(asset, parsed_by_ref.get(asset.ref)))
            seen.add(rel_id)
            return

        if element.tag == f"{{{WORD_NS}}}t":
            text_buffer.append(element.text or "")
            return

        for child in element:
            walk(child)

    walk(block)
    flush_text()
    return lines


def _write_docx_with_asset_sentinels(
    source_docx: Path, destination_docx: Path, rel_to_asset: dict[str, EmbeddedObject]
) -> None:
    with zipfile.ZipFile(source_docx) as source, zipfile.ZipFile(
        destination_docx, "w", zipfile.ZIP_DEFLATED
    ) as destination:
        for item in source.infolist():
            payload = source.read(item.filename)
            if item.filename == "word/document.xml":
                payload = _document_xml_with_asset_sentinels(payload, rel_to_asset)
            destination.writestr(item, payload)


def _document_xml_with_asset_sentinels(
    document_xml: bytes, rel_to_asset: dict[str, EmbeddedObject]
) -> bytes:
    for prefix, uri in OOXML_NAMESPACES.items():
        ET.register_namespace(prefix, uri)
    root = ET.fromstring(document_xml)
    parent_by_child = {child: parent for parent in root.iter() for child in parent}
    replaced_rel_ids: set[str] = set()
    for element in list(root.iter()):
        rel_id = _element_rel_id(element)
        if not rel_id or rel_id not in rel_to_asset or rel_id in replaced_rel_ids:
            continue
        run = _nearest_run(element, parent_by_child)
        if run is None:
            run = element
        run.clear()
        text = ET.SubElement(run, f"{{{WORD_NS}}}t")
        text.text = f"[[EDP_ASSET:{rel_to_asset[rel_id].ref}]]"
        replaced_rel_ids.add(rel_id)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _nearest_run(
    element: ET.Element, parent_by_child: dict[ET.Element, ET.Element]
) -> ET.Element | None:
    current: ET.Element | None = element
    while current is not None:
        if current.tag == f"{{{WORD_NS}}}r":
            return current
        current = parent_by_child.get(current)
    return None


def _element_rel_id(element: ET.Element) -> str | None:
    for attr in REL_ATTRS:
        rel_id = element.attrib.get(attr)
        if rel_id:
            return rel_id
    return None


def _convert_docx_to_markdown(input_path: Path) -> str:
    from markitdown import MarkItDown

    result = MarkItDown().convert(str(input_path))
    return str(result.text_content)


def _parse_clean_main_docx(
    input_path: Path, main_parser: str, parser_artifact_dir: Path | None
) -> str:
    if main_parser == "markitdown":
        markdown = normalize_docx_notes_markdown(input_path, _convert_docx_to_markdown(input_path))
        if parser_artifact_dir is not None:
            parser_artifact_dir.mkdir(parents=True, exist_ok=True)
            (parser_artifact_dir / f"{input_path.stem}.md").write_text(markdown, encoding="utf-8")
        return markdown
    result = parse_main_document(
        input_path,
        main_parser,
        parser_artifact_dir or input_path.parent / "__parser_artifacts__",
    )
    if result.markdown:
        return result.markdown
    if result.warnings:
        raise RuntimeError("; ".join(result.warnings))
    raise RuntimeError(f"{main_parser} did not return Markdown content")


# Matches an asset sentinel as written into the clean DOCX, tolerant of the
# Markdown escaping each parser applies: markitdown escapes underscores
# (``[[EDP\_ASSET:image\_01]]``), pandoc escapes brackets
# (``\[\[EDP_ASSET:image_01\]\]``). An optional backslash is allowed before
# each ``[``, ``]`` and the ``_`` in ``EDP_ASSET``; the captured ref is
# unescaped before lookup.
_SENTINEL_RE = re.compile(r"\\?\[\\?\[EDP\\?_ASSET:(?P<ref>[^[\]]+)\\?\]\\?\]")


def _replace_asset_sentinels(markdown: str, placeholders: dict[str, str]) -> str:
    found: set[str] = set()

    def _sub(match: re.Match[str]) -> str:
        ref = match.group("ref").replace("\\", "")
        placeholder = placeholders.get(ref)
        if placeholder is None:
            return match.group(0)
        found.add(ref)
        return placeholder

    rendered = _SENTINEL_RE.sub(_sub, markdown)
    missing = [ref for ref in placeholders if ref not in found]
    if missing:
        rendered = rendered.rstrip() + "\n\n" + "\n\n".join(placeholders[ref] for ref in missing)
    return rendered.rstrip() + "\n"


def _asset_placeholders(
    extraction: ExtractionResult, parsed_by_ref: dict[str, ParsedPackage]
) -> dict[str, str]:
    placeholders: dict[str, str] = {}
    for embedded in extraction.objects:
        placeholders[embedded.ref] = "\n".join(
            _placeholder_lines(embedded, parsed_by_ref.get(embedded.ref))
        )
    for image in extraction.images:
        placeholders[image.ref] = "\n".join(_image_placeholder_lines(image))
    for chart in extraction.charts:
        placeholders[chart.ref] = "\n".join(_chart_placeholder_lines(chart))
    for diagram in extraction.diagrams:
        placeholders[diagram.ref] = "\n".join(_diagram_placeholder_lines(diagram))
    return placeholders


def _placeholder_lines(embedded: EmbeddedObject, parsed: ParsedPackage | None) -> list[str]:
    if embedded.kind == "chart":
        return _chart_placeholder_lines(embedded)
    if embedded.kind == "diagram":
        return _diagram_placeholder_lines(embedded)
    return [_attachment_reference(embedded, parsed)]


def _image_placeholder_lines(image: EmbeddedObject) -> list[str]:
    return [_image_reference(image)]


def _chart_placeholder_lines(chart: EmbeddedObject) -> list[str]:
    return [_chart_reference(chart)]


def _diagram_placeholder_lines(diagram: EmbeddedObject) -> list[str]:
    return [_diagram_reference(diagram)]


def _asset_entries(
    extraction: ExtractionResult, parsed_by_ref: dict[str, ParsedPackage]
) -> list[dict[str, str]]:
    entries = [
        _attachment_entry(embedded, parsed_by_ref.get(embedded.ref))
        for embedded in extraction.objects
    ]
    entries.extend(_image_entry(image) for image in extraction.images)
    entries.extend(_chart_entry(chart) for chart in extraction.charts)
    entries.extend(_diagram_entry(diagram) for diagram in extraction.diagrams)
    return sorted(entries, key=lambda entry: (int(entry["position"]), entry["ref"]))


def _attachment_entry(
    embedded: EmbeddedObject, parsed: ParsedPackage | None
) -> dict[str, str]:
    tables = _table_paths(embedded.ref, parsed)
    preview_path = _preview_path(embedded)
    if parsed is not None:
        target_path = f"structured/attachments/{embedded.ref}/"
        entry_point = f"structured/attachments/{embedded.ref}/content.md"
        link_target = f"attachments/{embedded.ref}/content.md"
    elif preview_path:
        target_path = preview_path
        entry_point = preview_path
        link_target = preview_path.removeprefix("structured/")
    else:
        target_path = f"raw/embedded/{embedded.path.name}"
        entry_point = ""
        link_target = f"../raw/embedded/{embedded.path.name}"
    return {
        "position": str(embedded.position),
        "ref": embedded.ref,
        "kind": "attachment",
        "filename": embedded.filename,
        "type": embedded.type,
        "source_path": embedded.source_path,
        "markdown_reference": _attachment_reference(embedded, parsed),
        "target_path": target_path,
        "entry_point": entry_point,
        "tables": ";".join(tables),
        "link_target": link_target,
    }


def _image_entry(image: EmbeddedObject) -> dict[str, str]:
    target_path = f"structured/assets/images/{image.path.name}"
    return {
        "position": str(image.position),
        "ref": image.ref,
        "kind": "image",
        "filename": image.filename,
        "type": image.type,
        "source_path": image.source_path,
        "markdown_reference": _image_reference(image),
        "target_path": target_path,
        "entry_point": "",
        "tables": "",
        "link_target": f"../{target_path}",
    }


def _chart_entry(chart: EmbeddedObject) -> dict[str, str]:
    preview_path = _preview_path(chart) or f"structured/resources/{chart.ref}/preview.md"
    return {
        "position": str(chart.position),
        "ref": chart.ref,
        "kind": "chart",
        "filename": chart.filename,
        "type": chart.type,
        "source_path": chart.source_path,
        "markdown_reference": _chart_reference(chart),
        "target_path": preview_path,
        "entry_point": preview_path,
        "tables": "",
        "link_target": preview_path.removeprefix("structured/"),
    }


def _diagram_entry(diagram: EmbeddedObject) -> dict[str, str]:
    preview_path = _preview_path(diagram) or f"structured/resources/{diagram.ref}/preview.md"
    return {
        "position": str(diagram.position),
        "ref": diagram.ref,
        "kind": "diagram",
        "filename": diagram.filename,
        "type": diagram.type,
        "source_path": diagram.source_path,
        "markdown_reference": _diagram_reference(diagram),
        "target_path": preview_path,
        "entry_point": preview_path,
        "tables": "",
        "link_target": preview_path.removeprefix("structured/"),
    }


def _render_child_files(entries: list[dict[str, str]]) -> str:
    lines = [
        "# Child Files",
        "",
        "| ref | kind | filename | path | entry point | tables |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    if entries:
        for entry in entries:
            lines.append(
                "| "
                + " | ".join(
                    _markdown_table_cell(entry[column])
                    for column in ("ref", "kind", "filename", "target_path", "entry_point", "tables")
                )
                + " |"
            )
    else:
        lines.append("|  |  |  |  |  |  |")

    lines.extend(["", "## YAML Index", "", "```yaml", *_child_files_yaml_lines(entries), "```"])
    return "\n".join(lines).rstrip() + "\n"


def _child_files_yaml_lines(entries: list[dict[str, str]]) -> list[str]:
    if not entries:
        return ["child_files: []"]

    lines = ["child_files:"]
    for entry in entries:
        lines.extend(
            [
                f"  - ref: {_yaml_scalar(entry['ref'])}",
                f"    kind: {_yaml_scalar(entry['kind'])}",
                f"    filename: {_yaml_scalar(entry['filename'])}",
                f"    markdown_reference: {_yaml_scalar(entry['markdown_reference'])}",
                f"    path: {_yaml_scalar(entry['target_path'])}",
                f"    entry_point: {_yaml_scalar(entry['entry_point'])}",
            ]
        )
        tables = [table for table in entry["tables"].split(";") if table]
        if tables:
            lines.append("    tables:")
            lines.extend(f"      - {_yaml_scalar(table)}" for table in tables)
        else:
            lines.append("    tables: []")
    return lines


def _write_position_map(path: Path, entries: list[dict[str, str]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=POSITION_MAP_COLUMNS)
        writer.writeheader()
        for entry in entries:
            writer.writerow({column: entry[column] for column in POSITION_MAP_COLUMNS})


def _write_resource_index(path: Path, extraction: ExtractionResult) -> None:
    entries = [
        *(_resource_entry(embedded) for embedded in extraction.objects),
        *(_resource_entry(image) for image in extraction.images),
        *(_resource_entry(chart) for chart in extraction.charts),
        *(_resource_entry(diagram) for diagram in extraction.diagrams),
    ]
    entries.sort(key=lambda entry: (int(entry["position"]), str(entry["ref"])))
    with path.open("w", encoding="utf-8") as stream:
        for entry in entries:
            stream.write(json.dumps(entry, ensure_ascii=False) + "\n")


def _resource_entry(asset: EmbeddedObject) -> dict[str, object]:
    return {
        "resource_id": asset.resource_id or asset.ref,
        "ref": asset.ref,
        "kind": asset.kind,
        "container_path": asset.source_path,
        "relationship_type": asset.relationship_type,
        "filename": asset.filename,
        "original_filename": asset.original_filename or asset.filename,
        "content_type": asset.content_type,
        "detected_mime": asset.detected_mime,
        "extension": asset.extension,
        "size_bytes": asset.size_bytes,
        "sha256": asset.sha256,
        "position": asset.position,
        "anchor": asset.anchor or {},
        "parse_policy": asset.parse_policy or {},
        "parse_status": asset.parse_status or {},
        "risk": asset.risk or {"risk_level": "unassessed", "flags": []},
        "path": _resource_asset_path(asset),
    }


def _resource_asset_path(asset: EmbeddedObject) -> str:
    if asset.kind == "image":
        return f"structured/assets/images/{asset.path.name}"
    if asset.kind == "chart":
        return _preview_path(asset) or f"raw/charts/{asset.path.name}"
    if asset.kind == "diagram":
        return _preview_path(asset) or f"raw/diagrams/{asset.path.name}"
    return f"raw/embedded/{asset.path.name}"


def _attachment_reference(embedded: EmbeddedObject, parsed: ParsedPackage | None) -> str:
    preview_path = _preview_path(embedded)
    if parsed is not None:
        target = f"attachments/{embedded.ref}/content.md"
    elif preview_path:
        target = preview_path.removeprefix("structured/")
    else:
        target = f"../raw/embedded/{embedded.path.name}"
    return f"[{_markdown_link_text(embedded.filename)}]({target})"


def _image_reference(image: EmbeddedObject) -> str:
    return f"![](assets/images/{image.path.name})"


def _chart_reference(chart: EmbeddedObject) -> str:
    preview_path = _preview_path(chart) or f"../raw/charts/{chart.path.name}"
    base = chart.ref.replace("_", " ").title()
    chart_title = _chart_title(chart)
    title = f"{base}: {chart_title}" if chart_title else base
    return f"[{_markdown_link_text(title)}]({preview_path.removeprefix('structured/')})"


def _diagram_reference(diagram: EmbeddedObject) -> str:
    preview_path = _preview_path(diagram) or f"../raw/diagrams/{diagram.path.name}"
    base = f"SmartArt {diagram.ref.rsplit('_', 1)[-1]}"
    diagram_title = _diagram_title(diagram)
    title = f"{base}: {diagram_title}" if diagram_title else base
    return f"[{_markdown_link_text(title)}]({preview_path.removeprefix('structured/')})"


def _chart_title(chart: EmbeddedObject) -> str:
    try:
        root = ET.fromstring(chart.path.read_bytes())
    except (OSError, ET.ParseError):
        return ""
    for text in root.findall(f".//{{{CHART_NS}}}title//{{{DRAWING_NS}}}t"):
        value = (text.text or "").strip()
        if value:
            return value
    return ""


def _diagram_title(diagram: EmbeddedObject) -> str:
    try:
        root = ET.fromstring(diagram.path.read_bytes())
    except (OSError, ET.ParseError):
        return ""
    texts = []
    for text in root.findall(f".//{{{DRAWING_NS}}}t"):
        value = (text.text or "").strip()
        if value and value not in texts:
            texts.append(value)
    return " / ".join(texts[:4])


def _manifest_entry(embedded: EmbeddedObject, parsed: ParsedPackage | None) -> dict[str, object]:
    if parsed is None:
        preview_path = _preview_path(embedded)
        return {
            "ref": embedded.ref,
            "filename": embedded.filename,
            "type": embedded.type,
            "description": (
                "Embedded attachment with shallow preview"
                if preview_path
                else "Embedded attachment preserved without parsing"
            ),
            "path": f"raw/embedded/{embedded.path.name}",
            "entry_point": preview_path,
            "tables": [],
        }

    return {
        "ref": embedded.ref,
        "filename": embedded.filename,
        "type": "xlsx",
        "description": "Embedded XLSX workbook",
        "path": f"structured/attachments/{embedded.ref}/",
        "entry_point": "content.md",
        "tables": _table_paths(embedded.ref, parsed),
    }


def _image_manifest_entry(image: EmbeddedObject) -> dict[str, object]:
    return {
        "ref": image.ref,
        "filename": image.filename,
        "type": image.type,
        "content_type": image.content_type,
        "path": f"structured/assets/images/{image.path.name}",
        "source_path": image.source_path,
    }


def _chart_manifest_entry(chart: EmbeddedObject) -> dict[str, object]:
    preview_path = _preview_path(chart)
    return {
        "ref": chart.ref,
        "filename": chart.filename,
        "type": chart.type,
        "description": "Word chart with shallow preview" if preview_path else "Word chart preserved without preview",
        "path": preview_path or f"raw/charts/{chart.path.name}",
        "entry_point": preview_path,
        "source_path": chart.source_path,
    }


def _diagram_manifest_entry(diagram: EmbeddedObject) -> dict[str, object]:
    preview_path = _preview_path(diagram)
    return {
        "ref": diagram.ref,
        "filename": diagram.filename,
        "type": diagram.type,
        "description": (
            "Word SmartArt with shallow preview"
            if preview_path
            else "Word SmartArt preserved without preview"
        ),
        "path": preview_path or f"raw/diagrams/{diagram.path.name}",
        "entry_point": preview_path,
        "source_path": diagram.source_path,
    }


def _table_paths(ref: str, parsed: ParsedPackage | None) -> list[str]:
    if parsed is None:
        return []
    return [
        f"structured/attachments/{ref}/tables/{table.csv_path.name}"
        for table in parsed.tables
    ]


def _preview_path(embedded: EmbeddedObject) -> str | None:
    preview_path = embedded.parse_status.get("preview_path")
    return preview_path if isinstance(preview_path, str) else None


def _copy_file(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() != destination.resolve():
        shutil.copy2(source, destination)


def _render_parse_log(status: str, warnings: list[str]) -> str:
    lines = [f"status: {status}"]
    if warnings:
        lines.append("warnings:")
        lines.extend(f"- {warning}" for warning in warnings)
    else:
        lines.append("warnings: []")
    return "\n".join(lines) + "\n"


def _yaml_scalar(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


def _markdown_link_text(value: str) -> str:
    return value.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]")


def _markdown_table_cell(value: str) -> str:
    return value.replace("\\", "\\\\").replace("|", "\\|").replace("\n", " ")
