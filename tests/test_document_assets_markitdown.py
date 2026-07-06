from __future__ import annotations

import csv
import hashlib
import json
import struct
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from conftest import OFFICE_NS, PKG_REL_NS, REL_NS, WORD_NS, make_xlsx_bytes
from edp import extract_document_assets, merge_parent_with_attachments, parse_attachment_package
from edp.merger import _replace_asset_sentinels


DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DIAGRAM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


def test_docx_image_is_extracted_and_reinserted_from_markitdown(
    tmp_path: Path, monkeypatch
) -> None:
    docx = _write_docx_with_image(tmp_path / "image.docx")

    def fake_convert(input_path: Path) -> str:
        assert input_path == tmp_path / "package" / "structured" / "_meta" / "work" / "clean" / "main.docx"
        with ZipFile(input_path) as archive:
            patched_xml = archive.read("word/document.xml").decode("utf-8")
        assert "[[EDP_ASSET:image_01]]" in patched_xml
        return "Intro\n\n[[EDP_ASSET:image_01]]\n\nOutro"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    content = package.content_path.read_text(encoding="utf-8")
    assert "Intro" in content
    assert "![](assets/images/image_01.png)" in content
    assert "../raw/embedded/image_" not in content
    assert "image:" not in content
    assert "Outro" in content
    top_level_dirs = {
        path.name for path in (tmp_path / "package").iterdir() if path.is_dir()
    }
    assert top_level_dirs == {"raw", "structured"}
    assert not (tmp_path / "package" / "assets").exists()
    assert not (tmp_path / "package" / "__meta__").exists()
    assert (tmp_path / "package" / "structured" / "assets" / "images" / "image_01.png").exists()
    assert (tmp_path / "package" / "structured" / "_meta" / "work" / "clean" / "main.docx").exists()
    assert (tmp_path / "package" / "structured" / "_meta" / "parsers" / "markitdown" / "main.md").exists()

    child_files = (tmp_path / "package" / "structured" / "child_files.md").read_text(
        encoding="utf-8"
    )
    assert "| image_01 | image | image1.png | structured/assets/images/image_01.png |  |  |" in child_files
    assert "## Reference Definitions" not in child_files
    assert "## YAML Index" in child_files
    assert 'ref: "image_01"' in child_files
    assert 'markdown_reference: "![](assets/images/image_01.png)"' in child_files

    with (tmp_path / "package" / "structured" / "position_map.csv").open(
        newline="", encoding="utf-8"
    ) as stream:
        rows = list(csv.DictReader(stream))
    assert rows == [
        {
            "position": "1",
            "ref": "image_01",
            "kind": "image",
            "filename": "image1.png",
            "type": "png",
            "source_path": "word/media/image1.png",
            "markdown_reference": "![](assets/images/image_01.png)",
            "target_path": "structured/assets/images/image_01.png",
            "entry_point": "",
            "tables": "",
        }
    ]

    resource_entries = [
        json.loads(line)
        for line in (tmp_path / "package" / "structured" / "embedded_resources.jsonl")
        .read_text(encoding="utf-8")
        .splitlines()
    ]
    assert resource_entries[0]["path"] == "structured/assets/images/image_01.png"

    manifest = json.loads(package.manifest_path.read_text(encoding="utf-8"))
    assert manifest["content_map"]["child_files"] == "structured/child_files.md"
    assert manifest["content_map"]["clean_main"] == "structured/_meta/work/clean/main.docx"
    assert manifest["content_map"]["main_parser"] == "markitdown"
    assert manifest["content_map"]["main_parser_artifacts"] == "structured/_meta/parsers/markitdown/"
    assert manifest["content_map"]["position_map"] == "structured/position_map.csv"
    assert manifest["content_stats"]["image_count"] == 1
    assert manifest["content_map"]["images"] == [
        {
            "ref": "image_01",
            "filename": "image1.png",
            "type": "png",
            "content_type": "image/png",
            "path": "structured/assets/images/image_01.png",
            "source_path": "word/media/image1.png",
        }
    ]


def test_clean_main_docx_preserves_word_namespace_prefixes(
    tmp_path: Path, monkeypatch
) -> None:
    docx = _write_docx_with_image(tmp_path / "image-prefixes.docx")

    def fake_convert(input_path: Path) -> str:
        with ZipFile(input_path) as archive:
            patched_xml = archive.read("word/document.xml").decode("utf-8")
        assert 'xmlns:w="' in patched_xml
        assert "<w:document" in patched_xml
        assert "ns0:" not in patched_xml
        return "Intro\n\n[[EDP_ASSET:image_01]]\n\nOutro"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")

    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    assert package.warnings == []


def test_pipeline_markitdown_normalizes_docx_notes(tmp_path: Path, monkeypatch) -> None:
    docx = _write_docx_with_notes(tmp_path / "notes.docx")
    broken_markdown = (
        "Alpha [[1]](#footnote-1) beta.\n\n"
        "Omega [[2]](#endnote-1) done.\n\n"
        "1. Foot note text [↑](#footnote-ref-1)\n"
        "2. End note text [↑](#endnote-ref-1)\n"
    )
    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", lambda _path: broken_markdown)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    content = package.content_path.read_text(encoding="utf-8")
    assert "[^1]" in content
    assert "[^1]: Foot note text" in content
    assert "[[1]](#footnote-1)" not in content
    assert "[↑](#footnote-ref-1)" not in content


def test_vml_image_is_reinserted_in_place_from_markitdown(
    tmp_path: Path, monkeypatch
) -> None:
    docx = _write_docx_with_vml_image(tmp_path / "vml-image.docx")

    def fake_convert(input_path: Path) -> str:
        with ZipFile(input_path) as archive:
            patched_xml = archive.read("word/document.xml").decode("utf-8")
        assert "Intro" in patched_xml
        assert "[[EDP_ASSET:image_01]]" in patched_xml
        assert "rId1" not in patched_xml
        return "Intro\n\n[[EDP_ASSET:image_01]]\n\nOutro"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    content = package.content_path.read_text(encoding="utf-8")
    assert content == "Intro\n\n![](assets/images/image_01.png)\n\nOutro\n"
    assert (tmp_path / "package" / "structured" / "assets" / "images" / "image_01.png").exists()


def test_table_cell_image_sentinel_is_replaced_without_fallback_append(
    tmp_path: Path, monkeypatch
) -> None:
    docx = _write_docx_with_table_cell_image(tmp_path / "table-image.docx")

    def fake_convert(input_path: Path) -> str:
        with ZipFile(input_path) as archive:
            patched_xml = archive.read("word/document.xml").decode("utf-8")
        assert "<w:tc>" in patched_xml
        assert "[[EDP_ASSET:image_01]]" in patched_xml
        return "| Label | Asset |\n| --- | --- |\n| Logo | [[EDP_ASSET:image_01]] |\n"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    content = package.content_path.read_text(encoding="utf-8")
    assert content == "| Label | Asset |\n| --- | --- |\n| Logo | ![](assets/images/image_01.png) |\n"


def test_table_cell_attachment_sentinel_is_replaced_without_fallback_append(
    tmp_path: Path, monkeypatch
) -> None:
    xlsx = make_xlsx_bytes([["Metric", "Value"], ["Revenue", 1200]], sheet_name="Q3")
    docx = _write_docx_with_table_cell_embedding(tmp_path / "table-attachment.docx", "book.xlsx", xlsx)

    def fake_convert(input_path: Path) -> str:
        with ZipFile(input_path) as archive:
            patched_xml = archive.read("word/document.xml").decode("utf-8")
        assert "<w:tc>" in patched_xml
        assert "[[EDP_ASSET:attachment_01]]" in patched_xml
        return "| Label | Asset |\n| --- | --- |\n| Workbook | [[EDP_ASSET:attachment_01]] |\n"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    content = package.content_path.read_text(encoding="utf-8")
    assert content == (
        "| Label | Asset |\n"
        "| --- | --- |\n"
        "| Workbook | [book.xlsx](resources/attachment_01/preview.md) |\n"
    )


def test_ole_preview_image_is_not_registered_as_standalone_asset(
    tmp_path: Path, monkeypatch
) -> None:
    xlsx = make_xlsx_bytes([["Metric", "Value"], ["Revenue", 1200]], sheet_name="Q3")
    docx = _write_docx_with_ole_preview_image(
        tmp_path / "ole-preview.docx", "book.xlsx", xlsx
    )

    def fake_convert(input_path: Path) -> str:
        with ZipFile(input_path) as archive:
            patched_xml = archive.read("word/document.xml").decode("utf-8")
        assert "[[EDP_ASSET:attachment_01]]" in patched_xml
        assert "image_01" not in patched_xml
        return "Intro\n\n[[EDP_ASSET:attachment_01]]\n\nOutro"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")

    assert [asset.ref for asset in extraction.objects] == ["attachment_01"]
    assert extraction.images == []

    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")
    content = package.content_path.read_text(encoding="utf-8")
    assert content == "Intro\n\n[book.xlsx](resources/attachment_01/preview.md)\n\nOutro\n"
    assert not (tmp_path / "package" / "structured" / "assets" / "images" / "image_01.png").exists()


def test_docx_chart_is_extracted_as_structured_resource(tmp_path: Path, monkeypatch) -> None:
    docx = _write_docx_with_chart(tmp_path / "chart.docx")

    def fake_convert(input_path: Path) -> str:
        with ZipFile(input_path) as archive:
            patched_xml = archive.read("word/document.xml").decode("utf-8")
        assert "[[EDP_ASSET:chart_01]]" in patched_xml
        assert "rIdChart1" not in patched_xml
        return "Intro\n\n[[EDP_ASSET:chart_01]]\n\nOutro"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    assert [(chart.ref, chart.kind, chart.filename) for chart in extraction.charts] == [
        ("chart_01", "chart", "chart1.xml")
    ]

    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    content = package.content_path.read_text(encoding="utf-8")
    assert content == "Intro\n\n[Chart 01: I am a diagram](resources/chart_01/preview.md)\n\nOutro\n"

    preview = (tmp_path / "package" / "structured" / "resources" / "chart_01" / "preview.md").read_text(
        encoding="utf-8"
    )
    assert "# Chart Preview chart_01" in preview
    assert "Title: I am a diagram" in preview
    assert "| Category | 系列 1 | 系列 2 |" in preview
    assert "| 类别1 | 4.3 | 2.4 |" in preview

    preview_json = json.loads(
        (tmp_path / "package" / "structured" / "resources" / "chart_01" / "preview.json").read_text(
            encoding="utf-8"
        )
    )
    assert preview_json["title"] == "I am a diagram"
    assert preview_json["series"] == ["系列 1", "系列 2"]
    assert preview_json["categories"] == ["类别1", "类别2"]
    assert preview_json["data_rows"] == [
        {"category": "类别1", "values": {"系列 1": 4.3, "系列 2": 2.4}},
        {"category": "类别2", "values": {"系列 1": 2.5, "系列 2": 4.4}},
    ]

    resources = [
        json.loads(line)
        for line in (tmp_path / "package" / "structured" / "embedded_resources.jsonl")
        .read_text(encoding="utf-8")
        .splitlines()
        if line
    ]
    assert resources == [
        {
            "resource_id": "chart_01",
            "ref": "chart_01",
            "kind": "chart",
            "container_path": "word/charts/chart1.xml",
            "relationship_type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            "filename": "chart1.xml",
            "original_filename": "chart1.xml",
            "content_type": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            "detected_mime": "application/xml",
            "extension": ".xml",
            "size_bytes": resources[0]["size_bytes"],
            "sha256": resources[0]["sha256"],
            "position": 1,
            "anchor": {"block_id": "p_0002", "paragraph_text": "", "page": None},
            "parse_policy": {"mode": "shallow_preview", "reason": "allowlisted chart preview"},
            "parse_status": {
                "status": "preview_extracted",
                "preview_path": "structured/resources/chart_01/preview.md",
                "preview_json_path": "structured/resources/chart_01/preview.json",
                "full_parse": False,
            },
            "risk": {"risk_level": "unassessed", "flags": []},
            "path": "structured/resources/chart_01/preview.md",
        }
    ]

    manifest = json.loads(package.manifest_path.read_text(encoding="utf-8"))
    assert manifest["content_map"]["charts"] == [
        {
            "ref": "chart_01",
            "filename": "chart1.xml",
            "type": "chart",
            "description": "Word chart with shallow preview",
            "path": "structured/resources/chart_01/preview.md",
            "entry_point": "structured/resources/chart_01/preview.md",
            "source_path": "word/charts/chart1.xml",
        }
    ]


def test_docx_smartart_is_extracted_as_structured_resource(
    tmp_path: Path, monkeypatch
) -> None:
    docx = _write_docx_with_smartart(tmp_path / "smartart.docx")

    def fake_convert(input_path: Path) -> str:
        with ZipFile(input_path) as archive:
            patched_xml = archive.read("word/document.xml").decode("utf-8")
        assert "[[EDP_ASSET:diagram_01]]" in patched_xml
        assert "rIdDiagramData1" not in patched_xml
        return "Intro\n\n[[EDP_ASSET:diagram_01]]\n\nOutro"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    assert [(diagram.ref, diagram.kind, diagram.filename) for diagram in extraction.diagrams] == [
        ("diagram_01", "diagram", "data1.xml")
    ]

    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    content = package.content_path.read_text(encoding="utf-8")
    assert content == (
        "Intro\n\n"
        "[SmartArt 01: one / two / three / four](resources/diagram_01/preview.md)\n\n"
        "Outro\n"
    )

    preview = (
        tmp_path / "package" / "structured" / "resources" / "diagram_01" / "preview.md"
    ).read_text(encoding="utf-8")
    assert "# SmartArt Preview diagram_01" in preview
    assert "Source part: `word/diagrams/data1.xml`" in preview
    assert "- one" in preview
    assert "- four" in preview

    preview_json = json.loads(
        (
            tmp_path / "package" / "structured" / "resources" / "diagram_01" / "preview.json"
        ).read_text(encoding="utf-8")
    )
    assert preview_json["texts"] == ["one", "two", "three", "four"]
    assert preview_json["companion_parts"] == [
        "word/diagrams/drawing1.xml",
        "word/diagrams/layout1.xml",
        "word/diagrams/quickStyle1.xml",
        "word/diagrams/colors1.xml",
    ]

    resources = [
        json.loads(line)
        for line in (tmp_path / "package" / "structured" / "embedded_resources.jsonl")
        .read_text(encoding="utf-8")
        .splitlines()
        if line
    ]
    assert resources == [
        {
            "resource_id": "diagram_01",
            "ref": "diagram_01",
            "kind": "diagram",
            "container_path": "word/diagrams/data1.xml",
            "relationship_type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
            "filename": "data1.xml",
            "original_filename": "data1.xml",
            "content_type": "application/xml",
            "detected_mime": "application/xml",
            "extension": ".xml",
            "size_bytes": resources[0]["size_bytes"],
            "sha256": resources[0]["sha256"],
            "position": 1,
            "anchor": {"block_id": "p_0002", "paragraph_text": "", "page": None},
            "parse_policy": {"mode": "shallow_preview", "reason": "allowlisted SmartArt preview"},
            "parse_status": {
                "status": "preview_extracted",
                "preview_path": "structured/resources/diagram_01/preview.md",
                "preview_json_path": "structured/resources/diagram_01/preview.json",
                "full_parse": False,
            },
            "risk": {"risk_level": "unassessed", "flags": []},
            "path": "structured/resources/diagram_01/preview.md",
        }
    ]

    manifest = json.loads(package.manifest_path.read_text(encoding="utf-8"))
    assert manifest["content_map"]["diagrams"] == [
        {
            "ref": "diagram_01",
            "filename": "data1.xml",
            "type": "diagram",
            "description": "Word SmartArt with shallow preview",
            "path": "structured/resources/diagram_01/preview.md",
            "entry_point": "structured/resources/diagram_01/preview.md",
            "source_path": "word/diagrams/data1.xml",
        }
    ]


def test_non_xlsx_attachment_is_preserved_without_parsing(tmp_path: Path) -> None:
    docx = _write_docx_with_embedding(tmp_path / "notes.docx", "notes.txt", b"raw notes\n")

    extraction = extract_document_assets(docx, tmp_path / "extract")
    assert [(obj.ref, obj.filename, obj.type) for obj in extraction.objects] == [
        ("attachment_01", "notes.txt", "txt")
    ]
    assert extraction.objects[0].path.read_bytes() == b"raw notes\n"

    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")
    content = package.content_path.read_text(encoding="utf-8")
    assert "[notes.txt](resources/attachment_01/preview.md)" in content
    assert "embedded_object:" not in content

    raw_attachment = tmp_path / "package" / "raw" / "embedded" / "attachment_01.txt"
    assert raw_attachment.read_bytes() == b"raw notes\n"
    assert not (tmp_path / "package" / "structured" / "attachments").exists()
    assert not (tmp_path / "package" / "structured" / "attachments" / "attachment_01").exists()
    preview = tmp_path / "package" / "structured" / "resources" / "attachment_01" / "preview.md"
    assert preview.read_text(encoding="utf-8") == "# Resource Preview attachment_01\n\n```text\nraw notes\n\n```\n"

    child_files = (tmp_path / "package" / "structured" / "child_files.md").read_text(
        encoding="utf-8"
    )
    assert (
        "| attachment_01 | attachment | notes.txt | structured/resources/attachment_01/preview.md | "
        "structured/resources/attachment_01/preview.md |  |"
    ) in child_files
    assert "## Reference Definitions" not in child_files
    assert 'markdown_reference: "[notes.txt](resources/attachment_01/preview.md)"' in child_files

    manifest = json.loads(package.manifest_path.read_text(encoding="utf-8"))
    assert manifest["content_map"]["embedded_objects"] == [
        {
            "ref": "attachment_01",
            "filename": "notes.txt",
            "type": "txt",
            "description": "Embedded attachment with shallow preview",
            "path": "raw/embedded/attachment_01.txt",
            "entry_point": "structured/resources/attachment_01/preview.md",
            "tables": [],
        }
    ]
    assert manifest["content_map"]["embedded_resources"] == "structured/embedded_resources.jsonl"


def test_embedded_resources_jsonl_records_assets_metadata_and_anchor(
    tmp_path: Path, monkeypatch
) -> None:
    xlsx = make_xlsx_bytes([["Name", "Score"], ["alpha", 7]], sheet_name="Scores")
    docx = _write_docx_with_embedding_and_image(tmp_path / "mixed.docx", "book.xlsx", xlsx)

    def fake_convert(input_path: Path) -> str:
        return "Intro\n\n[[EDP_ASSET:attachment_01]]\n\n[[EDP_ASSET:image_01]]\n\nOutro"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    resources_path = tmp_path / "package" / "structured" / "embedded_resources.jsonl"
    resources = [
        json.loads(line)
        for line in resources_path.read_text(encoding="utf-8").splitlines()
        if line
    ]

    assert [resource["ref"] for resource in resources] == ["attachment_01", "image_01"]
    attachment = resources[0]
    assert attachment["kind"] == "attachment"
    assert attachment["container_path"] == "word/embeddings/book.xlsx"
    assert attachment["relationship_type"].endswith("/package")
    assert attachment["original_filename"] == "book.xlsx"
    assert attachment["detected_mime"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert attachment["extension"] == ".xlsx"
    assert attachment["size_bytes"] == len(xlsx)
    assert attachment["sha256"] == hashlib.sha256(xlsx).hexdigest()
    assert attachment["anchor"] == {
        "block_id": "p_0002",
        "paragraph_text": "",
        "page": None,
    }
    assert attachment["parse_policy"] == {
        "mode": "shallow_preview",
        "reason": "allowlisted attachment preview",
    }
    assert attachment["parse_status"] == {
        "status": "preview_extracted",
        "preview_path": "structured/resources/attachment_01/preview.md",
        "preview_json_path": "structured/resources/attachment_01/preview.json",
        "full_parse": False,
    }
    assert attachment["risk"] == {"risk_level": "unassessed", "flags": []}

    image = resources[1]
    assert image["kind"] == "image"
    assert image["detected_mime"] == "image/png"
    assert image["parse_policy"]["mode"] == "preserved_image"
    assert image["parse_status"] == {
        "status": "preserved",
        "preview_path": None,
        "preview_json_path": None,
        "full_parse": False,
    }
    assert package.warnings == []


def test_xlsx_attachment_gets_shallow_preview_without_full_table_package(
    tmp_path: Path, monkeypatch
) -> None:
    xlsx = make_xlsx_bytes([["Metric", "Value"], ["Revenue", 1200], ["Profit", 320]], sheet_name="Q3")
    docx = _write_docx_with_embedding(tmp_path / "book.docx", "book.xlsx", xlsx)

    def fake_convert(input_path: Path) -> str:
        return "Intro\n\n[[EDP_ASSET:attachment_01]]\n\nOutro"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    content = package.content_path.read_text(encoding="utf-8")
    assert "[book.xlsx](resources/attachment_01/preview.md)" in content
    preview = (
        tmp_path / "package" / "structured" / "resources" / "attachment_01" / "preview.md"
    ).read_text(encoding="utf-8")
    assert "# Resource Preview attachment_01" in preview
    assert "## Sheet: Q3" in preview
    assert "| Metric | Value |" in preview
    assert "| Revenue | 1200 |" in preview

    preview_json = json.loads(
        (
            tmp_path / "package" / "structured" / "resources" / "attachment_01" / "preview.json"
        ).read_text(encoding="utf-8")
    )
    assert preview_json["type"] == "xlsx"
    assert preview_json["sheets"][0]["name"] == "Q3"
    assert preview_json["sheets"][0]["preview_rows"][2] == ["Profit", 320]
    assert not (
        tmp_path
        / "package"
        / "structured"
        / "attachments"
        / "attachment_01"
        / "tables"
        / "table_001.csv"
    ).exists()


def test_replace_asset_sentinels_handles_parser_escaping() -> None:
    """Sentinels must be replaced in place regardless of parser escaping.

    markitdown escapes underscores (``EDP\\_ASSET``, ``image\\_01``); pandoc
    escapes brackets (``\\[\\[...\\]\\]``). Both must resolve to the placeholder
    inline, with no literal ``EDP_ASSET`` left behind and no end-of-doc append.
    """
    image_placeholder = {"image_01": "![](assets/images/image_01.png)"}
    placeholders = {
        "image_01": "![](assets/images/image_01.png)",
        "attachment_01": "[book.xlsx](attachments/attachment_01/content.md)",
    }

    forms = [
        ("raw", "[[EDP_ASSET:image_01]]"),
        ("markitdown", "[[EDP\\_ASSET:image\\_01]]"),
        ("pandoc", "\\[\\[EDP_ASSET:image_01\\]\\]"),
    ]
    for _label, sentinel in forms:
        markdown = f"Intro\n\n{sentinel}\n\nOutro"
        rendered = _replace_asset_sentinels(markdown, image_placeholder)
        assert "EDP_ASSET" not in rendered, sentinel
        assert "Intro\n\n![](assets/images/image_01.png)\n\nOutro\n" == rendered, sentinel

    # Mixed escaping within one document.
    markdown = (
        "Intro\n\n\\[\\[EDP_ASSET:attachment_01\\]\\]\n\n"
        "[[EDP\\_ASSET:image\\_01]]\n\nOutro"
    )
    rendered = _replace_asset_sentinels(markdown, placeholders)
    assert "EDP_ASSET" not in rendered
    assert rendered == (
        "Intro\n\n[book.xlsx](attachments/attachment_01/content.md)\n\n"
        "![](assets/images/image_01.png)\n\nOutro\n"
    )


def test_replace_asset_sentinels_appends_missing_refs() -> None:
    """A sentinel the parser dropped entirely is still appended at the end."""
    placeholders = {"image_01": "![](assets/images/image_01.png)"}
    rendered = _replace_asset_sentinels("Intro\n\nOutro", placeholders)
    assert rendered == "Intro\n\nOutro\n\n![](assets/images/image_01.png)\n"


def test_pdf_and_unknown_bin_are_registered_without_preview(tmp_path: Path, monkeypatch) -> None:
    docx = _write_docx_with_embedding(tmp_path / "payload.docx", "oleObject1.bin", b"ole payload")
    monkeypatch.setattr(
        "edp.extractor._extract_ole10_native_payload",
        lambda payload: ("payload.pdf", b"%PDF-1.7\n"),
    )

    extraction = extract_document_assets(docx, tmp_path / "extract")
    merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    resources_path = tmp_path / "package" / "structured" / "embedded_resources.jsonl"
    [resource] = [
        json.loads(line)
        for line in resources_path.read_text(encoding="utf-8").splitlines()
        if line
    ]
    assert resource["filename"] == "payload.pdf"
    assert resource["detected_mime"] == "application/pdf"
    assert resource["parse_policy"] == {
        "mode": "extracted_only",
        "reason": "preview disabled for this attachment type",
    }
    assert resource["parse_status"] == {
        "status": "extracted_only",
        "preview_path": None,
        "preview_json_path": None,
        "full_parse": False,
    }
    assert not (tmp_path / "package" / "structured" / "resources" / "attachment_01").exists()


def test_parent_markitdown_failure_falls_back_to_xml_text_and_marks_partial(
    tmp_path: Path, monkeypatch
) -> None:
    xlsx = make_xlsx_bytes([["Metric"], ["Fallback"]])
    docx = _write_docx_with_embedding(tmp_path / "fallback.docx", "book.xlsx", xlsx)

    def fail_convert(input_path: Path) -> str:
        raise RuntimeError("boom")

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fail_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    parsed = [
        parse_attachment_package(extraction.objects[0], tmp_path / "parsed" / "attachment_01")
    ]
    package = merge_parent_with_attachments(docx, extraction, parsed, tmp_path / "package")

    content = package.content_path.read_text(encoding="utf-8")
    assert "Intro" in content
    assert "[book.xlsx](attachments/attachment_01/content.md)" in content
    assert "Outro" in content
    assert package.warnings == ["MarkItDown failed to convert parent document: boom"]

    manifest = json.loads(package.manifest_path.read_text(encoding="utf-8"))
    assert manifest["parse_status"] == "partial"


def test_position_map_uses_document_order_across_attachments_and_images(
    tmp_path: Path, monkeypatch
) -> None:
    xlsx = make_xlsx_bytes([["Metric"], ["Order"]])
    docx = _write_docx_with_embedding_and_image(
        tmp_path / "mixed-assets.docx", "book.xlsx", xlsx
    )

    def fake_convert(input_path: Path) -> str:
        return "Intro\n\n[[EDP_ASSET:attachment_01]]\n\n[[EDP_ASSET:image_01]]\n\nOutro"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract")
    parsed = [
        parse_attachment_package(extraction.objects[0], tmp_path / "parsed" / "attachment_01")
    ]
    merge_parent_with_attachments(docx, extraction, parsed, tmp_path / "package")

    with (tmp_path / "package" / "structured" / "position_map.csv").open(
        newline="", encoding="utf-8"
    ) as stream:
        rows = list(csv.DictReader(stream))

    assert [(row["position"], row["ref"], row["kind"]) for row in rows] == [
        ("1", "attachment_01", "attachment"),
        ("2", "image_01", "image"),
    ]


def test_unknown_bin_attachment_is_preserved_for_document_asset_extraction(tmp_path: Path) -> None:
    docx = _write_docx_with_embedding(tmp_path / "unknown.docx", "oleObject1.bin", b"not ole")

    extraction = extract_document_assets(docx, tmp_path / "extract")

    assert [(obj.ref, obj.filename, obj.type) for obj in extraction.objects] == [
        ("attachment_01", "oleObject1.bin", "bin")
    ]
    assert extraction.objects[0].path.read_bytes() == b"not ole"
    assert extraction.warnings == ["Unknown OLE embedded object preserved as binary: oleObject1.bin"]


def test_bin_attachment_is_unwrapped_before_preserving(tmp_path: Path, monkeypatch) -> None:
    docx = _write_docx_with_embedding(tmp_path / "wrapped.docx", "oleObject1.bin", b"ole payload")
    monkeypatch.setattr(
        "edp.extractor._extract_ole10_native_payload",
        lambda payload: ("payload.pdf", b"%PDF-1.7\n"),
    )

    extraction = extract_document_assets(docx, tmp_path / "extract")

    assert [(obj.ref, obj.filename, obj.type) for obj in extraction.objects] == [
        ("attachment_01", "payload.pdf", "pdf")
    ]
    assert extraction.objects[0].path.name == "attachment_01.pdf"
    assert extraction.objects[0].path.read_bytes() == b"%PDF-1.7\n"
    assert extraction.warnings == []


def test_unsafe_unwrap_recovers_ole10native_source_filename_in_markdown(
    tmp_path: Path, monkeypatch
) -> None:
    xlsx = make_xlsx_bytes([["Order"], [123]], sheet_name="Orders")
    ole_payload = _ole10_native_payload("Orders.xlsx", xlsx)
    docx = _write_docx_with_embedding(tmp_path / "orders.docx", "oleObject1.bin", ole_payload)

    def fake_convert(input_path: Path) -> str:
        with ZipFile(input_path) as archive:
            patched_xml = archive.read("word/document.xml").decode("utf-8")
        assert "[[EDP_ASSET:attachment_01]]" in patched_xml
        return "Attached file: [[EDP_ASSET:attachment_01]]\n"

    monkeypatch.setattr("edp.merger._convert_docx_to_markdown", fake_convert)

    extraction = extract_document_assets(docx, tmp_path / "extract", unsafe_unwrap=True)
    package = merge_parent_with_attachments(docx, extraction, [], tmp_path / "package")

    assert [(obj.ref, obj.filename, obj.type) for obj in extraction.objects] == [
        ("attachment_01", "Orders.xlsx", "xlsx")
    ]
    assert extraction.objects[0].path.name == "attachment_01.xlsx"
    assert extraction.objects[0].path.read_bytes() == xlsx
    assert "[Orders.xlsx](resources/attachment_01/preview.md)" in package.content_path.read_text(
        encoding="utf-8"
    )


def test_unsafe_unwrap_scans_ole_streams_for_xlsx_when_native_parse_fails(
    tmp_path: Path, monkeypatch
) -> None:
    xlsx = make_xlsx_bytes([["Order"], [456]], sheet_name="Orders")
    docx = _write_docx_with_embedding(tmp_path / "orders.docx", "oleObject1.bin", b"ole payload")
    monkeypatch.setattr("edp.extractor._extract_ole10_native_payload", lambda payload: None)
    monkeypatch.setattr(
        "edp.extractor._ole_stream_payloads",
        lambda payload: [b"metadata\x00Orders.xlsx\x00", b"prefix" + xlsx],
    )

    extraction = extract_document_assets(docx, tmp_path / "extract", unsafe_unwrap=True)

    assert [(obj.ref, obj.filename, obj.type) for obj in extraction.objects] == [
        ("attachment_01", "Orders.xlsx", "xlsx")
    ]
    assert extraction.objects[0].path.name == "attachment_01.xlsx"
    assert extraction.objects[0].path.read_bytes() == xlsx
    assert extraction.warnings == []


def test_default_extraction_finds_directory_only_embedded_assets(tmp_path: Path) -> None:
    docx = _write_docx_with_unrelated_object_part(tmp_path / "objects.docx")

    extraction = extract_document_assets(docx, tmp_path / "extract")

    assert [(obj.ref, obj.filename, obj.type, obj.source_path) for obj in extraction.objects] == [
        ("attachment_01", "report.pdf", "pdf", "word/objects/report.pdf")
    ]
    assert extraction.objects[0].path.read_bytes() == b"%PDF-1.7\n"
    assert extraction.objects[0].detected_mime == "application/pdf"
    assert extraction.objects[0].anchor == {}


def test_default_extraction_deduplicates_relationship_and_directory_scan(
    tmp_path: Path,
) -> None:
    xlsx = make_xlsx_bytes([["Metric"], ["Duplicate"]])
    docx = _write_docx_with_embedding_and_duplicate_object(
        tmp_path / "duplicate.docx",
        "book.xlsx",
        xlsx,
    )

    extraction = extract_document_assets(docx, tmp_path / "extract")

    assert [(obj.ref, obj.filename, obj.source_path) for obj in extraction.objects] == [
        ("attachment_01", "book.xlsx", "word/embeddings/book.xlsx")
    ]
    assert extraction.objects[0].anchor == {
        "block_id": "p_0002",
        "paragraph_text": "",
        "page": None,
    }


def _write_docx_with_image(path: Path) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" xmlns:a="{DRAWING_NS}">'
        "<w:body>"
        "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>"
        '<w:p><w:r><w:drawing><a:blip r:embed="rId1"/></w:drawing></w:r></w:p>'
        "<w:p><w:r><w:t>Outro</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/image1.png"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/media/image1.png", b"\x89PNG\r\n\x1a\n")
    return path


def _write_docx_with_vml_image(path: Path) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" xmlns:v="urn:schemas-microsoft-com:vml">'
        "<w:body>"
        "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>"
        '<w:p><w:r><w:pict><v:shape><v:imagedata r:id="rId1"/></v:shape></w:pict></w:r></w:p>'
        "<w:p><w:r><w:t>Outro</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/image1.png"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/media/image1.png", b"\x89PNG\r\n\x1a\n")
    return path


def _write_docx_with_table_cell_image(path: Path) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" xmlns:a="{DRAWING_NS}">'
        "<w:body>"
        "<w:tbl><w:tr>"
        "<w:tc><w:p><w:r><w:t>Logo</w:t></w:r></w:p></w:tc>"
        '<w:tc><w:p><w:r><w:drawing><a:blip r:embed="rId1"/></w:drawing></w:r></w:p></w:tc>'
        "</w:tr></w:tbl>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/image1.png"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/media/image1.png", b"\x89PNG\r\n\x1a\n")
    return path


def _write_docx_with_table_cell_embedding(path: Path, filename: str, payload: bytes) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" xmlns:o="{OFFICE_NS}">'
        "<w:body>"
        "<w:tbl><w:tr>"
        "<w:tc><w:p><w:r><w:t>Workbook</w:t></w:r></w:p></w:tc>"
        '<w:tc><w:p><w:r><w:object><o:OLEObject r:id="rId1"/></w:object></w:r></w:p></w:tc>'
        "</w:tr></w:tbl>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" '
        f'Target="embeddings/{filename}"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr(f"word/embeddings/{filename}", payload)
    return path


def _write_docx_with_ole_preview_image(path: Path, filename: str, payload: bytes) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" '
        f'xmlns:o="{OFFICE_NS}" xmlns:v="urn:schemas-microsoft-com:vml">'
        "<w:body>"
        "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>"
        "<w:p><w:r><w:object>"
        '<v:shape><v:imagedata r:id="rIdPreview"/></v:shape>'
        '<o:OLEObject r:id="rIdEmbed" ProgID="Excel.Sheet.12"/>'
        "</w:object></w:r></w:p>"
        "<w:p><w:r><w:t>Outro</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rIdPreview" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/image1.png"/>'
        '<Relationship Id="rIdEmbed" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" '
        f'Target="embeddings/{filename}"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/media/image1.png", b"\x89PNG\r\n\x1a\n")
        docx.writestr(f"word/embeddings/{filename}", payload)
    return path


def _write_docx_with_chart(path: Path) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" '
        f'xmlns:a="{DRAWING_NS}" xmlns:c="{CHART_NS}">'
        "<w:body>"
        "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>"
        '<w:p><w:r><w:drawing><a:graphic><a:graphicData '
        'uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart r:id="rIdChart1"/>'
        "</a:graphicData></a:graphic></w:drawing></w:r></w:p>"
        "<w:p><w:r><w:t>Outro</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rIdChart1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
        'Target="charts/chart1.xml"/>'
        "</Relationships>"
    )
    chart_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<c:chartSpace xmlns:c="{CHART_NS}" xmlns:a="{DRAWING_NS}">'
        "<c:chart>"
        "<c:title><c:tx><c:rich><a:p><a:r><a:t>I am a diagram</a:t></a:r></a:p></c:rich></c:tx></c:title>"
        "<c:plotArea><c:barChart>"
        "<c:ser>"
        '<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>系列 1</c:v></c:pt></c:strCache></c:strRef></c:tx>'
        '<c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>类别1</c:v></c:pt><c:pt idx="1"><c:v>类别2</c:v></c:pt></c:strCache></c:strRef></c:cat>'
        '<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>4.3</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt></c:numCache></c:numRef></c:val>'
        "</c:ser>"
        "<c:ser>"
        '<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>系列 2</c:v></c:pt></c:strCache></c:strRef></c:tx>'
        '<c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>类别1</c:v></c:pt><c:pt idx="1"><c:v>类别2</c:v></c:pt></c:strCache></c:strRef></c:cat>'
        '<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>2.4</c:v></c:pt><c:pt idx="1"><c:v>4.4</c:v></c:pt></c:numCache></c:numRef></c:val>'
        "</c:ser>"
        "</c:barChart></c:plotArea>"
        "</c:chart></c:chartSpace>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/charts/chart1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/charts/chart1.xml", chart_xml)
    return path


def _write_docx_with_smartart(path: Path) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" '
        f'xmlns:a="{DRAWING_NS}" xmlns:dgm="{DIAGRAM_NS}">'
        "<w:body>"
        "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>"
        "<w:p><w:r><w:drawing><a:graphic><a:graphicData "
        'uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        '<dgm:relIds r:dm="rIdDiagramData1" r:lo="rIdDiagramLayout1" '
        'r:qs="rIdDiagramQuickStyle1" r:cs="rIdDiagramColors1"/>'
        "</a:graphicData></a:graphic></w:drawing></w:r></w:p>"
        "<w:p><w:r><w:t>Outro</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rIdDiagramData1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" '
        'Target="diagrams/data1.xml"/>'
        '<Relationship Id="rIdDiagramDrawing1" '
        'Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" '
        'Target="diagrams/drawing1.xml"/>'
        '<Relationship Id="rIdDiagramLayout1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" '
        'Target="diagrams/layout1.xml"/>'
        '<Relationship Id="rIdDiagramQuickStyle1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" '
        'Target="diagrams/quickStyle1.xml"/>'
        '<Relationship Id="rIdDiagramColors1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" '
        'Target="diagrams/colors1.xml"/>'
        "</Relationships>"
    )
    data_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<dgm:dataModel xmlns:dgm="{DIAGRAM_NS}" xmlns:a="{DRAWING_NS}">'
        "<dgm:ptLst>"
        "<dgm:pt><dgm:t><a:p><a:r><a:t>one</a:t></a:r></a:p></dgm:t></dgm:pt>"
        "<dgm:pt><dgm:t><a:p><a:r><a:t>two</a:t></a:r></a:p></dgm:t></dgm:pt>"
        "<dgm:pt><dgm:t><a:p><a:r><a:t>three</a:t></a:r></a:p></dgm:t></dgm:pt>"
        "<dgm:pt><dgm:t><a:p><a:r><a:t>four</a:t></a:r></a:p></dgm:t></dgm:pt>"
        "</dgm:ptLst>"
        "</dgm:dataModel>"
    )
    drawing_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" '
        f'xmlns:a="{DRAWING_NS}"><dsp:spTree><dsp:sp><dsp:txBody><a:p><a:r><a:t>'
        "one"
        "</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/diagrams/data1.xml", data_xml)
        docx.writestr("word/diagrams/drawing1.xml", drawing_xml)
        docx.writestr("word/diagrams/layout1.xml", "<layout/>")
        docx.writestr("word/diagrams/quickStyle1.xml", "<quickStyle/>")
        docx.writestr("word/diagrams/colors1.xml", "<colors/>")
    return path


def _write_docx_with_unrelated_object_part(path: Path) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}">'
        "<w:body><w:p><w:r><w:t>Object scan</w:t></w:r></w:p></w:body>"
        "</w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}"></Relationships>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="pdf" ContentType="application/pdf"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/objects/report.pdf", b"%PDF-1.7\n")
    return path


def _write_docx_with_embedding_and_duplicate_object(
    path: Path, filename: str, payload: bytes
) -> Path:
    docx_path = _write_docx_with_embedding(path, filename, payload)
    with ZipFile(docx_path, "a", ZIP_DEFLATED) as docx:
        docx.writestr(f"word/objects/{filename}", payload)
    return docx_path


def _write_docx_with_notes(path: Path) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>'
        "<w:p>"
        "<w:r><w:t>Alpha </w:t></w:r>"
        '<w:r><w:footnoteReference w:id="1"/></w:r>'
        "<w:r><w:t> beta.</w:t></w:r>"
        "</w:p>"
        "<w:p>"
        "<w:r><w:t>Omega </w:t></w:r>"
        '<w:r><w:endnoteReference w:id="2"/></w:r>'
        "<w:r><w:t> done.</w:t></w:r>"
        "</w:p>"
        "</w:body></w:document>"
    )
    footnotes_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{WORD_NS}">'
        '<w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p/></w:footnote>'
        '<w:footnote w:id="1"><w:p><w:r><w:t>Foot note text</w:t></w:r></w:p></w:footnote>'
        "</w:footnotes>"
    )
    endnotes_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:endnotes xmlns:w="{WORD_NS}">'
        '<w:endnote w:type="separator" w:id="-1"><w:p/></w:endnote>'
        '<w:endnote w:type="continuationSeparator" w:id="0"><w:p/></w:endnote>'
        '<w:endnote w:id="2"><w:p><w:r><w:t>End note text</w:t></w:r></w:p></w:endnote>'
        "</w:endnotes>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/footnotes.xml", footnotes_xml)
        docx.writestr("word/endnotes.xml", endnotes_xml)
    return path


def _ole10_native_payload(filename: str, payload: bytes) -> bytes:
    native = (
        struct.pack("<I", 0)
        + filename.encode("utf-8")
        + b"\x00"
        + b"\x00"
        + (b"\x00" * 8)
        + b"\x00"
        + struct.pack("<I", len(payload))
        + payload
    )
    return b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + native


def _write_docx_with_embedding(path: Path, filename: str, payload: bytes) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" xmlns:o="{OFFICE_NS}">'
        "<w:body>"
        "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>"
        '<w:p><w:r><w:object><o:OLEObject r:id="rId1"/></w:object></w:r></w:p>'
        "<w:p><w:r><w:t>Outro</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" '
        f'Target="embeddings/{filename}"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="txt" ContentType="text/plain"/>'
        '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>'
        '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr(f"word/embeddings/{filename}", payload)
    return path


def _write_docx_with_embedding_and_image(path: Path, filename: str, payload: bytes) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" '
        f'xmlns:o="{OFFICE_NS}" xmlns:a="{DRAWING_NS}">'
        "<w:body>"
        "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>"
        '<w:p><w:r><w:object><o:OLEObject r:id="rId1"/></w:object></w:r></w:p>'
        '<w:p><w:r><w:drawing><a:blip r:embed="rId2"/></w:drawing></w:r></w:p>'
        "<w:p><w:r><w:t>Outro</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" '
        f'Target="embeddings/{filename}"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/image1.png"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr(f"word/embeddings/{filename}", payload)
        docx.writestr("word/media/image1.png", b"\x89PNG\r\n\x1a\n")
    return path
