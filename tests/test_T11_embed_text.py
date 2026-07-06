from __future__ import annotations

import json
from pathlib import Path

from edp import extract_embedded_xlsx, merge_parent_with_attachments, parse_xlsx_package


def test_extracts_embedded_xlsx_with_stable_ref_and_manifest(
    embedded_xlsx_docx: Path, tmp_path: Path
) -> None:
    extraction = extract_embedded_xlsx(embedded_xlsx_docx, tmp_path / "extract")
    assert extraction.warnings == []
    assert [obj.ref for obj in extraction.objects] == ["attachment_01"]

    embedded = extraction.objects[0]
    assert embedded.filename == "profit.xlsx"
    assert embedded.rel_id == "rId1"
    assert embedded.path.name == "attachment_01.xlsx"
    assert embedded.path.exists()

    parsed = parse_xlsx_package(embedded.path, tmp_path / "parsed" / embedded.ref, embedded.ref)
    package = merge_parent_with_attachments(
        embedded_xlsx_docx,
        extraction,
        [parsed],
        tmp_path / "package",
    )

    manifest = json.loads(package.manifest_path.read_text(encoding="utf-8"))
    assert manifest["parse_status"] == "success"
    assert manifest["content_map"]["embedded_objects"] == [
        {
            "ref": "attachment_01",
            "filename": "profit.xlsx",
            "type": "xlsx",
            "description": "Embedded XLSX workbook",
            "path": "structured/attachments/attachment_01/",
            "entry_point": "content.md",
            "tables": ["structured/attachments/attachment_01/tables/table_001.csv"],
        }
    ]


def test_docx_without_embedded_objects_succeeds_with_empty_attachment_list(tmp_path: Path) -> None:
    from tests.conftest import make_docx_with_embeddings

    docx_path = make_docx_with_embeddings(tmp_path / "plain.docx", [])
    extraction = extract_embedded_xlsx(docx_path, tmp_path / "extract")
    package = merge_parent_with_attachments(docx_path, extraction, [], tmp_path / "package")

    manifest = json.loads(package.manifest_path.read_text(encoding="utf-8"))
    assert manifest["parse_status"] == "success"
    assert manifest["content_map"]["embedded_objects"] == []
    assert "Quarterly report" in package.content_path.read_text(encoding="utf-8")
