from __future__ import annotations

import csv
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from edp import extract_embedded_xlsx, parse_xlsx_package
from edp import merge_parent_with_attachments
from tests.conftest import OFFICE_NS, PKG_REL_NS, REL_NS, WORD_NS


def test_xlsx_package_contains_markdown_csv_and_json(
    embedded_xlsx_docx: Path, tmp_path: Path
) -> None:
    extraction = extract_embedded_xlsx(embedded_xlsx_docx, tmp_path / "extract")
    parsed = parse_xlsx_package(
        extraction.objects[0].path,
        tmp_path / "parsed" / "attachment_01",
        "attachment_01",
    )

    assert parsed.content_path.exists()
    assert [table.csv_path.name for table in parsed.tables] == ["table_001.csv"]
    assert [table.json_path.name for table in parsed.tables] == ["table_001.json"]

    content = parsed.content_path.read_text(encoding="utf-8")
    assert "# Embedded Workbook attachment_01" in content
    assert "## Sheet: Q3" in content
    assert "Revenue" in content
    assert "1200" in content

    with parsed.tables[0].csv_path.open(newline="", encoding="utf-8") as stream:
        rows = list(csv.reader(stream))
    assert rows[1] == ["Revenue", "1200"]

    table_json = json.loads(parsed.tables[0].json_path.read_text(encoding="utf-8"))
    assert table_json["sheet"] == "Q3"
    assert table_json["rows"][2] == ["Profit", 320]


def test_multiple_xlsx_embeddings_are_numbered_in_document_order(tmp_path: Path) -> None:
    from tests.conftest import make_docx_with_embeddings, make_xlsx_bytes

    first = make_xlsx_bytes([["first"], [1]], sheet_name="First")
    second = make_xlsx_bytes([["second"], [2]], sheet_name="Second")
    docx = make_docx_with_embeddings(
        tmp_path / "multi.docx",
        [
            ("second-created.xlsx", first),
            ("first-created.xlsx", second),
        ],
    )

    extraction = extract_embedded_xlsx(docx, tmp_path / "extract")

    assert [(obj.ref, obj.filename, obj.position) for obj in extraction.objects] == [
        ("attachment_01", "second-created.xlsx", 1),
        ("attachment_02", "first-created.xlsx", 2),
    ]


def test_inline_placeholder_stays_between_surrounding_text(tmp_path: Path) -> None:
    from tests.conftest import make_xlsx_bytes

    xlsx = make_xlsx_bytes([["Metric"], ["Inline"]])
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" xmlns:o="{OFFICE_NS}">'
        "<w:body><w:p>"
        "<w:r><w:t>Before object</w:t></w:r>"
        '<w:r><w:object><o:OLEObject r:id="rId1" ProgID="Excel.Sheet.12"/></w:object></w:r>'
        "<w:r><w:t>After object</w:t></w:r>"
        "</w:p></w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" '
        'Target="embeddings/inline.xlsx"/>'
        "</Relationships>"
    )
    docx_path = tmp_path / "inline.docx"
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/embeddings/inline.xlsx", xlsx)

    extraction = extract_embedded_xlsx(docx_path, tmp_path / "extract")
    parsed = [
        parse_xlsx_package(extraction.objects[0].path, tmp_path / "parsed" / "attachment_01", "attachment_01")
    ]
    package = merge_parent_with_attachments(docx_path, extraction, parsed, tmp_path / "package")
    content = package.content_path.read_text(encoding="utf-8")

    before_index = content.index("Before object")
    placeholder_index = content.index("[inline.xlsx](attachments/attachment_01/content.md)")
    after_index = content.index("After object")
    assert before_index < placeholder_index < after_index
