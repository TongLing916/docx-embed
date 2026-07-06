from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
OFFICE_NS = "urn:schemas-microsoft-com:office:office"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def make_xlsx_bytes(rows: list[list[object]], sheet_name: str = "Summary") -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    for row in rows:
        sheet.append(row)

    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def make_docx_with_embeddings(
    path: Path,
    embeddings: list[tuple[str, bytes]],
    intro: str = "Quarterly report",
    outro: str = "End of report",
) -> Path:
    paragraphs = [
        f'<w:p><w:r><w:t>{intro}</w:t></w:r></w:p>',
    ]
    relationships: list[str] = []

    for index, (filename, _) in enumerate(embeddings, start=1):
        rel_id = f"rId{index}"
        relationships.append(
            f'<Relationship Id="{rel_id}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" '
            f'Target="embeddings/{filename}"/>'
        )
        paragraphs.append(
            "<w:p><w:r><w:object>"
            f'<o:OLEObject r:id="{rel_id}" ProgID="Excel.Sheet.12"/>'
            "</w:object></w:r></w:p>"
        )

    paragraphs.append(f'<w:p><w:r><w:t>{outro}</w:t></w:r></w:p>')

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" xmlns:o="{OFFICE_NS}">'
        "<w:body>"
        + "".join(paragraphs)
        + "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        + "".join(relationships)
        + "</Relationships>"
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
        for filename, payload in embeddings:
            docx.writestr(f"word/embeddings/{filename}", payload)

    return path


@pytest.fixture
def embedded_xlsx_docx(tmp_path: Path) -> Path:
    xlsx = make_xlsx_bytes(
        [
            ["Metric", "Value"],
            ["Revenue", 1200],
            ["Profit", 320],
        ],
        sheet_name="Q3",
    )
    return make_docx_with_embeddings(tmp_path / "report.docx", [("profit.xlsx", xlsx)])
