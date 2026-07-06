from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from edp import extract_embedded_xlsx, merge_parent_with_attachments, parse_xlsx_package


def test_parent_content_contains_reference_pointing_to_attachment(
    embedded_xlsx_docx: Path, tmp_path: Path
) -> None:
    extraction = extract_embedded_xlsx(embedded_xlsx_docx, tmp_path / "extract")
    parsed = [
        parse_xlsx_package(obj.path, tmp_path / "parsed" / obj.ref, obj.ref)
        for obj in extraction.objects
    ]
    package = merge_parent_with_attachments(
        embedded_xlsx_docx,
        extraction,
        parsed,
        tmp_path / "package",
    )

    content = package.content_path.read_text(encoding="utf-8")
    assert "Quarterly report" in content
    assert "[profit.xlsx](attachments/attachment_01/content.md)" in content
    assert "embedded_object:" not in content
    assert "End of report" in content

    child_files = (tmp_path / "package" / "structured" / "child_files.md").read_text(
        encoding="utf-8"
    )
    assert (
        "| attachment_01 | attachment | profit.xlsx | structured/attachments/attachment_01/ | "
        "structured/attachments/attachment_01/content.md | "
        "structured/attachments/attachment_01/tables/table_001.csv |"
    ) in child_files
    assert "## Reference Definitions" not in child_files

    entry_point = tmp_path / "package" / "structured" / "attachments" / "attachment_01" / "content.md"
    table = tmp_path / "package" / "structured" / "attachments" / "attachment_01" / "tables" / "table_001.csv"
    assert entry_point.exists()
    assert table.exists()


def test_corrupt_ole_embedding_warns_without_blocking_direct_xlsx(tmp_path: Path) -> None:
    from tests.conftest import make_docx_with_embeddings, make_xlsx_bytes

    good_xlsx = make_xlsx_bytes([["Name", "Score"], ["alpha", 7]])
    docx = make_docx_with_embeddings(
        tmp_path / "mixed.docx",
        [
            ("broken.bin", b"not an OLE compound file"),
            ("good.xlsx", good_xlsx),
        ],
    )

    extraction = extract_embedded_xlsx(docx, tmp_path / "extract")

    assert [obj.ref for obj in extraction.objects] == ["attachment_01"]
    assert extraction.objects[0].filename == "good.xlsx"
    assert len(extraction.warnings) == 1
    assert "broken.bin" in extraction.warnings[0]


def test_valid_non_xlsx_zip_bin_warns_without_blocking_direct_xlsx(tmp_path: Path) -> None:
    from tests.conftest import make_docx_with_embeddings, make_xlsx_bytes

    fake_zip = BytesIO()
    with ZipFile(fake_zip, "w", ZIP_DEFLATED) as archive:
        archive.writestr("notes.txt", "not a workbook")

    good_xlsx = make_xlsx_bytes([["Name", "Score"], ["beta", 9]])
    docx = make_docx_with_embeddings(
        tmp_path / "fake-zip.docx",
        [
            ("fake.bin", fake_zip.getvalue()),
            ("good.xlsx", good_xlsx),
        ],
    )

    extraction = extract_embedded_xlsx(docx, tmp_path / "extract")

    assert [obj.ref for obj in extraction.objects] == ["attachment_01"]
    assert extraction.objects[0].filename == "good.xlsx"
    assert len(extraction.warnings) == 1
    assert "fake.bin" in extraction.warnings[0]
