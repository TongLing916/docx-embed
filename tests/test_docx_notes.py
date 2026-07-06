from __future__ import annotations

from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZIP_DEFLATED, ZipFile

from edp.main_parser import parse_main_document


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def test_markitdown_broken_anchors_are_normalized_to_markdown_footnotes(
    tmp_path: Path, monkeypatch
) -> None:
    source = _write_docx_with_notes(tmp_path / "notes.docx")

    broken_markdown = (
        "Alpha [[1]](#footnote-1) beta.\n\n"
        "Omega [[2]](#endnote-1) done.\n\n"
        "1. Foot note text [↑](#footnote-ref-1)\n"
        "2. End note text [↑](#endnote-ref-1)\n"
    )

    monkeypatch.setattr(
        "edp.main_parser._convert_docx_with_markitdown_images",
        lambda *_args: broken_markdown,
    )

    result = parse_main_document(source, "markitdown", tmp_path / "artifacts")

    assert result.markdown == (
        "Alpha [^1] beta.\n\n"
        "Omega [^2] done.\n\n"
        "[^1]: Foot note text\n\n"
        "[^2]: End note text\n"
    )
    assert (tmp_path / "artifacts" / "notes.md").read_text(encoding="utf-8") == result.markdown


def test_mineru_missing_note_markers_are_restored_from_docx(
    tmp_path: Path, monkeypatch
) -> None:
    source = _write_docx_with_notes(tmp_path / "notes.docx")
    payload = _zip_response("result.md", "Alpha beta.\n\nOmega done.\n")

    class FakeResponse:
        headers = {"Content-Type": "application/zip"}

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def read(self) -> bytes:
            return payload

    monkeypatch.setenv("MINERU_API_KEY", "mineru-token")
    monkeypatch.setattr("edp.main_parser.request.urlopen", lambda *_args, **_kwargs: FakeResponse())

    result = parse_main_document(source, "mineru", tmp_path / "artifacts")

    assert result.markdown == (
        "Alpha [^1] beta.\n\n"
        "Omega [^2] done.\n\n"
        "[^1]: Foot note text\n\n"
        "[^2]: End note text\n"
    )
    clean_markdown = tmp_path / "artifacts" / "zip" / "result.md"
    assert clean_markdown.read_text(encoding="utf-8") == result.markdown


def test_pandoc_markdown_footnotes_are_not_duplicated(tmp_path: Path, monkeypatch) -> None:
    source = _write_docx_with_notes(tmp_path / "notes.docx")

    def fake_run(cmd, capture_output, text, cwd):
        markdown_path = Path(cwd) / "notes.md"
        markdown_path.write_text(
            "Alpha [^1] beta.\n\n"
            "Omega [^2] done.\n\n"
            "[^1]: Foot note text\n\n"
            "[^2]: End note text\n",
            encoding="utf-8",
        )
        return SimpleNamespace(returncode=0, stderr="", stdout="")

    monkeypatch.setattr("edp.main_parser.shutil.which", lambda name: "/usr/bin/pandoc")
    monkeypatch.setattr("edp.main_parser.subprocess.run", fake_run)

    result = parse_main_document(source, "pandoc", tmp_path / "artifacts")

    assert result.markdown.count("[^1]:") == 1
    assert result.markdown.count("[^2]:") == 1
    assert "[[1]](#footnote-1)" not in result.markdown


def _zip_response(markdown_name: str, markdown: str) -> bytes:
    stream = BytesIO()
    with ZipFile(stream, "w", ZIP_DEFLATED) as archive:
        archive.writestr(markdown_name, markdown)
    return stream.getvalue()


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
