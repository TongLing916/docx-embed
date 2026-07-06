from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from tests.conftest import PKG_REL_NS, REL_NS, WORD_NS
from edp.main_parser import parse_main_document


DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"


def test_markitdown_parser_preserves_docx_images_as_relative_references(
    tmp_path: Path,
) -> None:
    source = _write_docx_with_inline_image(tmp_path / "image.docx")

    result = parse_main_document(source, "markitdown", tmp_path / "artifacts")

    assert result.warnings == []
    assert "Before image" in result.markdown
    assert "After image" in result.markdown
    assert "![Diagram](images/doc_001.png)" in result.markdown
    assert (tmp_path / "artifacts" / "images" / "doc_001.png").read_bytes() == b"png"
    assert result.artifacts["clean_markdown"] == tmp_path / "artifacts" / "image.md"


def test_markitdown_parser_skips_ole_preview_images(tmp_path: Path) -> None:
    source = _write_docx_with_ole_preview_and_image(tmp_path / "ole-preview.docx")

    result = parse_main_document(source, "markitdown", tmp_path / "artifacts")

    assert result.warnings == []
    assert "Before object" in result.markdown
    assert "After object" in result.markdown
    assert "![Diagram](images/doc_001.png)" in result.markdown
    assert "doc_002" not in result.markdown
    assert sorted(p.name for p in (tmp_path / "artifacts" / "images").iterdir()) == [
        "doc_001.png"
    ]


def _write_docx_with_inline_image(path: Path) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" '
        f'xmlns:a="{DRAWING_NS}" xmlns:wp="{WP_NS}" xmlns:pic="{PIC_NS}">'
        "<w:body>"
        "<w:p><w:r><w:t>Before image</w:t></w:r></w:p>"
        "<w:p><w:r><w:drawing><wp:inline>"
        "<wp:docPr id=\"1\" name=\"Diagram\" descr=\"Diagram\"/>"
        "<a:graphic><a:graphicData><pic:pic><pic:blipFill>"
        '<a:blip r:embed="rId1"/>'
        "</pic:blipFill></pic:pic></a:graphicData></a:graphic>"
        "</wp:inline></w:drawing></w:r></w:p>"
        "<w:p><w:r><w:t>After image</w:t></w:r></w:p>"
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
        docx.writestr("word/media/image1.png", b"png")
    return path


def _write_docx_with_ole_preview_and_image(path: Path) -> Path:
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}" '
        f'xmlns:a="{DRAWING_NS}" xmlns:wp="{WP_NS}" xmlns:pic="{PIC_NS}" '
        'xmlns:o="urn:schemas-microsoft-com:office:office">'
        "<w:body>"
        "<w:p><w:r><w:t>Before object</w:t></w:r></w:p>"
        "<w:p><w:r><w:drawing><wp:inline>"
        "<wp:docPr id=\"1\" name=\"Diagram\" descr=\"Diagram\"/>"
        "<a:graphic><a:graphicData><pic:pic><pic:blipFill>"
        '<a:blip r:embed="rIdImage"/>'
        "</pic:blipFill></pic:pic></a:graphicData></a:graphic>"
        "</wp:inline></w:drawing></w:r></w:p>"
        "<w:p><w:r>"
        '<w:object><v:shape xmlns:v="urn:schemas-microsoft-com:vml">'
        '<v:imagedata r:id="rIdPreview"/>'
        "</v:shape>"
        '<o:OLEObject r:id="rIdOle"/>'
        "</w:object>"
        "</w:r></w:p>"
        "<w:p><w:r><w:t>After object</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rIdImage" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/image1.png"/>'
        '<Relationship Id="rIdPreview" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/preview.png"/>'
        '<Relationship Id="rIdOle" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" '
        'Target="embeddings/oleObject1.bin"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>'
        "</Types>"
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/media/image1.png", b"png")
        docx.writestr("word/media/preview.png", b"preview")
        docx.writestr("word/embeddings/oleObject1.bin", b"ole")
    return path
