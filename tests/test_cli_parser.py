from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

from examples.run_parser import run_parser
from tests.conftest import make_docx_with_embeddings
import examples.run_parser as run_parser_module


def test_ragflow_markdown_is_not_a_public_pure_parser() -> None:
    assert "ragflow" in run_parser_module.PARSERS
    assert "ragflow-markdown" not in run_parser_module.PARSERS


def test_pure_markitdown_mode_writes_parser_markdown_without_attachment_package(
    tmp_path: Path, monkeypatch
) -> None:
    docx = make_docx_with_embeddings(tmp_path / "pure.docx", [])

    def fake_parse(input_path: Path, parser_name: str, artifact_dir: Path):
        assert input_path == docx
        assert parser_name == "markitdown"
        artifact_dir.mkdir(parents=True)
        artifact = artifact_dir / "pure.md"
        artifact.write_text("# Pure MarkItDown\n", encoding="utf-8")
        return SimpleNamespace(
            parser="markitdown",
            markdown="# Pure MarkItDown\n",
            artifacts={"clean_markdown": artifact},
            warnings=[],
        )

    monkeypatch.setattr("examples.run_parser.parse_main_document", fake_parse, raising=False)

    exit_code = run_parser(docx, tmp_path / "package", "markitdown")

    assert exit_code == 0
    assert (tmp_path / "package" / "structured" / "content.md").read_text(encoding="utf-8") == "# Pure MarkItDown\n"
    assert not (tmp_path / "package" / "structured" / "attachments").exists()
    assert not (tmp_path / "package" / "structured" / "child_files.md").exists()

    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["parse_status"] == "success"
    assert manifest["content_map"]["parser"] == "markitdown"
    assert manifest["content_map"]["parser_artifacts"] == {
        "clean_markdown": "structured/_meta/parsers/markitdown/pure.md"
    }


def test_pure_markitdown_mode_copies_parser_images_and_rewrites_links(
    tmp_path: Path, monkeypatch
) -> None:
    docx = make_docx_with_embeddings(tmp_path / "pure-markitdown-images.docx", [])

    def fake_parse(input_path: Path, parser_name: str, artifact_dir: Path):
        assert input_path == docx
        assert parser_name == "markitdown"
        image_dir = artifact_dir / "images"
        image_dir.mkdir(parents=True)
        markdown = "# Pure MarkItDown\n\n![diagram](images/doc_001.png)\n"
        markdown_path = artifact_dir / "pure.md"
        markdown_path.write_text(markdown, encoding="utf-8")
        (image_dir / "doc_001.png").write_bytes(b"png")
        return SimpleNamespace(
            parser="markitdown",
            markdown=markdown,
            artifacts={"clean_markdown": markdown_path},
            warnings=[],
        )

    monkeypatch.setattr("examples.run_parser.parse_main_document", fake_parse, raising=False)

    exit_code = run_parser(docx, tmp_path / "package", "markitdown")

    assert exit_code == 0
    content = (tmp_path / "package" / "structured" / "content.md").read_text(encoding="utf-8")
    assert "![diagram](assets/images/markitdown/images/doc_001.png)" in content
    assert (tmp_path / "package" / "structured" / "assets" / "images" / "markitdown" / "images" / "doc_001.png").read_bytes() == b"png"


def test_pure_mineru_mode_routes_original_docx_to_mineru_parser(
    tmp_path: Path, monkeypatch
) -> None:
    docx = make_docx_with_embeddings(tmp_path / "pure-mineru.docx", [])

    def fake_parse(input_path: Path, parser_name: str, artifact_dir: Path):
        assert input_path == docx
        assert parser_name == "mineru"
        assert artifact_dir == tmp_path / "package" / "structured" / "_meta" / "parsers" / "mineru"
        return SimpleNamespace(
            parser="mineru",
            markdown="# Pure MinerU\n",
            artifacts={},
            warnings=[],
        )

    monkeypatch.setattr("examples.run_parser.parse_main_document", fake_parse, raising=False)

    exit_code = run_parser(docx, tmp_path / "package", "mineru")

    assert exit_code == 0
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["content_map"]["parser"] == "mineru"
    assert (tmp_path / "package" / "structured" / "content.md").read_text(encoding="utf-8") == "# Pure MinerU\n"


def test_pure_mineru_mode_copies_zip_images_and_rewrites_links(
    tmp_path: Path, monkeypatch
) -> None:
    docx = make_docx_with_embeddings(tmp_path / "pure-mineru-images.docx", [])

    def fake_parse(input_path: Path, parser_name: str, artifact_dir: Path):
        assert input_path == docx
        assert parser_name == "mineru"
        markdown_dir = artifact_dir / "zip" / "document" / "office"
        image_dir = markdown_dir / "images"
        image_dir.mkdir(parents=True)
        markdown = (
            "# Pure MinerU\n\n"
            "![chart](images/page_1.png)\n\n"
            '<img src="images/page_2.jpg"/>\n'
        )
        markdown_path = markdown_dir / "document.md"
        markdown_path.write_text(markdown, encoding="utf-8")
        (image_dir / "page_1.png").write_bytes(b"png")
        (image_dir / "page_2.jpg").write_bytes(b"jpg")
        return SimpleNamespace(
            parser="mineru",
            markdown=markdown,
            artifacts={"clean_markdown": markdown_path},
            warnings=[],
        )

    monkeypatch.setattr("examples.run_parser.parse_main_document", fake_parse, raising=False)

    exit_code = run_parser(docx, tmp_path / "package", "mineru")

    assert exit_code == 0
    content = (tmp_path / "package" / "structured" / "content.md").read_text(encoding="utf-8")
    assert "![chart](assets/images/mineru/images/page_1.png)" in content
    assert '<img src="assets/images/mineru/images/page_2.jpg"/>' in content
    assert (tmp_path / "package" / "structured" / "assets" / "images" / "mineru" / "images" / "page_1.png").read_bytes() == b"png"
    assert (tmp_path / "package" / "structured" / "assets" / "images" / "mineru" / "images" / "page_2.jpg").read_bytes() == b"jpg"


def test_pure_mineru_backend_variant_copies_zip_images_under_variant_name(
    tmp_path: Path, monkeypatch
) -> None:
    docx = make_docx_with_embeddings(tmp_path / "pure-mineru-hybrid-images.docx", [])

    def fake_parse(input_path: Path, parser_name: str, artifact_dir: Path):
        assert input_path == docx
        assert parser_name == "mineru-hybrid-engine"
        assert artifact_dir == tmp_path / "package" / "structured" / "_meta" / "parsers" / "mineru-hybrid-engine"
        markdown_dir = artifact_dir / "zip" / "document" / "office"
        image_dir = markdown_dir / "images"
        image_dir.mkdir(parents=True)
        markdown = (
            "# Pure MinerU Hybrid\n\n"
            "![chart](images/page_1.png)\n\n"
            '<img src="images/page_2.jpg"/>\n'
        )
        markdown_path = markdown_dir / "document.md"
        markdown_path.write_text(markdown, encoding="utf-8")
        (image_dir / "page_1.png").write_bytes(b"png")
        (image_dir / "page_2.jpg").write_bytes(b"jpg")
        return SimpleNamespace(
            parser="mineru-hybrid-engine",
            markdown=markdown,
            artifacts={"clean_markdown": markdown_path},
            warnings=[],
        )

    monkeypatch.setattr("examples.run_parser.parse_main_document", fake_parse, raising=False)

    exit_code = run_parser(docx, tmp_path / "package", "mineru-hybrid-engine")

    assert exit_code == 0
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["content_map"]["parser"] == "mineru-hybrid-engine"
    content = (tmp_path / "package" / "structured" / "content.md").read_text(encoding="utf-8")
    assert "![chart](assets/images/mineru-hybrid-engine/images/page_1.png)" in content
    assert '<img src="assets/images/mineru-hybrid-engine/images/page_2.jpg"/>' in content
    assert (
        tmp_path / "package" / "structured" / "assets" / "images" / "mineru-hybrid-engine" / "images" / "page_1.png"
    ).read_bytes() == b"png"
    assert (
        tmp_path / "package" / "structured" / "assets" / "images" / "mineru-hybrid-engine" / "images" / "page_2.jpg"
    ).read_bytes() == b"jpg"


def test_pure_docling_mode_routes_original_docx_to_docling_parser(
    tmp_path: Path, monkeypatch
) -> None:
    docx = make_docx_with_embeddings(tmp_path / "pure-docling.docx", [])

    def fake_parse(input_path: Path, parser_name: str, artifact_dir: Path):
        assert input_path == docx
        assert parser_name == "docling"
        assert artifact_dir == tmp_path / "package" / "structured" / "_meta" / "parsers" / "docling"
        return SimpleNamespace(
            parser="docling",
            markdown="# Pure Docling\n",
            artifacts={},
            warnings=[],
        )

    monkeypatch.setattr("examples.run_parser.parse_main_document", fake_parse, raising=False)

    exit_code = run_parser(docx, tmp_path / "package", "docling")

    assert exit_code == 0
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["content_map"]["parser"] == "docling"
    assert (tmp_path / "package" / "structured" / "content.md").read_text(encoding="utf-8") == "# Pure Docling\n"


def test_pure_pandoc_mode_invokes_pandoc_and_rewrites_extracted_media(
    tmp_path: Path, monkeypatch
) -> None:
    docx = make_docx_with_embeddings(tmp_path / "pure-pandoc.docx", [])

    def fake_invoke(input_docx: Path, output_md: Path, media_dir: Path):
        assert input_docx == docx
        output_md.parent.mkdir(parents=True, exist_ok=True)
        media_dir.mkdir(parents=True, exist_ok=True)
        (media_dir / "image1.png").write_bytes(b"png")
        markdown = "# Pure Pandoc\n\n![diagram](media/image1.png)\n"
        output_md.write_text(markdown, encoding="utf-8")
        return markdown, []

    monkeypatch.setattr(run_parser_module, "_invoke_pandoc", fake_invoke)

    exit_code = run_parser(docx, tmp_path / "package", "pandoc")

    assert exit_code == 0
    content = (tmp_path / "package" / "structured" / "content.md").read_text(encoding="utf-8")
    assert content.startswith("# Pure Pandoc")
    assert "![diagram](assets/images/pandoc/media/image1.png)" in content
    assert (
        tmp_path / "package" / "structured" / "assets" / "images" / "pandoc" / "media" / "image1.png"
    ).read_bytes() == b"png"
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["content_map"]["parser"] == "pandoc"
    assert manifest["parse_status"] == "success"
    assert (tmp_path / "package" / "raw" / "original.docx").exists()


def test_pure_pandoc_mode_warns_when_pandoc_missing(tmp_path: Path, monkeypatch) -> None:
    docx = make_docx_with_embeddings(tmp_path / "pure-pandoc-missing.docx", [])
    monkeypatch.setattr(
        run_parser_module, "_invoke_pandoc", lambda *_a, **_k: ("", ["pandoc is not installed"])
    )

    exit_code = run_parser(docx, tmp_path / "package", "pandoc")

    assert exit_code == 2
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["parse_status"] == "failed"
    assert "pandoc is not installed" in manifest["parse_warnings"][0]


def test_pure_pandoc_invokes_pandoc_with_relative_extract_media(
    tmp_path: Path, monkeypatch
) -> None:
    # Pandoc must run from the markdown directory with ``--extract-media=.`` so
    # image links come out relative to the markdown (not CWD-relative/absolute
    # paths that break when the package is moved or re-read from elsewhere).
    docx = make_docx_with_embeddings(tmp_path / "pure-pandoc-rel.docx", [])

    captured = {}

    def fake_run(cmd, **kwargs):
        captured["cmd"] = cmd
        captured["cwd"] = kwargs.get("cwd")
        # Simulate pandoc: write the markdown and the extracted media file
        # under the cwd's ./media/ folder (pandoc preserves the docx media/
        # subpath).
        cwd = Path(kwargs.get("cwd") or ".")
        (cwd / "media").mkdir(parents=True, exist_ok=True)
        (cwd / "media" / "image1.png").write_bytes(b"png")
        output_name = cmd[cmd.index("-o") + 1]
        (cwd / output_name).write_text(
            "# Pure Pandoc\n\n![diagram](./media/image1.png)\n", encoding="utf-8"
        )
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(run_parser_module.shutil, "which", lambda _: "/usr/local/bin/pandoc")
    monkeypatch.setattr(run_parser_module.subprocess, "run", fake_run)

    exit_code = run_parser(docx, tmp_path / "package", "pandoc")

    assert exit_code == 0
    # Relative extract-media, run from the markdown directory.
    assert "--extract-media=." in captured["cmd"]
    assert captured["cmd"][captured["cmd"].index("-o") + 1] == "main.md"
    assert Path(captured["cwd"]) == (tmp_path / "package" / "structured" / "_meta" / "parsers" / "pandoc")
    # The relative link was rewired into the package's assets tree.
    content = (tmp_path / "package" / "structured" / "content.md").read_text(encoding="utf-8")
    assert "![diagram](assets/images/pandoc/media/image1.png)" in content
    assert (
        tmp_path / "package" / "structured" / "assets" / "images" / "pandoc" / "media" / "image1.png"
    ).read_bytes() == b"png"
