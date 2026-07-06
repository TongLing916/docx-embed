from __future__ import annotations

import json
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

import examples.run_pipeline as run_pipeline_module
from examples.run_pipeline import main, run_pipeline
from tests.conftest import make_docx_with_embeddings


def test_ragflow_markdown_is_not_a_public_pipeline_parser() -> None:
    assert "ragflow" in run_pipeline_module.MAIN_PARSERS
    assert "ragflow-markdown" not in run_pipeline_module.MAIN_PARSERS


def test_cli_preserves_partial_package_when_xlsx_parse_fails(tmp_path: Path) -> None:
    malformed_xlsx = tmp_path / "malformed.xlsx"
    with ZipFile(malformed_xlsx, "w", ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", "<Types/>")
        workbook.writestr("xl/workbook.xml", "<not-a-workbook/>")

    docx = make_docx_with_embeddings(
        tmp_path / "malformed.docx",
        [("malformed.xlsx", malformed_xlsx.read_bytes())],
    )

    exit_code = run_pipeline(docx, tmp_path / "package")

    assert exit_code == 2
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["parse_status"] == "partial"
    assert "Failed to preview malformed.xlsx" in manifest["parse_warnings"][0]
    assert manifest["content_map"]["embedded_resources"] == "structured/embedded_resources.jsonl"
    assert not (
        tmp_path / "package" / "structured" / "attachments" / "attachment_01" / "content.md"
    ).exists()


def test_pipeline_accepts_docling_as_main_parser(tmp_path: Path, monkeypatch) -> None:
    docx = make_docx_with_embeddings(tmp_path / "docling.docx", [])

    def fake_parse(input_path, parser_name, artifact_dir):
        assert parser_name == "docling"
        artifact_dir.mkdir(parents=True, exist_ok=True)
        artifact = artifact_dir / "main.md"
        artifact.write_text("# Docling Parent\n", encoding="utf-8")
        return type(
            "Result",
            (),
            {
                "parser": "docling",
                "markdown": "# Docling Parent\n",
                "artifacts": {"clean_markdown": artifact},
                "warnings": [],
            },
        )()

    monkeypatch.setattr("edp.merger.parse_main_document", fake_parse)

    exit_code = run_pipeline(docx, tmp_path / "package", main_parser="docling")

    assert exit_code == 0
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["content_map"]["main_parser"] == "docling"
    assert manifest["content_map"]["main_parser_artifacts"] == "structured/_meta/parsers/docling/"


def test_pipeline_accepts_mineru_backend_variant(tmp_path: Path, monkeypatch) -> None:
    docx = make_docx_with_embeddings(tmp_path / "mineru-vlm.docx", [])

    def fake_parse(input_path, parser_name, artifact_dir):
        assert parser_name == "mineru-vlm-engine"
        assert artifact_dir == tmp_path / "package" / "structured" / "_meta" / "parsers" / "mineru-vlm-engine"
        artifact_dir.mkdir(parents=True, exist_ok=True)
        artifact = artifact_dir / "main.md"
        artifact.write_text("# MinerU VLM Parent\n", encoding="utf-8")
        return type(
            "Result",
            (),
            {
                "parser": "mineru-vlm-engine",
                "markdown": "# MinerU VLM Parent\n",
                "artifacts": {"clean_markdown": artifact},
                "warnings": [],
            },
        )()

    monkeypatch.setattr("edp.merger.parse_main_document", fake_parse)

    exit_code = run_pipeline(docx, tmp_path / "package", main_parser="mineru-vlm-engine")

    assert exit_code == 0
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["content_map"]["main_parser"] == "mineru-vlm-engine"
    assert manifest["content_map"]["main_parser_artifacts"] == "structured/_meta/parsers/mineru-vlm-engine/"


def test_pipeline_accepts_pandoc_as_main_parser(tmp_path: Path, monkeypatch) -> None:
    docx = make_docx_with_embeddings(tmp_path / "pandoc.docx", [])

    def fake_parse(input_path, parser_name, artifact_dir):
        assert parser_name == "pandoc"
        artifact_dir.mkdir(parents=True, exist_ok=True)
        artifact = artifact_dir / "main.md"
        artifact.write_text("# Pandoc Parent\n", encoding="utf-8")
        return type(
            "Result",
            (),
            {
                "parser": "pandoc",
                "markdown": "# Pandoc Parent\n",
                "artifacts": {"clean_markdown": artifact},
                "warnings": [],
            },
        )()

    monkeypatch.setattr("edp.merger.parse_main_document", fake_parse)

    exit_code = run_pipeline(docx, tmp_path / "package", main_parser="pandoc")

    assert exit_code == 0
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["content_map"]["main_parser"] == "pandoc"
    assert manifest["content_map"]["main_parser_artifacts"] == "structured/_meta/parsers/pandoc/"
    assert (tmp_path / "package" / "structured" / "content.md").read_text(encoding="utf-8").startswith(
        "# Pandoc Parent"
    )


def test_pipeline_accepts_ragflow_as_main_parser(tmp_path: Path, monkeypatch) -> None:
    docx = make_docx_with_embeddings(tmp_path / "ragflow.docx", [])

    def fake_parse(input_path, parser_name, artifact_dir):
        assert parser_name == "ragflow"
        artifact_dir.mkdir(parents=True, exist_ok=True)
        artifact = artifact_dir / "main.md"
        artifact.write_text("# RAGFlow Parent\n", encoding="utf-8")
        return type(
            "Result",
            (),
            {
                "parser": "ragflow",
                "markdown": "# RAGFlow Parent\n",
                "artifacts": {"clean_markdown": artifact},
                "warnings": [],
            },
        )()

    monkeypatch.setattr("edp.merger.parse_main_document", fake_parse)

    exit_code = run_pipeline(docx, tmp_path / "package", main_parser="ragflow")

    assert exit_code == 0
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["content_map"]["main_parser"] == "ragflow"
    assert manifest["content_map"]["main_parser_artifacts"] == "structured/_meta/parsers/ragflow/"


def test_pipeline_default_extraction_finds_object_directory_assets(tmp_path: Path) -> None:
    docx = make_docx_with_embeddings(tmp_path / "object-dir.docx", [])
    with ZipFile(docx, "a", ZIP_DEFLATED) as archive:
        archive.writestr("word/objects/report.pdf", b"%PDF-1.7\n")

    exit_code = run_pipeline(docx, tmp_path / "package")

    assert exit_code == 0
    manifest = json.loads((tmp_path / "package" / "manifest.json").read_text(encoding="utf-8"))
    assert "asset_strategy" not in manifest["content_map"]
    assert manifest["content_map"]["embedded_objects"][0]["filename"] == "report.pdf"


def test_pipeline_passes_unsafe_unwrap_to_document_extraction(tmp_path: Path, monkeypatch) -> None:
    docx = make_docx_with_embeddings(tmp_path / "unsafe.docx", [])
    seen: dict[str, object] = {}

    def fake_extract(input_docx, work_dir, *, unsafe_unwrap=False):
        seen["input_docx"] = input_docx
        seen["unsafe_unwrap"] = unsafe_unwrap
        from edp.models import ExtractionResult

        work_dir.mkdir(parents=True, exist_ok=True)
        raw_dir = work_dir / "raw"
        raw_dir.mkdir(parents=True, exist_ok=True)
        return ExtractionResult(
            input_docx=input_docx,
            work_dir=work_dir,
            raw_original_path=raw_dir / "original.docx",
            embedded_dir=raw_dir / "embedded",
            media_dir=raw_dir / "media",
        )

    monkeypatch.setattr("examples.run_pipeline.extract_document_assets", fake_extract)

    exit_code = run_pipeline(docx, tmp_path / "package", unsafe_unwrap_embedded=True)

    assert exit_code == 0
    assert seen == {"input_docx": docx, "unsafe_unwrap": True}


def test_pipeline_cli_accepts_unsafe_unwrap_flag(tmp_path: Path, monkeypatch) -> None:
    docx = make_docx_with_embeddings(tmp_path / "unsafe-cli.docx", [])
    seen: dict[str, object] = {}

    def fake_run_pipeline(input_docx, output_dir, *, main_parser="markitdown", unsafe_unwrap_embedded=False):
        seen["input_docx"] = input_docx
        seen["output_dir"] = output_dir
        seen["main_parser"] = main_parser
        seen["unsafe_unwrap_embedded"] = unsafe_unwrap_embedded
        return 0

    monkeypatch.setattr("examples.run_pipeline.run_pipeline", fake_run_pipeline)

    exit_code = main([str(docx), str(tmp_path / "package"), "--unsafe-unwrap-embedded"])

    assert exit_code == 0
    assert seen == {
        "input_docx": docx,
        "output_dir": tmp_path / "package",
        "main_parser": "markitdown",
        "unsafe_unwrap_embedded": True,
    }


def test_pipeline_cli_rejects_removed_asset_strategy_flag(tmp_path: Path) -> None:
    docx = make_docx_with_embeddings(tmp_path / "removed-flag.docx", [])

    with pytest.raises(SystemExit) as exc:
        main([str(docx), str(tmp_path / "package"), "--asset-strategy", "ragflow-like"])

    assert exc.value.code == 2
