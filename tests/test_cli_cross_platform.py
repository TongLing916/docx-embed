from __future__ import annotations

import json
from pathlib import Path

import edp.cli as cli
from tests.conftest import make_docx_with_embeddings


def test_cli_pipeline_subcommand_runs_pipeline(tmp_path: Path, monkeypatch) -> None:
    docx = make_docx_with_embeddings(tmp_path / "source.docx", [])
    seen: dict[str, object] = {}

    def fake_run_pipeline(
        input_docx,
        output_dir,
        *,
        main_parser="markitdown",
        unsafe_unwrap_embedded=False,
    ):
        seen["input_docx"] = input_docx
        seen["output_dir"] = output_dir
        seen["main_parser"] = main_parser
        seen["unsafe_unwrap_embedded"] = unsafe_unwrap_embedded
        return 0

    monkeypatch.setattr(cli, "run_pipeline", fake_run_pipeline)

    exit_code = cli.main(
        [
            "pipeline",
            str(docx),
            str(tmp_path / "package"),
            "--main-parser",
            "ragflow",
            "--unsafe-unwrap-embedded",
        ]
    )

    assert exit_code == 0
    assert seen == {
        "input_docx": docx,
        "output_dir": tmp_path / "package",
        "main_parser": "ragflow",
        "unsafe_unwrap_embedded": True,
    }


def test_cli_parser_subcommand_runs_parser(tmp_path: Path, monkeypatch) -> None:
    docx = make_docx_with_embeddings(tmp_path / "source.docx", [])
    seen: dict[str, object] = {}

    def fake_run_parser(input_docx, output_dir, parser_name):
        seen["input_docx"] = input_docx
        seen["output_dir"] = output_dir
        seen["parser_name"] = parser_name
        return 2

    monkeypatch.setattr(cli, "run_parser", fake_run_parser)

    exit_code = cli.main(["parser", "pandoc", str(docx), str(tmp_path / "package")])

    assert exit_code == 2
    assert seen == {
        "input_docx": docx,
        "output_dir": tmp_path / "package",
        "parser_name": "pandoc",
    }


def test_cli_batch_recurses_skips_office_lock_files_and_preserves_relative_layout(
    tmp_path: Path, monkeypatch
) -> None:
    source_dir = tmp_path / "input"
    nested_dir = source_dir / "nested"
    nested_dir.mkdir(parents=True)
    docx = make_docx_with_embeddings(nested_dir / "report.docx", [])
    lock_docx = make_docx_with_embeddings(source_dir / "~$report.docx", [])
    (source_dir / "notes.txt").write_text("skip", encoding="utf-8")
    seen: list[tuple[Path, Path, str, bool]] = []

    def fake_run_pipeline(
        input_docx,
        output_dir,
        *,
        main_parser="markitdown",
        unsafe_unwrap_embedded=False,
    ):
        seen.append((input_docx, output_dir, main_parser, unsafe_unwrap_embedded))
        output_dir.mkdir(parents=True, exist_ok=True)
        (output_dir / "manifest.json").write_text(
            json.dumps({"input": input_docx.name}),
            encoding="utf-8",
        )
        return 0

    monkeypatch.setattr(cli, "run_pipeline", fake_run_pipeline)

    exit_code = cli.main(
        [
            "batch",
            str(source_dir),
            str(tmp_path / "packages"),
            "--main-parser",
            "pandoc",
            "--unsafe-unwrap-embedded",
        ]
    )

    assert exit_code == 0
    assert seen == [(docx, tmp_path / "packages" / "nested" / "report", "pandoc", True)]
    assert lock_docx.exists()


def test_cli_batch_returns_warning_status_when_any_document_warns(
    tmp_path: Path, monkeypatch
) -> None:
    source_dir = tmp_path / "input"
    source_dir.mkdir()
    docx = make_docx_with_embeddings(source_dir / "warn.docx", [])

    def fake_run_pipeline(*_args, **_kwargs):
        return 2

    monkeypatch.setattr(cli, "run_pipeline", fake_run_pipeline)

    exit_code = cli.main(["batch", str(source_dir), str(tmp_path / "packages")])

    assert exit_code == 2
    assert docx.exists()
