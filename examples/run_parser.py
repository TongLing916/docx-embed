from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from edp import DocumentPackage, build_manifest, parse_main_document
from edp.path_utils import is_external_or_absolute_path


MINERU_PARSERS = {"mineru", "mineru-pipeline", "mineru-vlm-engine", "mineru-hybrid-engine"}
PARSERS = {
    "markitdown",
    *MINERU_PARSERS,
    "docling",
    "pandoc",
    "ragflow",
}


def run_parser(input_docx: Path, output_dir: Path, parser_name: str) -> int:
    if parser_name not in PARSERS:
        raise ValueError(f"Unsupported parser: {parser_name}")

    if parser_name == "pandoc":
        package = _run_pure_pandoc(input_docx, output_dir)
    else:
        package = _run_pure_main_parser(input_docx, output_dir, parser_name)
    print(f"Package written to: {package.package_dir}")
    print(f"Manifest: {package.manifest_path}")
    if package.warnings:
        print("Warnings:")
        for warning in package.warnings:
            print(f"- {warning}")
        return 2
    return 0


def _run_pure_main_parser(input_docx: Path, output_dir: Path, parser_name: str) -> DocumentPackage:
    package_dir = Path(output_dir)
    raw_dir = package_dir / "raw"
    structured_dir = package_dir / "structured"
    meta_dir = structured_dir / "_meta"
    parser_dir = meta_dir / "parsers" / parser_name
    raw_dir.mkdir(parents=True, exist_ok=True)
    structured_dir.mkdir(parents=True, exist_ok=True)
    meta_dir.mkdir(parents=True, exist_ok=True)

    original_path = raw_dir / "original.docx"
    if input_docx.resolve() != original_path.resolve():
        shutil.copy2(input_docx, original_path)

    result = parse_main_document(input_docx, parser_name, parser_dir)
    return _assemble_pure_package(
        package_dir, result.parser, result.markdown, result.artifacts, result.warnings
    )


def _run_pure_pandoc(input_docx: Path, output_dir: Path) -> DocumentPackage:
    """Run ``pandoc input.docx -o main.md --wrap=preserve`` as a pure baseline.

    Pandoc converts DOCX straight to Markdown via its own reader (no layout
    model), so it is a useful structural baseline alongside MarkItDown/MinerU/
    Docling. Images are extracted with ``--extract-media`` and rewired through
    the same asset pipeline as the other pure parsers.
    """
    package_dir = Path(output_dir)
    raw_dir = package_dir / "raw"
    structured_dir = package_dir / "structured"
    meta_dir = structured_dir / "_meta"
    parser_dir = meta_dir / "parsers" / "pandoc"
    raw_dir.mkdir(parents=True, exist_ok=True)
    structured_dir.mkdir(parents=True, exist_ok=True)
    meta_dir.mkdir(parents=True, exist_ok=True)
    parser_dir.mkdir(parents=True, exist_ok=True)

    original_path = raw_dir / "original.docx"
    if input_docx.resolve() != original_path.resolve():
        shutil.copy2(input_docx, original_path)

    markdown_path = parser_dir / "main.md"
    media_dir = parser_dir / "media"
    markdown, warnings = _invoke_pandoc(input_docx, markdown_path, media_dir)
    artifacts = {"clean_markdown": markdown_path} if markdown_path.exists() else {}
    return _assemble_pure_package(
        package_dir, "pandoc", markdown, artifacts, warnings
    )


def _invoke_pandoc(
    input_docx: Path, output_md: Path, media_dir: Path
) -> tuple[str, list[str]]:
    """Shell out to ``pandoc``; return ``(markdown, warnings)``.

    Isolated so tests can monkeypatch the subprocess without needing the
    ``pandoc`` binary installed. Runs from the markdown file's directory with
    ``--extract-media=.`` so pandoc emits image links relative to that
    directory (and the markdown) rather than as CWD-relative or absolute
    paths — pandoc preserves the docx internal ``media/`` subpath, so the
    extracted files land in ``media_dir`` (``<cwd>/media``) and the
    downstream asset rewriter can resolve them.
    """
    if not shutil.which("pandoc"):
        return "", ["pandoc is not installed"]
    cmd = [
        "pandoc",
        str(Path(input_docx).resolve()),
        "-o",
        output_md.name,
        "--wrap=preserve",
        "--extract-media=.",
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, cwd=str(output_md.parent))
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()[:1000]
        return "", [f"pandoc failed (exit {proc.returncode}): {detail}"]
    markdown = output_md.read_text(encoding="utf-8") if output_md.exists() else ""
    warnings: list[str] = []
    if proc.stderr.strip():
        warnings.append(f"pandoc stderr: {proc.stderr.strip()[:1000]}")
    return markdown, warnings


def _assemble_pure_package(
    package_dir: Path,
    parser_name: str,
    markdown: str,
    artifacts: dict[str, Any],
    warnings: list[str],
) -> DocumentPackage:
    structured_dir = package_dir / "structured"
    meta_dir = structured_dir / "_meta"
    content_path = structured_dir / "content.md"
    markdown = _materialize_parser_assets(markdown, artifacts, package_dir, parser_name)
    content_path.write_text(markdown, encoding="utf-8")

    warnings = list(warnings)
    status = "success" if not warnings else "partial"
    if not markdown:
        status = "failed"

    content_map = {
        "main": "structured/content.md",
        "parser": parser_name,
        "parser_artifacts": _relative_artifacts(artifacts, package_dir),
    }
    manifest = build_manifest(package_dir, status, warnings, content_map)
    manifest_path = package_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (meta_dir / "parse_log.txt").write_text(_render_parse_log(status, warnings), encoding="utf-8")

    return DocumentPackage(
        package_dir=package_dir,
        content_path=content_path,
        manifest_path=manifest_path,
        warnings=warnings,
    )


def _relative_artifacts(artifacts: dict[str, Any], package_dir: Path) -> dict[str, str]:
    return {key: _relative_path(Path(value), package_dir) for key, value in artifacts.items()}


def _materialize_parser_assets(
    markdown: str, artifacts: dict[str, Any], package_dir: Path, parser_name: str
) -> str:
    clean_markdown = artifacts.get("clean_markdown")
    if clean_markdown is None:
        return markdown
    markdown_dir = Path(clean_markdown).parent
    structured_dir = package_dir / "structured"
    assets_dir = structured_dir / "assets" / "images" / parser_name

    def rewrite_asset_path(raw_path: str) -> str:
        if is_external_or_absolute_path(raw_path):
            return raw_path
        source = (markdown_dir / raw_path).resolve()
        if not source.exists() or not source.is_file():
            return raw_path
        destination = assets_dir / Path(raw_path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
        return Path(os.path.relpath(destination, structured_dir)).as_posix()

    markdown = re.sub(
        r"!\[([^\]]*)\]\(([^)]+)\)",
        lambda match: f"![{match.group(1)}]({rewrite_asset_path(match.group(2))})",
        markdown,
    )
    return re.sub(
        r'(<img\b[^>]*\bsrc=")([^"]+)(")',
        lambda match: f"{match.group(1)}{rewrite_asset_path(match.group(2))}{match.group(3)}",
        markdown,
    )


def _relative_path(path: Path, package_dir: Path) -> str:
    try:
        return path.resolve().relative_to(package_dir.resolve()).as_posix()
    except ValueError:
        return path.as_posix()


def _render_parse_log(status: str, warnings: list[str]) -> str:
    lines = [f"status: {status}"]
    if warnings:
        lines.append("warnings:")
        lines.extend(f"- {warning}" for warning in warnings)
    else:
        lines.append("warnings: []")
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run a pure parser comparison package.")
    parser.add_argument("parser", choices=sorted(PARSERS))
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("output_dir", type=Path)
    args = parser.parse_args(argv)
    return run_parser(args.input_docx, args.output_dir, args.parser)


if __name__ == "__main__":
    raise SystemExit(main())
