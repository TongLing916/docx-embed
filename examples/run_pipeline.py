from __future__ import annotations

import argparse
from pathlib import Path
import sys


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from edp import (
    DocumentPackage,
    extract_document_assets,
    merge_parent_with_attachments,
)


MINERU_PARSERS = {"mineru", "mineru-pipeline", "mineru-vlm-engine", "mineru-hybrid-engine"}
MAIN_PARSERS = {
    "markitdown",
    *MINERU_PARSERS,
    "docling",
    "pandoc",
    "ragflow",
}


def run_pipeline(
    input_docx: Path,
    output_dir: Path,
    *,
    main_parser: str = "markitdown",
    unsafe_unwrap_embedded: bool = False,
) -> int:
    if main_parser not in MAIN_PARSERS:
        raise ValueError(f"Unsupported main parser: {main_parser}")

    work_dir = output_dir / "structured" / "_meta" / "work"
    extraction = extract_document_assets(
        input_docx,
        work_dir / "extract",
        unsafe_unwrap=unsafe_unwrap_embedded,
    )

    package = merge_parent_with_attachments(
        input_docx,
        extraction,
        [],
        output_dir,
        main_parser=main_parser,
    )

    return _print_package_result(package)


def _print_package_result(package: DocumentPackage) -> int:
    print(f"Package written to: {package.package_dir}")
    print(f"Manifest: {package.manifest_path}")
    if package.warnings:
        print("Warnings:")
        for warning in package.warnings:
            print(f"- {warning}")
        return 2
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the EDP V0.1 DOCX embedded-XLSX pipeline.")
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument(
        "--main-parser",
        choices=sorted(MAIN_PARSERS),
        default="markitdown",
        help="Parser used for main-document body inside pipeline mode.",
    )
    parser.add_argument(
        "--unsafe-unwrap-embedded",
        action="store_true",
        help="Best-effort unwrap embedded OLE .bin payloads before preserving them.",
    )
    args = parser.parse_args(argv)
    return run_pipeline(
        args.input_docx,
        args.output_dir,
        main_parser=args.main_parser,
        unsafe_unwrap_embedded=args.unsafe_unwrap_embedded,
    )


if __name__ == "__main__":
    raise SystemExit(main())
