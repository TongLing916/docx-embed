from __future__ import annotations

import argparse
from pathlib import Path

from examples.run_parser import PARSERS, run_parser
from examples.run_pipeline import MAIN_PARSERS, run_pipeline


def run_batch(
    input_dir: Path,
    output_dir: Path,
    *,
    main_parser: str = "markitdown",
    unsafe_unwrap_embedded: bool = False,
) -> int:
    statuses: list[int] = []
    for docx in _iter_docx_files(input_dir):
        relative = docx.relative_to(input_dir)
        destination = output_dir / relative.with_suffix("")
        status = run_pipeline(
            docx,
            destination,
            main_parser=main_parser,
            unsafe_unwrap_embedded=unsafe_unwrap_embedded,
        )
        statuses.append(status)
    if any(status not in {0, 2} for status in statuses):
        return 1
    if any(status == 2 for status in statuses):
        return 2
    return 0


def _iter_docx_files(input_dir: Path) -> list[Path]:
    return sorted(
        path
        for path in input_dir.rglob("*.docx")
        if path.is_file() and not path.name.startswith("~$")
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run Embedded Document Parser tools.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    pipeline = subparsers.add_parser("pipeline", help="Convert one DOCX with the EDP pipeline.")
    pipeline.add_argument("input_docx", type=Path)
    pipeline.add_argument("output_dir", type=Path)
    pipeline.add_argument(
        "--main-parser",
        choices=sorted(MAIN_PARSERS),
        default="markitdown",
        help="Parser used for main-document body inside pipeline mode.",
    )
    pipeline.add_argument(
        "--unsafe-unwrap-embedded",
        action="store_true",
        help="Best-effort unwrap embedded OLE .bin payloads before preserving them.",
    )

    parser_cmd = subparsers.add_parser("parser", help="Run a pure parser comparison package.")
    parser_cmd.add_argument("parser", choices=sorted(PARSERS))
    parser_cmd.add_argument("input_docx", type=Path)
    parser_cmd.add_argument("output_dir", type=Path)

    batch = subparsers.add_parser("batch", help="Convert every DOCX under a directory.")
    batch.add_argument("input_dir", type=Path)
    batch.add_argument("output_dir", type=Path)
    batch.add_argument(
        "--main-parser",
        choices=sorted(MAIN_PARSERS),
        default="markitdown",
        help="Parser used for main-document body inside pipeline mode.",
    )
    batch.add_argument(
        "--unsafe-unwrap-embedded",
        action="store_true",
        help="Best-effort unwrap embedded OLE .bin payloads before preserving them.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "pipeline":
        return run_pipeline(
            args.input_docx,
            args.output_dir,
            main_parser=args.main_parser,
            unsafe_unwrap_embedded=args.unsafe_unwrap_embedded,
        )
    if args.command == "parser":
        return run_parser(args.input_docx, args.output_dir, args.parser)
    if args.command == "batch":
        return run_batch(
            args.input_dir,
            args.output_dir,
            main_parser=args.main_parser,
            unsafe_unwrap_embedded=args.unsafe_unwrap_embedded,
        )
    raise RuntimeError(f"Unsupported command: {args.command}")


if __name__ == "__main__":
    raise SystemExit(main())
