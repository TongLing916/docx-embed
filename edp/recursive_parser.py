"""Dispatch embedded attachments to format-specific parsers.

This module is kept minimal — the real parsing logic lives in subpackages:

* :mod:`edp.xlsx.parser` — XLSX workbook → structured tables
* Future: :mod:`edp.pptx.parser` — PPTX slides → structured content
"""

from __future__ import annotations

from pathlib import Path

from edp.models import EmbeddedObject, ParsedPackage


def parse_attachment_package(
    embedded: EmbeddedObject, output_dir: str | Path
) -> ParsedPackage:
    """Parse supported embedded attachments into structured child packages.

    Currently supports:

    * ``xlsx`` — XLSX workbooks parsed via :func:`edp.xlsx.parser.parse_xlsx_package`.
    """

    if embedded.type == "xlsx":
        from edp.xlsx.parser import parse_xlsx_package

        return parse_xlsx_package(embedded.path, output_dir, embedded.ref)
    raise ValueError(
        f"Unsupported attachment type for parsing: {embedded.type}"
    )
