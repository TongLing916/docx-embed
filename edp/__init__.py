"""Embedded Document Parser (EDP) — extract, parse, and reattach embedded assets.

Subpackage layout:

* ``edp.common`` — shared utilities (Markdown normalizer, etc.)
* ``edp.xlsx`` — XLSX workbook parsing and asset extraction
* Top-level — DOCX-specific modules (extractor, merger, parsers)
"""

from __future__ import annotations

from edp.common.markdown_normalizer import (
    normalize_markdown_cell_text,
    normalize_markdown_for_reading,
)
from edp.extractor import extract_document_assets, extract_embedded_xlsx
from edp.main_parser import MainParseResult, parse_main_document
from edp.manifest_builder import build_manifest
from edp.merger import merge_parent_with_attachments
from edp.models import (
    DocumentPackage,
    EmbeddedObject,
    ExtractionResult,
    ParsedPackage,
    ParsedTable,
)
from edp.recursive_parser import parse_attachment_package
from edp.resource_preview import build_resource_previews
from edp.xlsx.assets import XlsxAssetCollection, extract_xlsx_assets
from edp.xlsx.parser import parse_xlsx_package

__all__ = [
    "DocumentPackage",
    "EmbeddedObject",
    "ExtractionResult",
    "MainParseResult",
    "ParsedPackage",
    "ParsedTable",
    "XlsxAssetCollection",
    "build_manifest",
    "build_resource_previews",
    "extract_document_assets",
    "extract_embedded_xlsx",
    "extract_xlsx_assets",
    "merge_parent_with_attachments",
    "normalize_markdown_cell_text",
    "normalize_markdown_for_reading",
    "parse_attachment_package",
    "parse_main_document",
    "parse_xlsx_package",
]
