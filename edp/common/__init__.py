"""Common utilities shared across document format parsers (DOCX, XLSX, etc.)."""

from edp.common.markdown_normalizer import (
    normalize_markdown_cell_text,
    normalize_markdown_for_reading,
)

__all__ = [
    "normalize_markdown_cell_text",
    "normalize_markdown_for_reading",
]
