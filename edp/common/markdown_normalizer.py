"""Markdown normalization utilities for consistent rendering across parsers.

These helpers are shared between the DOCX and XLSX code paths: XLSX cell text
can contain embedded newlines (wrap-text) and extra whitespace that must be
normalized without breaking the GFM table structure. The DOCX merger uses
``normalize_markdown_for_reading`` to clean presentation whitespace.

Portability rule:
  - ``normalize_markdown_cell_text`` converts ``\\n`` → ``<br>`` so that a
    single logical table row stays on one physical Markdown line.
  - ``normalize_markdown_for_reading`` collapses repeated blank lines and
    trims trailing whitespace without touching fenced code blocks.
"""

from __future__ import annotations

import re


_FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})")
_LIST_ITEM_RE = re.compile(
    r"^(?P<indent>[ \t]*)(?P<marker>[-+*]|\d+[.)])[ \t]+(?P<body>.*)$"
)


def normalize_markdown_for_reading(markdown: str) -> str:
    """Normalize presentation whitespace without changing Markdown structure.

    Collapses repeated blank lines into a single blank line, trims trailing
    whitespace on each line, and preserves fenced code blocks and YAML
    front-matter as-is.
    """

    if not markdown:
        return markdown

    lines = markdown.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    normalized: list[str] = []
    in_fenced_code = False
    in_front_matter = lines and lines[0] == "---"
    blank_pending = False

    for index, line in enumerate(lines):
        if in_front_matter:
            normalized.append(line)
            if index and line in {"---", "..."}:
                in_front_matter = False
            continue

        fence = _FENCE_RE.match(line)
        if in_fenced_code or fence:
            normalized.append(line)
            if fence:
                in_fenced_code = not in_fenced_code
            blank_pending = False
            continue

        cleaned = _normalize_markdown_line(line)
        if not cleaned:
            if normalized and not blank_pending:
                normalized.append("")
                blank_pending = True
            continue

        normalized.append(cleaned)
        blank_pending = False

    while normalized and not normalized[-1]:
        normalized.pop()
    return "\n".join(normalized) + "\n"


def normalize_markdown_cell_text(value: str) -> str:
    """Normalize table-cell text: collapse whitespace, keep logical lines as ``<br>``.

    A markdown table cell that contains a literal ``\\n`` would split the table
    row across multiple physical lines and shatter the GFM table. This function
    converts ``\\n`` → ``<br>`` so the row stays on one line while preserving
    the intent of multi-line cell content. Leading/trailing whitespace within
    each logical line is collapsed.
    """

    logical_lines = value.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    return "<br>".join(
        _collapse_inline_whitespace(line).strip() for line in logical_lines
    )


def _normalize_markdown_line(line: str) -> str:
    """Strip trailing whitespace and collapse inline runs of whitespace."""

    if line.startswith("    "):
        return line.rstrip()
    list_item = _LIST_ITEM_RE.match(line)
    if list_item:
        return (
            f"{list_item.group('indent')}{list_item.group('marker')} "
            f"{_collapse_inline_whitespace(list_item.group('body')).strip()}"
        )
    return _collapse_inline_whitespace(line).strip()


def _collapse_inline_whitespace(value: str) -> str:
    """Replace any run of whitespace characters with a single space."""

    return re.sub(r"\s+", " ", value)
