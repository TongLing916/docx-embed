"""V0.2 structural-fidelity scorers.

Each scorer compares a parser output package against a checklist ground-truth
dict (produced by ``scripts/build_groundtruth.py``) and returns a float in
``[0, 1]`` or ``None`` when the dimension does not apply to that document
(e.g. a doc with no embedded objects skips ``embedded_object_recall``).

``score`` aggregates the applicable dimensions with equal weight into a
``doc_score``. ``score_gold`` adds deeper diff metrics for the curated
full-gold subset.
"""

from __future__ import annotations

import csv
import json
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


DIMENSIONS = (
    "embedded_object_recall",
    "image_recall",
    "table_recall",
    "nested_table_recall",
    "nested_table_asset_recall",
    "heading_recall",
    "heading_tree_score",
    "key_text_hit",
    "checkbox_recall",
    "chart_text_recall",
    "smartart_text_recall",
    "footnote_recall",
    "footnote_anchor_accuracy",
    "hyperlink_recall",
    "toc_recall",
    "textbox_recall",
    "numbering_recall",
    "table_numbering_recall",
    "formula_recall",
    "table_cell_match",
    "asset_anchor_accuracy",
    "embedded_content_hit",
    "markdown_portability",
)

_IMG_MD_RE = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")
_IMG_HTML_RE = re.compile(r'<img\b[^>]*\bsrc="([^"]+)"')
_HEADING_RE = re.compile(r"^\s{0,3}(#{1,6})\s+(.*)$")


@dataclass
class DimensionResult:
    """Score for one dimension plus the list of unmatched expected items."""

    score: float | None
    misses: list[str] = field(default_factory=list)


def score(package_dir: str | Path, gt: dict[str, Any]) -> dict[str, Any]:
    """Score one parser package against its checklist ground truth."""

    diagnosed = diagnose(package_dir, gt)
    results: dict[str, float | None] = {dim: diagnosed[dim].score for dim in DIMENSIONS}
    applicable = [value for value in results.values() if value is not None]
    results["doc_score"] = sum(applicable) / len(applicable) if applicable else 0.0
    return results


def diagnose(package_dir: str | Path, gt: dict[str, Any]) -> dict[str, DimensionResult]:
    """Score one package and also return which expected items were not recovered.

    Each dimension maps to a :class:`DimensionResult` whose ``misses`` lists a
    human-readable identifier for every expected item the package failed to
    recover (e.g. ``"table rows=3 cols=2"``). Empty when the dimension does not
    apply or everything was matched.
    """

    package_dir = Path(package_dir)
    content = _read_content(package_dir)
    manifest = _read_manifest(package_dir)
    expected = gt.get("expected", {})

    return {
        "embedded_object_recall": _embedded_object_recall(manifest, expected),
        "image_recall": _image_recall(content, manifest, expected),
        "table_recall": _table_recall(content, expected),
        "nested_table_recall": _nested_table_recall(content, expected),
        "nested_table_asset_recall": _nested_table_asset_recall(
            _read_full_content(package_dir), expected
        ),
        "heading_recall": _heading_recall(content, expected),
        "heading_tree_score": _sampled_heading_tree_score(content, expected),
        "key_text_hit": _key_text_hit(_read_full_content(package_dir), expected),
        "checkbox_recall": _checkbox_recall(content, expected),
        "chart_text_recall": _chart_text_recall(_read_full_content(package_dir), expected),
        "smartart_text_recall": _smartart_text_recall(_read_full_content(package_dir), expected),
        "footnote_recall": _footnote_recall(content, expected),
        "footnote_anchor_accuracy": _footnote_anchor_accuracy(content, expected),
        "hyperlink_recall": _hyperlink_recall(content, expected),
        "toc_recall": _toc_recall(content, expected),
        "textbox_recall": _textbox_recall(content, expected),
        "numbering_recall": _numbering_recall(content, expected),
        "table_numbering_recall": _table_numbering_recall(content, expected),
        "formula_recall": _formula_recall(content, expected),
        "table_cell_match": _sampled_table_cell_match(content, expected),
        "asset_anchor_accuracy": _asset_anchor_accuracy(content, manifest, expected),
        "embedded_content_hit": _embedded_content_hit(package_dir, expected),
        "markdown_portability": _markdown_portability(package_dir, content, expected),
    }


def score_gold(package_dir: str | Path, gold_dir: str | Path) -> dict[str, float | None]:
    """Deeper diff metrics for the curated full-gold subset."""

    package_dir = Path(package_dir)
    gold_dir = Path(gold_dir)
    content = _read_content(package_dir)
    gold_md = gold_dir / "gold.md"
    gold_text = gold_md.read_text(encoding="utf-8") if gold_md.exists() else ""

    return {
        "token_f1": _token_f1(gold_text, content),
        "table_cell_match": _table_cell_match(content, gold_dir),
        "heading_tree_score": _heading_tree_score(gold_text, content),
    }


def _embedded_object_recall(manifest: dict[str, Any], expected: dict[str, Any]) -> DimensionResult:
    expected_objects = expected.get("embedded_objects", [])
    if not expected_objects:
        return DimensionResult(None)
    actual = len(manifest.get("content_map", {}).get("embedded_objects", []))
    matched = min(actual, len(expected_objects))
    misses = [
        f"{obj.get('ref', obj.get('source_path', '?'))} type={obj.get('type', '?')}"
        for obj in expected_objects[matched:]
    ] if actual < len(expected_objects) else []
    return DimensionResult(matched / len(expected_objects), misses)


def _image_recall(content: str, manifest: dict[str, Any], expected: dict[str, Any]) -> DimensionResult:
    expected_images = expected.get("images", [])
    if not expected_images:
        return DimensionResult(None)
    # An image counts as recovered if it is either referenced inline in the
    # body markdown OR registered in the manifest's images list. Pipeline
    # methods register images in the manifest; pure parsers only inline them.
    inline_count = len(set(_IMG_MD_RE.findall(content) + _IMG_HTML_RE.findall(content)))
    manifest_count = len(manifest.get("content_map", {}).get("images", []))
    found = max(inline_count, manifest_count)
    matched = min(found, len(expected_images))
    misses = [
        f"{img.get('ref', img.get('filename', '?'))} {img.get('filename', '')}".strip()
        for img in expected_images[matched:]
    ] if found < len(expected_images) else []
    return DimensionResult(matched / len(expected_images), misses)


def _table_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    expected_tables = expected.get("tables", [])
    if not expected_tables:
        return DimensionResult(None)
    matches = _match_expected_tables(content, expected_tables)
    matched = sum(1 for _, found_match, _ in matches if found_match)
    misses = [label for _, found_match, label in matches if not found_match]
    return DimensionResult(matched / len(expected_tables), misses)


def _nested_table_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    expected_tables = expected.get("nested_tables", [])
    if not expected_tables:
        return DimensionResult(None)
    parsed_shapes = [shape for shape, _ in _parse_markdown_tables(content)]
    normalized_content = _normalize_text(content)
    matched = 0
    misses: list[str] = []
    for expected_table in expected_tables:
        target = (expected_table["rows"], expected_table["cols"])
        label = f"index={expected_table.get('index', '?')} rows={target[0]} cols={target[1]}"
        shape_found = any(_nested_shape_matches(shape, target) for shape in parsed_shapes)
        texts = [str(text) for text in expected_table.get("texts", [])]
        texts_found = all(
            (needle := _normalize_text(text)) and needle in normalized_content
            for text in texts
        )
        if shape_found and texts_found:
            matched += 1
        else:
            if not shape_found:
                misses.append(label)
            for text in texts:
                needle = _normalize_text(text)
                if not needle or needle not in normalized_content:
                    misses.append(f"{label} text={text[:40]}")
    return DimensionResult(matched / len(expected_tables), misses)


def _nested_table_asset_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    assets = expected.get("nested_table_assets", [])
    if not assets:
        return DimensionResult(None)
    normalized_content = _normalize_text(content)
    has_inline_image = bool(_IMG_MD_RE.search(content) or _IMG_HTML_RE.search(content))
    matched = 0
    misses: list[str] = []
    for asset in assets:
        kind = str(asset.get("kind", "asset"))
        label = f"table={asset.get('table_index', '?')} kind={kind} rel={asset.get('rel_id', '?')}"
        texts = [str(text) for text in asset.get("texts", []) if str(text)]
        if texts:
            missing_texts = [
                text
                for text in texts
                if not (needle := _normalize_text(text)) or needle not in normalized_content
            ]
            if not missing_texts:
                matched += 1
            else:
                misses.extend(f"{label} text={text[:40]}" for text in missing_texts)
        elif kind == "image" and has_inline_image:
            matched += 1
        else:
            misses.append(label)
    return DimensionResult(matched / len(assets), misses)


def _match_expected_tables(
    content: str,
    expected_tables: list[dict[str, Any]],
) -> list[tuple[dict[str, Any], bool, str]]:
    parsed_shapes = [shape for shape, _ in _parse_markdown_tables(content)]
    consumed: list[bool] = [False] * len(parsed_shapes)
    matches: list[tuple[dict[str, Any], bool, str]] = []
    for expected_table in expected_tables:
        target = (expected_table["rows"], expected_table["cols"])
        label = f"index={expected_table.get('index', '?')} rows={target[0]} cols={target[1]}"
        found_match = False
        for index, shape in enumerate(parsed_shapes):
            if consumed[index]:
                continue
            if _shape_matches(shape, target):
                consumed[index] = True
                found_match = True
                break
        matches.append((expected_table, found_match, label))
    return matches


def _shape_matches(shape: tuple[int, int], target: tuple[int, int]) -> bool:
    """Match table shapes with a tolerance for merged-cell column drift.

    Rows must match exactly. Columns may differ by one: merged cells (and the
    differing ways parsers restore a Word grid vs. a logical view) commonly
    shift the column count by one between ground truth and parsed output, so a
    strict ``cols`` equality would mark whole tables as missing.
    """
    rows_match = shape[0] == target[0]
    cols_close = abs(shape[1] - target[1]) <= 1
    return rows_match and cols_close


def _nested_shape_matches(shape: tuple[int, int], target: tuple[int, int]) -> bool:
    """Match nested table shapes with Pandoc grid-table padding tolerance."""

    rows_match = shape[0] == target[0] or shape[0] == target[0] + 1
    cols_close = abs(shape[1] - target[1]) <= 1
    return rows_match and cols_close


def _heading_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    expected_headings = expected.get("headings", [])
    if not expected_headings:
        return DimensionResult(None)
    content_headings = [_normalize(text) for _, text in _content_headings(content)]
    pointer = 0
    misses: list[str] = []
    matched = 0
    for expected_heading in expected_headings:
        needle = _normalize(expected_heading["text"])
        found_match = False
        while pointer < len(content_headings):
            candidate = content_headings[pointer]
            pointer += 1
            if needle and (needle in candidate or candidate in needle):
                matched += 1
                found_match = True
                break
        if not found_match:
            misses.append(f"L{expected_heading.get('level', '?')} {expected_heading['text']}")
    return DimensionResult(matched / len(expected_headings), misses)


def _key_text_hit(content: str, expected: dict[str, Any]) -> DimensionResult:
    key_texts = expected.get("key_texts", [])
    if not key_texts:
        return DimensionResult(None)
    normalized_content = _normalize_text(content)
    content_no_ws = re.sub(r"\s+", "", normalized_content)
    misses: list[str] = []
    matched = 0
    for text in key_texts:
        normalized_text = _normalize_text(text)
        if not normalized_text:
            misses.append(text)
            continue
        # Match either on whitespace-collapsed text, or on whitespace-stripped
        # text, so a key text survives per-framework rendering differences at
        # inline-tag boundaries (e.g. "<sup>2</sup>" -> " 2 " in one framework
        # vs "2" in another).
        if normalized_text in normalized_content or re.sub(r"\s+", "", normalized_text) in content_no_ws:
            matched += 1
        else:
            misses.append(text)
    return DimensionResult(matched / len(key_texts), misses)


def _checkbox_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    checkboxes = expected.get("checkboxes", [])
    if not checkboxes:
        return DimensionResult(None)
    expected_counts = Counter(
        _normalize_checkbox_text(str(item.get("text", "")))
        for item in checkboxes
        if _normalize_checkbox_text(str(item.get("text", "")))
    )
    if not expected_counts:
        return DimensionResult(None)
    normalized_content = _normalize_checkbox_text(content)
    actual_counts = Counter(
        {
            text: len(re.findall(re.escape(text), normalized_content))
            for text in expected_counts
        }
    )
    matched = sum(min(count, actual_counts.get(text, 0)) for text, count in expected_counts.items())
    misses: list[str] = []
    for text, count in expected_counts.items():
        missing = count - actual_counts.get(text, 0)
        misses.extend([text] * max(0, missing))
    return DimensionResult(matched / sum(expected_counts.values()), misses)


def _normalize_checkbox_text(text: str) -> str:
    normalized = _normalize_text(text)
    if not any(marker in normalized for marker in ("□", "■", "☑", "☒", "☐")):
        return ""
    return normalized


def _chart_text_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    charts = expected.get("charts", [])
    if not charts:
        return DimensionResult(None)
    tokens: list[str] = []
    for chart in charts:
        chart_tokens = chart.get("texts")
        if chart_tokens is None:
            chart_tokens = [
                chart.get("title", ""),
                *chart.get("series", []),
                *chart.get("categories", []),
            ]
        tokens.extend(str(token) for token in chart_tokens)
    normalized_tokens = [token for token in (_normalize_text(token) for token in tokens) if token]
    if not normalized_tokens:
        return DimensionResult(None)
    normalized_content = _normalize_text(content)
    matched = 0
    misses: list[str] = []
    for raw_token, token in zip(tokens, (_normalize_text(token) for token in tokens)):
        if not token:
            continue
        if token in normalized_content:
            matched += 1
        else:
            misses.append(str(raw_token))
    return DimensionResult(matched / len(normalized_tokens), misses)


def _smartart_text_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    smartarts = expected.get("smartarts", [])
    if not smartarts:
        return DimensionResult(None)
    normalized_content = _normalize_text(content)
    expected_with_text = [
        smartart for smartart in smartarts if smartart.get("texts")
    ]
    if not expected_with_text:
        return DimensionResult(None)
    matched = 0
    misses: list[str] = []
    for smartart in expected_with_text:
        texts = [str(text) for text in smartart.get("texts", []) if str(text)]
        missing = [
            text
            for text in texts
            if not (needle := _normalize_text(text)) or needle not in normalized_content
        ]
        if not missing:
            matched += 1
        else:
            misses.extend(str(text) for text in missing)
    return DimensionResult(matched / len(expected_with_text), misses)


def _footnote_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    footnotes = expected.get("footnotes", [])
    if not footnotes:
        return DimensionResult(None)
    matched = 0
    misses: list[str] = []
    normalized_content = _normalize_text(content)
    for footnote in footnotes:
        text = str(footnote.get("text", ""))
        if _normalize_text(text) and _normalize_text(text) in normalized_content:
            matched += 1
        else:
            misses.append(f"id={footnote.get('id', '?')} text={text[:40]}")
    return DimensionResult(matched / len(footnotes), misses)


def _footnote_anchor_accuracy(content: str, expected: dict[str, Any]) -> DimensionResult:
    footnotes = expected.get("footnotes", [])
    if not footnotes:
        return DimensionResult(None)
    matched = 0
    misses: list[str] = []
    lowered = content.lower()
    for footnote in footnotes:
        footnote_id = str(footnote.get("id", "")).strip()
        before = str(footnote.get("anchor_before", "")).lower()
        after = str(footnote.get("anchor_after", "")).lower()
        before_pos = lowered.find(before) if before else -1
        after_pos = lowered.find(after, before_pos + len(before)) if before_pos >= 0 and after else -1
        marker_positions = _footnote_marker_positions(content, footnote_id)
        if before_pos >= 0 and after_pos >= 0 and any(before_pos <= pos <= after_pos for pos in marker_positions):
            matched += 1
        else:
            misses.append(f"id={footnote_id or '?'} anchor={footnote.get('anchor_before', '')}...")
    return DimensionResult(matched / len(footnotes), misses)


def _hyperlink_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    hyperlinks = expected.get("hyperlinks", [])
    if not hyperlinks:
        return DimensionResult(None)
    matched = 0
    misses: list[str] = []
    lowered = content.lower()
    for link in hyperlinks:
        url = str(link.get("url", "")).strip()
        # A hyperlink counts as recovered if the URL string survives anywhere
        # in the body markdown -- accepting Markdown ``[anchor](url)``, HTML
        # ``<a href="url">``, or a bare URL without per-format parsing. The bare
        # anchor text without the URL does not count: the point is whether the
        # link target was preserved.
        if url and url.lower() in lowered:
            matched += 1
        else:
            misses.append(f"url={url or '?'} anchor={link.get('anchor', '')}")
    return DimensionResult(matched / len(hyperlinks), misses)


def _textbox_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    textboxes = expected.get("textboxes", [])
    if not textboxes:
        return DimensionResult(None)
    # A text box is a floating ``wps:txbx`` shape whose ``w:txbxContent`` text
    # is outside the body flow and commonly dropped by parsers. A box counts as
    # recovered when its text survives in the body markdown; matching is on
    # normalized text so inline-mark/whitespace rendering differences do not
    # mask a real recovery. Two boxes with identical text are deduped to one
    # expected entry: substring presence cannot distinguish duplicates, and the
    # point is whether the text-box content was extracted at all.
    normalized_content = _normalize_text(content)
    matched = 0
    misses: list[str] = []
    for box in textboxes:
        text = str(box.get("text", ""))
        needle = _normalize_text(text)
        if needle and needle in normalized_content:
            matched += 1
        else:
            misses.append(f"textbox text={text[:50]}")
    return DimensionResult(matched / len(textboxes), misses)


def _numbering_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    items = expected.get("numbered_items", [])
    if not items:
        return DimensionResult(None)
    # Word auto-numbers list items via the numbering part (``<w:numPr>`` or a
    # list paragraph style), so the visible number is not stored in the run
    # text. Parsers that rebuild numbers themselves commonly restart at 1 per
    # list or flatten sub-items into the main sequence, so the rendered number
    # drifts from Word's continued sequence. An item counts as recovered when
    # a single body line begins with the expected number (as a list marker)
    # followed by the item text -- this catches both a dropped number and a
    # wrong number, while ignoring the same text appearing elsewhere without
    # the marker.
    matched = 0
    misses: list[str] = []
    for item in items:
        number = str(item.get("number", "")).strip()
        text = str(item.get("text", ""))
        if not number or not text:
            misses.append(f"numbered item number={number or '?'} text={text[:40]}")
            continue
        pattern = rf"(?m)^\s*{re.escape(number)}\s*[.)\]]\s*{re.escape(text)}"
        if re.search(pattern, content):
            matched += 1
        else:
            misses.append(f"number={number} text={text[:40]}")
    return DimensionResult(matched / len(items), misses)


def _table_numbering_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    items = expected.get("table_numbered_items", [])
    if not items:
        return DimensionResult(None)
    # Word auto-numbers table rows via the numbering part, so the visible
    # number is not stored in run text. Parsers that rebuild numbers must
    # place each number in the same row as its anchor text (the neighbouring
    # cell that all renderings preserve). An item counts as recovered when a
    # single content line contains both the number as a row marker -- ``N.``
    # at line start, after a ``|`` pipe, or inside a ``<td>`` (optionally
    # wrapped in ``<p>``) -- and the anchor text. This is rendering-agnostic
    # across pipe tables, pandoc simple/grid tables, and HTML tables, and
    # catches both a dropped number and a number that drifted to the wrong
    # row. The preceding-boundary anchor also prevents ``1.`` from matching
    # inside ``41.`` or ``11.``, and the same-line anchor requirement
    # disambiguates the remaining collisions.
    matched = 0
    misses: list[str] = []
    raw_lines = content.splitlines()
    norm_lines = [_normalize_text(line) for line in raw_lines]
    for item in items:
        number = str(item.get("number", "")).strip()
        anchor = str(item.get("anchor", ""))
        if not number or not anchor:
            misses.append(f"table_numbering number={number or '?'} anchor={anchor[:40]}")
            continue
        marker = re.compile(
            rf"(?:^|\||<td[^>]*>(?:<p>)?)\s*{re.escape(number)}\s*\."
        )
        anchor_norm = _normalize_text(anchor)
        found = any(
            marker.search(raw) and anchor_norm in norm
            for raw, norm in zip(raw_lines, norm_lines)
        )
        if found:
            matched += 1
        else:
            misses.append(f"table_numbering number={number} anchor={anchor[:40]}")
    return DimensionResult(matched / len(items), misses)


def _formula_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    formulas = expected.get("formulas", [])
    if not formulas:
        return DimensionResult(None)
    # An OMML formula is recovered when its distinctive tokens survive in the
    # body markdown. Parsers render formulas differently (LaTeX ``$$...$$``,
    # Unicode, plain text), so each expected formula lists tokens that should
    # survive any faithful rendering -- semantic operands plus structural
    # markers like ``\frac`` or ``^{up}``. Matching is on normalized text:
    # ``_normalize_text`` strips markdown escapes (``\_`` -> ``_``) but leaves
    # LaTeX commands intact (``\frac`` survives because the normalizer only
    # drops a backslash before punctuation, not before a letter). Score is
    # token-level recall so a partially-recovered formula (fraction kept,
    # super/subscript dropped) earns partial credit and the miss list names
    # exactly which structural pieces were lost.
    normalized_content = _normalize_text(content)
    content_no_ws = re.sub(r"\s+", "", normalized_content)
    total = 0
    matched = 0
    misses: list[str] = []
    for formula in formulas:
        label = str(formula.get("label", formula.get("latex", "")[:30] or "?"))
        for token in formula.get("tokens", []):
            needle = _normalize_text(str(token))
            total += 1
            # Match on whitespace-collapsed text, or whitespace-stripped text,
            # so a token survives per-framework spacing differences (e.g.
            # docling's ``A\times Q\times T`` vs the GT token ``A×Q×T``).
            if needle and (needle in normalized_content or re.sub(r"\s+", "", needle) in content_no_ws):
                matched += 1
            else:
                misses.append(f"{label}: {token}")
    return DimensionResult(matched / total if total else None, misses)


def _toc_recall(content: str, expected: dict[str, Any]) -> DimensionResult:
    toc = expected.get("toc", [])
    if not toc:
        return DimensionResult(None)
    # The TOC region is the block of body text between the "Table of Contents"
    # heading and the next heading line. Restricting matches to this region is
    # what separates "the TOC entry list was extracted" from "the same titles
    # appear later as body headings" -- without it every entry would trivially
    # match the body and the dimension would be meaningless.
    region = _toc_region(content)
    normalized_region = _normalize_text(region) if region else ""
    matched = 0
    misses: list[str] = []
    for entry in toc:
        text = str(entry.get("text", ""))
        needle = _normalize_text(text)
        if needle and needle in normalized_region:
            matched += 1
        else:
            misses.append(f"toc entry={text[:50]}")
    return DimensionResult(matched / len(toc), misses)


def _toc_region(content: str) -> str:
    """Body text between the "Table of Contents" heading and the next heading.

    Recognizes both ATX (``# Table of Contents``) and setext
    (``Table of Contents\\n=====``) headings; the RAGFlow Markdown export renders the TOC
    title as a setext heading, which the ATX-only matcher missed, dropping the
    whole TOC region and zeroing ``toc_recall`` even when every entry was
    present.
    """
    lines = content.splitlines()
    spans = _heading_line_spans(lines)
    toc_span = None
    for start, end, _, text in spans:
        if _normalize(text) == "table of contents":
            toc_span = (start, end)
            break
    if toc_span is None:
        return ""
    region_start = toc_span[1] + 1
    region_end = len(lines)
    for start, _, _, _ in spans:
        if start >= region_start:
            region_end = start
            break
    return "\n".join(lines[region_start:region_end])


def _heading_line_spans(lines: list[str]) -> list[tuple[int, int, int, str]]:
    """Return ``(start_line, end_line, level, text)`` for each heading.

    ATX headings occupy one line; setext headings occupy two (text + rule).
    Mirrors ``_content_headings`` but preserves line positions for region
    math.
    """
    spans: list[tuple[int, int, int, str]] = []
    for index, line in enumerate(lines):
        match = _HEADING_RE.match(line)
        if match:
            spans.append((index, index, len(match.group(1)), match.group(2).strip()))
            continue
        if index == 0:
            continue
        stripped = line.strip()
        previous = lines[index - 1].strip()
        if previous and re.fullmatch(r"=+", stripped):
            spans.append((index - 1, index, 1, previous))
        elif previous and re.fullmatch(r"-+", stripped):
            spans.append((index - 1, index, 2, previous))
    return spans


def _footnote_marker_positions(content: str, footnote_id: str) -> list[int]:
    if not footnote_id:
        return []
    patterns = [
        rf"\[\^{re.escape(footnote_id)}\]",
        rf"\[\[{re.escape(footnote_id)}\]\]\(#(?:footnote|endnote)-{re.escape(footnote_id)}\)",
        rf"#(?:footnote|endnote)-ref-{re.escape(footnote_id)}",
    ]
    positions: list[int] = []
    for pattern in patterns:
        positions.extend(match.start() for match in re.finditer(pattern, content, re.IGNORECASE))
    return sorted(set(positions))


def _sampled_heading_tree_score(content: str, expected: dict[str, Any]) -> DimensionResult:
    expected_headings = expected.get("headings", [])
    if not expected_headings:
        return DimensionResult(None)
    actual = _heading_nodes(_content_headings(content))
    expected_nodes = _heading_nodes(
        [(int(heading.get("level", 1)), str(heading.get("text", ""))) for heading in expected_headings]
    )
    possible = len(expected_nodes) * 3

    # Pass 1: match expected -> actual by text, in order.
    matches: list[tuple[dict[str, Any], dict[str, Any]]] = []
    misses: list[str] = []
    pointer = 0
    for expected_node in expected_nodes:
        expected_text = _normalize(expected_node["text"])
        found_index = None
        for index in range(pointer, len(actual)):
            candidate_text = _normalize(actual[index]["text"])
            if expected_text and (expected_text in candidate_text or candidate_text in expected_text):
                found_index = index
                break
        if found_index is None:
            misses.append(f"L{expected_node['level']} {expected_node['text']}")
            continue
        matches.append((expected_node, actual[found_index]))
        pointer = found_index + 1

    # Per-document uniform level offset (e.g. docling shifts every Word
    # heading down by one). Compare relative structure, not absolute # count.
    offset = _heading_level_offset([(exp["level"], cand["level"]) for exp, cand in matches])

    # Pass 2: award points.
    points = 0
    for expected_node, candidate in matches:
        points += 1  # found, in order
        if candidate["level"] - offset == expected_node["level"]:
            points += 1
        else:
            misses.append(
                f"level {expected_node['text']} expected=L{expected_node['level']} "
                f"actual=L{candidate['level']} offset={offset:+d}"
            )
        if [_normalize(text) for text in candidate["parents"]] == [
            _normalize(text) for text in expected_node["parents"]
        ]:
            points += 1
        else:
            misses.append(f"parent {expected_node['text']}")
    return DimensionResult(points / possible if possible else None, misses)


def _heading_nodes(headings: list[tuple[int, str]]) -> list[dict[str, Any]]:
    stack: list[tuple[int, str]] = []
    nodes: list[dict[str, Any]] = []
    for level, text in headings:
        stack = [(parent_level, parent_text) for parent_level, parent_text in stack if parent_level < level]
        nodes.append({"level": level, "text": text, "parents": [parent_text for _, parent_text in stack]})
        stack.append((level, text))
    return nodes


def _heading_level_offset(pairs: list[tuple[int, int]]) -> int:
    """Per-document uniform level offset: the most frequent
    ``(candidate_level - expected_level)`` over text-matched headings.

    Parsers like docling shift every Word heading down by a constant, which
    preserves the relative tree but breaks strict level equality. Returning
    that constant lets the scorer compare relative structure instead of
    absolute ``#`` count.

    Returns ``0`` (strict comparison) when fewer than 2 matched headings agree
    on an offset, so single-heading docs stay strict and genuine non-uniform
    drift is still penalized. Ties break toward the offset closest to 0.
    """
    if not pairs:
        return 0
    counts = Counter(cand - exp for exp, cand in pairs)
    best_offset, best_count = counts.most_common(1)[0]
    if best_count < 2:
        return 0
    top = [offset for offset, count in counts.items() if count == best_count]
    return min(top, key=lambda o: abs(o))


def _sampled_table_cell_match(content: str, expected: dict[str, Any]) -> DimensionResult:
    checks = expected.get("table_cell_checks", [])
    if not checks:
        return DimensionResult(None)
    parsed_tables = _parse_markdown_tables(content)
    total = 0
    matched = 0
    misses: list[str] = []
    for check in checks:
        table_index = _resolved_parsed_table_index(check, expected, parsed_tables)
        label = check.get("label", f"table_index={table_index}")
        cells = check.get("cells", [])
        if table_index >= len(parsed_tables):
            total += len(cells)
            misses.append(f"{label} missing table")
            continue
        _, parsed_rows = parsed_tables[table_index]
        for cell in cells:
            total += 1
            row = int(cell.get("row", 0))
            col = int(cell.get("col", 0))
            expected_text = _normalize_text(str(cell.get("text", "")))
            actual_text = ""
            if row < len(parsed_rows) and col < len(parsed_rows[row]):
                actual_text = _normalize_text(parsed_rows[row][col])
            if expected_text and (expected_text == actual_text or expected_text in actual_text):
                matched += 1
            else:
                misses.append(f"{label} row={row} col={col} expected={cell.get('text', '')}")
    return DimensionResult(matched / total if total else None, misses)


def _resolved_parsed_table_index(
    check: dict[str, Any],
    expected: dict[str, Any],
    parsed_tables: list[tuple[tuple[int, int], list[list[str]]]],
) -> int:
    if "parsed_table_index" in check:
        return int(check["parsed_table_index"])
    target_index = check.get("table_index")
    expected_tables = expected.get("tables", [])
    if target_index is None or not expected_tables:
        return int(target_index or 0)
    consumed: list[bool] = [False] * len(parsed_tables)
    for expected_table in expected_tables:
        target = (expected_table["rows"], expected_table["cols"])
        found_index = None
        for parsed_index, (shape, _) in enumerate(parsed_tables):
            if consumed[parsed_index]:
                continue
            if _shape_matches(shape, target):
                consumed[parsed_index] = True
                found_index = parsed_index
                break
        if expected_table.get("index") == target_index:
            return found_index if found_index is not None else len(parsed_tables)
    return int(target_index)


def _asset_anchor_accuracy(
    content: str,
    manifest: dict[str, Any],
    expected: dict[str, Any],
) -> DimensionResult:
    anchors = expected.get("asset_anchors", [])
    if not anchors:
        return DimensionResult(None)
    matched = 0
    misses: list[str] = []
    for anchor in anchors:
        ref = str(anchor.get("ref", ""))
        before = str(anchor.get("before", ""))
        after = str(anchor.get("after", ""))
        positions = _asset_positions(content, manifest, ref)
        before_pos = content.find(before) if before else -1
        after_pos = content.find(after, before_pos + len(before)) if before_pos >= 0 and after else -1
        if before_pos >= 0 and after_pos >= 0 and any(before_pos <= pos <= after_pos for pos in positions):
            matched += 1
        else:
            misses.append(f"{ref or '?'} between={before[:30]}...{after[:30]}")
    return DimensionResult(matched / len(anchors), misses)


def _asset_positions(content: str, manifest: dict[str, Any], ref: str) -> list[int]:
    needles = [ref]
    content_map = manifest.get("content_map", {})
    for key in ("images", "embedded_objects"):
        for item in content_map.get(key, []):
            if item.get("ref") == ref:
                needles.extend(str(item.get(field, "")) for field in ("path", "markdown_reference", "filename"))
    positions: list[int] = []
    for needle in needles:
        if not needle:
            continue
        start = 0
        while True:
            pos = content.find(needle, start)
            if pos == -1:
                break
            positions.append(pos)
            start = pos + len(needle)
    return sorted(set(positions))


def _embedded_content_hit(package_dir: Path, expected: dict[str, Any]) -> DimensionResult:
    checks = expected.get("embedded_content", [])
    if not checks:
        return DimensionResult(None)
    matched = 0
    total = 0
    misses: list[str] = []
    for check in checks:
        ref = str(check.get("ref", ""))
        text = _read_attachment_content(package_dir, ref)
        normalized = _normalize_text(text)
        for expected_text in check.get("texts", []):
            total += 1
            needle = _normalize_text(str(expected_text))
            if needle and needle in normalized:
                matched += 1
            else:
                misses.append(f"{ref or '?'} text={str(expected_text)[:40]}")
    return DimensionResult(matched / total if total else None, misses)


def _read_attachment_content(package_dir: Path, ref: str) -> str:
    structured = package_dir / "structured"
    candidates = [
        structured / "attachments" / ref / "content.md",
        structured / "resources" / ref / "preview.md",
        structured / "resources" / ref / "preview.json",
    ]
    return "\n".join(_read_existing_texts(candidates))


def _read_existing_texts(paths: list[Path]) -> list[str]:
    return [path.read_text(encoding="utf-8") for path in paths if path.is_file()]


def _embedded_resource_text_paths(package_dir: Path) -> list[Path]:
    structured = package_dir / "structured"
    paths = [
        *sorted((structured / "attachments").glob("*/content.md")),
        *sorted((structured / "resources").glob("*/preview.md")),
        *sorted((structured / "resources").glob("*/preview.json")),
    ]
    return [path for path in paths if path.is_file()]


def _markdown_portability(package_dir: Path, content: str, expected: dict[str, Any]) -> DimensionResult:
    if not expected.get("markdown_portability"):
        return DimensionResult(None)
    checks = [
        *_portable_image_checks(package_dir, content),
        *_portable_html_checks(content),
        *_portable_heading_checks(content),
    ]
    if not checks:
        return DimensionResult(1.0)
    passed = sum(1 for _, ok in checks if ok)
    misses = [label for label, ok in checks if not ok]
    return DimensionResult(passed / len(checks), misses)


def _portable_image_checks(package_dir: Path, content: str) -> list[tuple[str, bool]]:
    checks: list[tuple[str, bool]] = []
    links = _IMG_MD_RE.findall(content) + _IMG_HTML_RE.findall(content)
    for link in links:
        if link.startswith("data:"):
            checks.append((f"base64 image {link[:24]}", False))
            continue
        if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", link) or Path(link).is_absolute():
            checks.append((f"non-relative image path {link}", False))
            continue
        target = (package_dir / "structured" / link).resolve()
        try:
            target.relative_to(package_dir.resolve())
        except ValueError:
            checks.append((f"image path escapes package {link}", False))
            continue
        checks.append((f"image exists {link}", target.exists()))
    return checks


def _portable_html_checks(content: str) -> list[tuple[str, bool]]:
    checks: list[tuple[str, bool]] = []
    if re.search(r"<table\b", content, re.IGNORECASE):
        checks.append(("html table", False))
    if re.search(r"\bfile://", content, re.IGNORECASE):
        checks.append(("file url", False))
    return checks


def _portable_heading_checks(content: str) -> list[tuple[str, bool]]:
    lines = content.splitlines()
    checks: list[tuple[str, bool]] = []
    for index, line in enumerate(lines[1:], start=1):
        if lines[index - 1].strip() and re.fullmatch(r"(=+|-+)", line.strip()):
            checks.append((f"setext heading {lines[index - 1].strip()[:40]}", False))
    return checks


def _token_f1(gold_text: str, content: str) -> float | None:
    gold_tokens = set(_normalize(gold_text).split())
    content_tokens = set(_normalize(content).split())
    if not gold_tokens and not content_tokens:
        return 1.0
    if not gold_tokens or not content_tokens:
        return 0.0
    common = gold_tokens & content_tokens
    precision = len(common) / len(content_tokens)
    recall = len(common) / len(gold_tokens)
    if precision + recall == 0:
        return 0.0
    return 2 * precision * recall / (precision + recall)


def _table_cell_match(content: str, gold_dir: Path) -> float | None:
    gold_csvs = sorted(gold_dir.glob("gold_tables/*.csv"))
    if not gold_csvs:
        return None
    parsed_tables = _parse_markdown_tables(content)
    ratios: list[float] = []
    for index, gold_csv in enumerate(gold_csvs):
        gold_rows = _read_csv_rows(gold_csv)
        if not gold_rows:
            continue
        if index >= len(parsed_tables):
            ratios.append(0.0)
            continue
        _, parsed_rows = parsed_tables[index]
        ratios.append(_row_match_ratio(gold_rows, parsed_rows))
    return sum(ratios) / len(ratios) if ratios else None


def _heading_tree_score(gold_text: str, content: str) -> float | None:
    gold_levels = [level for level, _ in _content_headings(gold_text)]
    content_levels = [level for level, _ in _content_headings(content)]
    if not gold_levels:
        return None
    # Tolerate a per-document uniform level shift (e.g. docling's +1) so the
    # edit distance reflects structural drift, not a benign constant offset.
    offset = _heading_level_offset(
        [(gold_levels[i], content_levels[i]) for i in range(min(len(gold_levels), len(content_levels)))]
    )
    shifted = [level - offset for level in content_levels]
    distance = _edit_distance(gold_levels, shifted)
    max_len = max(len(gold_levels), len(shifted))
    return 1.0 - distance / max_len if max_len else 1.0


def _read_content(package_dir: Path) -> str:
    content_path = package_dir / "structured" / "content.md"
    return content_path.read_text(encoding="utf-8") if content_path.exists() else ""


def _read_full_content(package_dir: Path) -> str:
    """Main ``content.md`` plus embedded full parses and shallow previews.

    Used for ``key_text_hit`` so a key text may appear in either the parent
    body or any embedded resource content.
    """

    parts = [
        _read_content(package_dir),
        *_read_existing_texts(_embedded_resource_text_paths(package_dir)),
    ]
    return "\n".join(parts)


def _read_manifest(package_dir: Path) -> dict[str, Any]:
    manifest_path = package_dir / "manifest.json"
    if not manifest_path.exists():
        return {}
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def _parse_markdown_tables(content: str) -> list[tuple[tuple[int, int], list[list[str]]]]:
    """Return ``((rows, cols), cells)`` for each table in ``content``.

    Handles all four pandoc-native Markdown table formats plus inline HTML:
    GitHub pipe tables (markitdown/docling), pandoc grid tables (``+---+``
    borders), pandoc simple/multiline tables (``------`` rule lines), and
    inline HTML ``<table>`` blocks (MinerU). ``rows`` excludes
    separator/border rows. ``cols`` is the cell count of the header row.
    """

    return (
        _parse_pipe_tables(content)
        + _parse_embedded_grid_tables(content)
        + _parse_simple_tables(content)
        + _parse_html_tables(content)
    )


def _parse_pipe_tables(content: str) -> list[tuple[tuple[int, int], list[list[str]]]]:
    tables: list[tuple[tuple[int, int], list[list[str]]]] = []
    block: list[str] = []
    for line in content.splitlines():
        if line.lstrip().startswith("|"):
            block.append(line)
        elif block and _is_grid_separator_line(line):
            # Pandoc emits grid tables whose rows are separated by ``+---+``
            # borders. These border lines sit between ``|`` rows of one table;
            # keep them in the block so the table stays whole instead of being
            # split into one-row fragments.
            block.append(line)
        elif block:
            tables.append(_finalize_table_block(block))
            block = []
    if block:
        tables.append(_finalize_table_block(block))
    return tables


def _parse_embedded_grid_tables(content: str) -> list[tuple[tuple[int, int], list[list[str]]]]:
    """Parse pandoc grid tables nested inside an outer grid-table cell.

    Pandoc renders a DOCX table nested inside another table as a grid-table
    fragment embedded in one outer cell, e.g. ``| outer | +---+---+ |`` and
    ``|       | | a | b | |``. The ordinary pipe/grid parser sees only the
    outer table. For nested-table recall, peel off the outer cell prefix and
    parse the inner ``+---+`` / ``| ... |`` block as its own candidate table.
    """

    tables: list[tuple[tuple[int, int], list[list[str]]]] = []
    block: list[str] = []
    for line in content.splitlines():
        fragment = _embedded_grid_fragment(line)
        if fragment is not None and block:
            block.append(fragment)
        elif fragment is not None and _starts_embedded_grid_table(fragment):
            block.append(fragment)
        elif block:
            _append_embedded_grid_table(tables, block)
            block = []
    if block:
        _append_embedded_grid_table(tables, block)
    return tables


def _embedded_grid_fragment(line: str) -> str | None:
    match = re.match(r"^\s*\|[^|]*\|\s*(?P<fragment>[+|].*)$", line)
    if not match:
        return None
    fragment = match.group("fragment").rstrip()
    if _is_grid_separator_line(fragment) or fragment.startswith("|"):
        return fragment
    return None


def _starts_embedded_grid_table(fragment: str) -> bool:
    return _is_grid_separator_line(fragment) and fragment.count("+") >= 3


def _append_embedded_grid_table(
    tables: list[tuple[tuple[int, int], list[list[str]]]],
    block: list[str],
) -> None:
    if any(line.lstrip().startswith("|") for line in block) and any(
        _is_grid_separator_line(line) for line in block
    ):
        tables.append(_finalize_table_block(block))


def _is_grid_separator_line(line: str) -> bool:
    """A pandoc grid-table border line (full or partial).

    Full borders are ``+---+---+`` lines; partial borders — where only some
    columns end, e.g. ``|          |          +--------+`` — start with
    ``|`` but still carry a ``+---`` segment. Both mark logical-row
    boundaries in a grid table, so we recognize any line containing a ``+``
    followed by a run of ``-``/``=``.
    """
    return bool(re.search(r"\+[-=]{2,}", line))


def _parse_simple_tables(content: str) -> list[tuple[tuple[int, int], list[list[str]]]]:
    """Parse pandoc simple and multiline tables.

    These use dashed rule lines (``------``) instead of ``|`` or ``+``. A
    table has a gapped column-rule line (``------ ------``) whose dash-runs
    give the column count; a header row sits directly above it (optional),
    and data rows sit below. Simple tables put one row per line; multiline
    tables separate rows with blank lines (a row may span several physical
    lines). The region is typically wrapped by full-width ``------`` borders.

    We anchor on each gapped column-rule line, read the header above it, and
    collect data below until the closing border (the next rule line). Row
    counting follows the format: blank-separated blocks for multiline, plain
    non-blank lines for simple.
    """
    lines = content.splitlines()
    tables: list[tuple[tuple[int, int], list[list[str]]]] = []
    i = 0
    while i < len(lines):
        runs = _simple_rule_runs(lines[i])
        if runs is None or runs < 2:
            i += 1
            continue
        table, next_i = _build_simple_table(lines, i)
        if table is not None:
            tables.append(table)
            i = max(next_i, i + 1)
        else:
            i += 1
    return tables


def _build_simple_table(
    lines: list[str], col_rule_idx: int
) -> tuple[tuple[tuple[int, int], list[list[str]]], int] | None:
    col_runs = _simple_rule_runs(lines[col_rule_idx]) or 2

    # Header: contiguous non-blank, non-rule lines directly above the column
    # rule. A multiline header may span several physical lines but counts as
    # one logical row. The walk must stop at the first blank line: pandoc
    # multiline tables emit a gapped top rule with the header *below* it and
    # a blank line above the rule (e.g. a figure caption or heading). Walking
    # through blanks would absorb that unrelated text as a phantom header row
    # and inflate the row count by one.
    header_cells: list[list[str]] = []
    j = col_rule_idx - 1
    while j >= 0 and lines[j].strip() and _is_simple_table_text(lines[j]):
        header_cells.append(_split_simple_cells(lines[j], col_runs))
        j -= 1
    header_cells.reverse()
    header_rows = 1 if header_cells else 0

    # Data: lines below the column rule until the closing border (the next
    # rule line). Blank lines separate rows in multiline tables.
    data_lines: list[str] = []
    k = col_rule_idx + 1
    while k < len(lines):
        if _simple_rule_runs(lines[k]) is not None:
            k += 1  # consume the closing border
            break
        if not _is_simple_table_text(lines[k]):
            # A pipe/grid line or other construct ends the table without a
            # closing border.
            break
        data_lines.append(lines[k])
        k += 1

    data_rows = _count_simple_data_rows(data_lines, col_runs)
    cells = header_cells + _simple_data_cells(data_lines, col_runs)
    if data_rows == 0 and not header_cells:
        return None
    return ((header_rows + data_rows, col_runs), cells), k


def _count_simple_data_rows(data_lines: list[str], col_runs: int) -> int:
    non_blank = [line for line in data_lines if line.strip()]
    if not non_blank:
        return 0
    has_blank_rows = any(not line.strip() for line in data_lines)
    if has_blank_rows:
        # Multiline: rows are blank-separated blocks (a block may be several
        # physical lines for a wrapped cell).
        return len([b for b in re.split(r"\n\s*\n", "\n".join(data_lines)) if b.strip()])
    # Simple: each non-blank line is a row.
    return len(non_blank)


def _simple_data_cells(data_lines: list[str], col_runs: int) -> list[list[str]]:
    has_blank_rows = any(not line.strip() for line in data_lines)
    if has_blank_rows:
        blocks = [b for b in re.split(r"\n\s*\n", "\n".join(data_lines)) if b.strip()]
        return [_split_simple_cells(" ".join(b.split()), col_runs) for b in blocks]
    return [_split_simple_cells(line, col_runs) for line in data_lines if line.strip()]


def _split_simple_cells(line: str, col_runs: int) -> list[str]:
    # Best-effort column split for simple/multiline tables: cells are
    # separated by 2+ spaces. Pad/truncate to the rule's column count.
    cells = [cell.strip() for cell in re.split(r" {2,}", line.strip()) if cell.strip()]
    if len(cells) < col_runs:
        cells.extend([""] * (col_runs - len(cells)))
    return cells[:col_runs]


def _is_simple_table_text(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return True  # blank lines separate multiline rows
    if stripped.startswith("|"):
        return False
    if _is_grid_separator_line(stripped):
        return False
    if _simple_rule_runs(line) is not None:
        return False
    return True


def _simple_rule_runs(line: str) -> int | None:
    """Dash-run count for a pandoc simple/multiline rule line, else ``None``.

    A rule line is runs of dashes (``>=3`` each) separated by spaces, with no
    ``|`` or ``+`` (those belong to pipe/grid tables). Returns the number of
    dash-runs (the column count when the rule is a column rule).
    """
    stripped = line.strip()
    if "|" in stripped or "+" in stripped:
        return None
    if not re.fullmatch(r"-{3,}(\s+-{3,})*", stripped):
        return None
    return len(re.findall(r"-{2,}", stripped))


def _parse_html_tables(content: str) -> list[tuple[tuple[int, int], list[list[str]]]]:
    """Parse ``<table>`` blocks, honoring ``colspan`` for column count.

    MinerU emits merged cells with ``colspan``/``rowspan``. Counting raw
    ``<td>`` tags undercounts columns (a ``colspan="2"`` header occupies two
    grid columns). ``colspan`` is horizontal and directly determines the
    column count, so we expand it (duplicating the cell text across the spanned
    columns) and size the table by the widest row. ``rowspan`` is vertical and
    does not affect column count; MinerU's ``rowspan`` values are noisy, so we
    deliberately ignore them rather than reconstruct a fragile carried grid.
    """
    tables: list[tuple[tuple[int, int], list[list[str]]]] = []
    for match in re.finditer(r"<table\b[^>]*>(.*?)</table>", content, re.DOTALL | re.IGNORECASE):
        grid: list[list[str]] = []
        for row_match in re.finditer(r"<tr\b[^>]*>(.*?)</tr>", match.group(1), re.DOTALL | re.IGNORECASE):
            row: list[str] = []
            for cell_match in re.finditer(
                r"<t[dh]\b([^>]*)>(.*?)</t[dh]>", row_match.group(1), re.DOTALL | re.IGNORECASE
            ):
                text = _strip_html(cell_match.group(2))
                colspan = _read_span(cell_match.group(1) or "", "colspan")
                row.extend([text] * colspan)
            if row:
                grid.append(row)
        if grid:
            cols = max(len(row) for row in grid)
            tables.append(((len(grid), cols), grid))
    return tables


def _read_span(attrs: str, name: str) -> int:
    match = re.search(rf"{name}\s*=\s*\"?(\d+)", attrs, re.IGNORECASE)
    return int(match.group(1)) if match else 1


def _finalize_table_block(block: list[str]) -> tuple[tuple[int, int], list[list[str]]]:
    # Drop pandoc grid-table border lines (``+---+`` / ``+===+``, including
    # partial borders); they carry no cell data and would otherwise corrupt
    # column counting.
    grid_borders = [line for line in block if _is_grid_separator_line(line)]
    block = [line for line in block if not _is_grid_separator_line(line)]
    rows = [_split_pipe_cells(line) for line in block]
    rows = [row for row in rows if row]
    if not rows:
        return ((0, 0), [])
    cols = len(rows[0])
    if grid_borders:
        # Pandoc grid tables expand merged cells across multiple physical
        # ``|`` lines, so counting ``|`` lines over-counts rows (~2x). Border
        # lines (full + partial) are the true logical-row boundaries: every
        # logical row emits exactly one border, so border count == logical
        # rows including the header.
        return ((len(grid_borders), cols), rows)
    data_rows = [row for index, row in enumerate(rows) if not _is_separator_row(row)]
    return ((len(data_rows), cols), data_rows)


def _split_pipe_cells(line: str) -> list[str]:
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|") and not stripped.endswith(r"\|"):
        stripped = stripped[:-1]
    return [cell.replace(r"\|", "|").strip() for cell in re.split(r"(?<!\\)\|", stripped)]


def _is_separator_row(row: list[str]) -> bool:
    return bool(row) and all(re.fullmatch(r":?-{1,}:?", cell) for cell in row if cell != "")


def _content_headings(text: str) -> list[tuple[int, str]]:
    headings: list[tuple[int, str]] = []
    lines = text.splitlines()
    for index, line in enumerate(lines):
        match = _HEADING_RE.match(line)
        if match:
            headings.append((len(match.group(1)), match.group(2).strip()))
            continue
        if index == 0:
            continue
        stripped = line.strip()
        previous = lines[index - 1].strip()
        if previous and re.fullmatch(r"=+", stripped):
            headings.append((1, previous))
        elif previous and re.fullmatch(r"-+", stripped):
            headings.append((2, previous))
    return headings


def _read_csv_rows(path: Path) -> list[list[str]]:
    with path.open(encoding="utf-8") as handle:
        return [[cell.strip() for cell in row] for row in csv.reader(handle)]


def _row_match_ratio(gold_rows: list[list[str]], parsed_rows: list[list[str]]) -> float:
    max_rows = max(len(gold_rows), len(parsed_rows))
    if max_rows == 0:
        return 1.0
    max_cols = max((len(row) for row in gold_rows + parsed_rows), default=0)
    matches = 0
    for row_index in range(max_rows):
        gold_row = gold_rows[row_index] if row_index < len(gold_rows) else []
        parsed_row = parsed_rows[row_index] if row_index < len(parsed_rows) else []
        for col_index in range(max_cols):
            gold_cell = gold_row[col_index] if col_index < len(gold_row) else ""
            parsed_cell = parsed_row[col_index] if col_index < len(parsed_row) else ""
            if _normalize(gold_cell) == _normalize(parsed_cell):
                matches += 1
    return matches / (max_rows * max_cols) if max_cols else 0.0


def _edit_distance(left: list[int], right: list[int]) -> int:
    previous = list(range(len(right) + 1))
    for i, left_value in enumerate(left, start=1):
        current = [i]
        for j, right_value in enumerate(right, start=1):
            cost = 0 if left_value == right_value else 1
            current.append(min(current[j - 1] + 1, previous[j] + 1, previous[j - 1] + cost))
        previous = current
    return previous[-1]


def _normalize(text: str) -> str:
    # Strip markdown inline marks (``**bold**``, ``*italic*``, ``_italic_``,
    # `` `code` ``) so a heading rendered as ``## **1.1 Education**`` matches
    # the GT heading ``1.1 Education`` for both text recall and parent-chain
    # equality. Backslash escapes are stripped too.
    text = _strip_markdown_inline_marks(text)
    collapsed = re.sub(r"\s+", " ", text).strip().lower()
    return collapsed


_SUPERSCRIPTS = {
    "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
    "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
}
_SUBSCRIPTS = {
    "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4",
    "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9",
}
_SUPERSCRIPT_TAG_RE = re.compile(r"<sup[^>]*>(.*?)</sup>", re.DOTALL | re.IGNORECASE)
_SUBSCRIPT_TAG_RE = re.compile(r"<sub[^>]*>(.*?)</sub>", re.DOTALL | re.IGNORECASE)
_CARET_SUPER_RE = re.compile(r"\^(\d)")


def _normalize_supsub(text: str) -> str:
    """Flatten super/subscripts to plain digits: ²/^2/<sup>2</sup> -> 2."""

    text = _SUPERSCRIPT_TAG_RE.sub(r"\1", text)
    text = _SUBSCRIPT_TAG_RE.sub(r"\1", text)
    text = _CARET_SUPER_RE.sub(r"\1", text)
    for char, digit in _SUPERSCRIPTS.items():
        text = text.replace(char, digit)
    for char, digit in _SUBSCRIPTS.items():
        text = text.replace(char, digit)
    return text


def _normalize_text(text: str) -> str:
    """Normalize for semantic text matching across frameworks.

    Decodes HTML entities (``&gt;`` -> ``>``), flattens super/subscripts
    (``²`` / ``^2`` / ``<sup>2</sup>`` -> ``2``), and strips remaining inline
    tags so the same key text matches regardless of how a parser rendered it,
    then collapses whitespace and lowercases.
    """

    import html

    text = _normalize_supsub(text)
    text = _strip_markdown_inline_marks(text)
    # Strip real HTML tags only (``<p>``/``</td>``/``<img ...>``): the ``<`` must
    # be followed by a letter, ``/`` or ``!``, and the tag may not span a
    # newline. A bare ``<`` glued to a digit (e.g. markitdown's literal
    # ``<3800`` in a table cell) is not a tag — the old ``<[^>]+>`` pattern
    # greedily ate everything from such a ``<`` up to the next ``>`` (e.g. the
    # ``>`` of ``>2100``), erasing the key text we meant to match.
    without_tags = re.sub(r"</?[a-zA-Z!][^>\n]*>", "", text)
    normalized = html.unescape(without_tags)
    normalized = re.sub(r"[‐‑‒–—―−]+", "-", normalized)
    normalized = re.sub(r"-{2,}", "-", normalized)
    # Unify the multiplication sign with the LaTeX command: GT uses ``×``
    # (U+00D7) while docling/mineru render ``\times`` (which escape-stripping
    # turns into ``times``). Map the bare ``×`` to ``times`` so both sides
    # agree regardless of rendering.
    normalized = normalized.replace("×", "times")
    # Collapse spaces around an underscore between word chars so docling's
    # ``group \_ loss`` (spaces around the escaped underscore) matches the
    # token ``group_loss``. Only bridges ``word _ word`` / ``word_ word`` /
    # ``word _word``; a bare ``_`` elsewhere is untouched.
    normalized = re.sub(r"(\w)\s*_\s*(\w)", r"\1_\2", normalized)
    return re.sub(r"\s+", " ", normalized).strip().lower()


def _strip_markdown_inline_marks(text: str) -> str:
    # Strip LaTeX font-shape wrappers (``\mathbf{X}``, ``\mathrm{X}``, ...)
    # before backslash-escape stripping, leaving the content ``X`` (and a spare
    # ``}``). Pandoc wraps every math token this way, so ``S_{\mathbf{down}}``
    # would otherwise normalize to ``{mathbf{down}}`` and miss the GT token
    # ``_{down}``. The leftover ``}`` does not affect substring matching.
    text = re.sub(r"\\(?:mathbf|mathrm|mathit|mathsf|mathtt|text|mathcal|mathbb|mathfrak)\s*\{", "", text)
    text = re.sub(r"(\*\*|__)(.*?)\1", r"\2", text, flags=re.DOTALL)
    # Italic/word-emphasis. ``*`` is lenient (``*bold*`` valid anywhere). ``_``
    # follows the CommonMark intraword rule: a ``_`` adjacent to a word char is
    # literal, so ``S_{down}`` / ``group_loss`` keep their underscore instead of
    # being paired with a distant ``_`` and stripped. A ``_`` immediately
    # followed by ``{`` is a LaTeX subscript (``_{down}``) and is also literal.
    # Backslash-preceded delimiters are rejected too so ``\_`` / ``\\_`` never
    # act as delimiters.
    text = re.sub(r"(?<!\\)(\*)([^*]+)(?<!\\)\1(?!\w)", r"\2", text, flags=re.DOTALL)
    text = re.sub(r"(?<![\w\\])(_)(?!{)([^_]+)(?<![\w\\])\1(?!{)(?!\w)", r"\2", text, flags=re.DOTALL)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    # Strip markdown backslash escapes. Parsers like markitdown over-escape
    # underscores as ``\\_`` (escaped backslash + escaped underscore); a single
    # pass leaves ``\_`` behind, so loop until stable. Also strip a backslash
    # before an ASCII letter so LaTeX commands like ``\max`` / ``\min`` /
    # ``\times`` collapse to plain ``max`` / ``min`` / ``times``. ``\frac``
    # becomes ``frac`` too; that is fine because the GT token ``\frac`` is
    # normalized the same way, so both sides still match.
    escape_re = re.compile(r"\\([\\`*_{}\[\]()#+\-.!<>|=~'\"]|[A-Za-z])")
    while True:
        new_text = escape_re.sub(r"\1", text)
        if new_text == text:
            break
        text = new_text
    return text


def _strip_html(fragment: str) -> str:
    without_tags = re.sub(r"<[^>]+>", "", fragment)
    import html

    return html.unescape(without_tags).strip()
