from __future__ import annotations

from dataclasses import dataclass
import re
from pathlib import Path
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree as ET


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WORD_ID_ATTR = f"{{{WORD_NS}}}id"
WORD_TYPE_ATTR = f"{{{WORD_NS}}}type"


@dataclass(frozen=True)
class DocxNote:
    sequence: int
    kind: str
    source_id: str
    text: str
    before_context: str
    after_context: str


def normalize_docx_notes_markdown(docx_path: str | Path, markdown: str) -> str:
    notes = extract_docx_notes(docx_path)
    if not notes or not markdown:
        return markdown
    if _has_complete_markdown_notes(markdown, notes):
        return markdown

    normalized = _replace_broken_note_anchors(markdown, notes)
    normalized = _drop_broken_note_list(normalized)
    normalized = _insert_missing_note_markers(normalized, notes)
    normalized = _append_note_definitions(normalized, notes)
    return normalized


def extract_docx_notes(docx_path: str | Path) -> list[DocxNote]:
    source = Path(docx_path)
    if source.suffix.lower() != ".docx":
        return []

    try:
        with ZipFile(source) as archive:
            document_xml = archive.read("word/document.xml")
            footnotes = _read_note_part(archive, "word/footnotes.xml", "footnote")
            endnotes = _read_note_part(archive, "word/endnotes.xml", "endnote")
    except (BadZipFile, KeyError, OSError, ET.ParseError):
        return []

    try:
        document_root = ET.fromstring(document_xml)
    except ET.ParseError:
        return []

    refs = _document_note_refs(document_root)
    notes: list[DocxNote] = []
    for sequence, ref in enumerate(refs, start=1):
        note_text = (footnotes if ref["kind"] == "footnote" else endnotes).get(ref["source_id"], "")
        if not note_text:
            continue
        notes.append(
            DocxNote(
                sequence=sequence,
                kind=ref["kind"],
                source_id=ref["source_id"],
                text=note_text,
                before_context=ref["before_context"],
                after_context=ref["after_context"],
            )
        )
    return notes


def _read_note_part(archive: ZipFile, part_name: str, local_name: str) -> dict[str, str]:
    try:
        root = ET.fromstring(archive.read(part_name))
    except KeyError:
        return {}

    notes: dict[str, str] = {}
    for note in root:
        if _local_name(note.tag) != local_name:
            continue
        note_id = note.attrib.get(WORD_ID_ATTR)
        note_type = note.attrib.get(WORD_TYPE_ATTR)
        if note_id is None or note_type in {"separator", "continuationSeparator"}:
            continue
        text = _element_text(note)
        if text:
            notes[note_id] = text
    return notes


def _document_note_refs(root: ET.Element) -> list[dict[str, str]]:
    tokens: list[tuple[str, str, str]] = []
    _collect_document_tokens(root, tokens)

    refs: list[dict[str, str]] = []
    running_text = ""
    for index, token in enumerate(tokens):
        token_type = token[0]
        if token_type == "text":
            running_text += token[1]
            continue

        following_text = "".join(value for kind, value, _source_id in tokens[index + 1 :] if kind == "text")
        refs.append(
            {
                "kind": token_type,
                "source_id": token[1],
                "before_context": _context_suffix(running_text),
                "after_context": _context_prefix(following_text),
            }
        )
    return refs


def _collect_document_tokens(element: ET.Element, tokens: list[tuple[str, str, str]]) -> None:
    local_name = _local_name(element.tag)
    if local_name == "t" and element.text:
        tokens.append(("text", element.text, ""))
        return
    if local_name == "tab":
        tokens.append(("text", "\t", ""))
        return
    if local_name in {"br", "cr"}:
        tokens.append(("text", "\n", ""))
        return
    if local_name == "footnoteReference":
        note_id = element.attrib.get(WORD_ID_ATTR)
        if note_id:
            tokens.append(("footnote", note_id, ""))
        return
    if local_name == "endnoteReference":
        note_id = element.attrib.get(WORD_ID_ATTR)
        if note_id:
            tokens.append(("endnote", note_id, ""))
        return

    for child in element:
        _collect_document_tokens(child, tokens)
    if local_name == "p":
        tokens.append(("text", "\n", ""))


def _element_text(element: ET.Element) -> str:
    chunks: list[str] = []
    for item in element.iter():
        local_name = _local_name(item.tag)
        if local_name == "t" and item.text:
            chunks.append(item.text)
        elif local_name == "tab":
            chunks.append("\t")
        elif local_name in {"br", "cr", "p"} and chunks:
            chunks.append(" ")
    return _normalize_spaces("".join(chunks))


def _replace_broken_note_anchors(markdown: str, notes: list[DocxNote]) -> str:
    by_kind_ordinal: dict[tuple[str, int], int] = {}
    by_kind_source_id: dict[tuple[str, str], int] = {}
    kind_counts: dict[str, int] = {"footnote": 0, "endnote": 0}
    for note in notes:
        kind_counts[note.kind] += 1
        by_kind_ordinal[(note.kind, kind_counts[note.kind])] = note.sequence
        by_kind_source_id[(note.kind, note.source_id)] = note.sequence

    def markdown_anchor(match: re.Match[str]) -> str:
        kind = match.group("kind")
        index = int(match.group("index"))
        sequence = by_kind_ordinal.get((kind, index)) or by_kind_source_id.get((kind, str(index)))
        if sequence is None:
            return match.group(0)
        return f"[^{sequence}]"

    updated = re.sub(
        r"\[\[\d+\]\]\(#(?P<kind>footnote|endnote)-(?P<index>\d+)\)",
        markdown_anchor,
        markdown,
    )

    def html_anchor(match: re.Match[str]) -> str:
        kind = match.group("kind")
        index = int(match.group("index"))
        sequence = by_kind_ordinal.get((kind, index)) or by_kind_source_id.get((kind, str(index)))
        if sequence is None:
            return match.group(0)
        return f"[^{sequence}]"

    return re.sub(
        r"(?:<sup>)?\s*<a\s+href=[\"']#(?P<kind>footnote|endnote)-(?P<index>\d+)[\"']>\[\d+\]</a>\s*(?:</sup>)?",
        html_anchor,
        updated,
        flags=re.IGNORECASE,
    )


def _drop_broken_note_list(markdown: str) -> str:
    lines = markdown.splitlines()
    kept = [
        line
        for line in lines
        if not re.match(r"^\s*\d+\.\s+.*\s+\[↑\]\(#(?:footnote|endnote)-ref-\d+\)\s*$", line)
    ]
    return "\n".join(kept).rstrip() + ("\n" if markdown.endswith("\n") else "")


def _insert_missing_note_markers(markdown: str, notes: list[DocxNote]) -> str:
    updated = markdown
    search_start = 0
    for note in notes:
        marker = f"[^{note.sequence}]"
        if marker in updated:
            continue
        inserted = _insert_marker_after_context(updated, note.before_context, marker, search_start)
        if inserted == updated:
            updated = marker + updated
            search_start = len(marker)
        else:
            search_start = inserted.find(marker, search_start) + len(marker)
            updated = inserted
    return updated


def _insert_marker_after_context(markdown: str, context: str, marker: str, search_start: int) -> str:
    normalized_markdown, index_map = _normalized_index_map(markdown)
    normalized_context = _normalize_spaces(context)
    if not normalized_context:
        return markdown
    normalized_search_start = _normalized_offset_for_raw_index(index_map, search_start)
    match_index = normalized_markdown.find(normalized_context, normalized_search_start)
    if match_index < 0:
        return markdown
    raw_insert_at = index_map[match_index + len(normalized_context) - 1] + 1
    if markdown[raw_insert_at : raw_insert_at + len(marker)] == marker:
        return markdown
    return markdown[:raw_insert_at] + f" {marker}" + markdown[raw_insert_at:]


def _append_note_definitions(markdown: str, notes: list[DocxNote]) -> str:
    body = re.sub(r"\n{3,}", "\n\n", markdown).rstrip()
    definitions = [
        f"[^{note.sequence}]: {note.text}"
        for note in notes
        if not re.search(rf"(?m)^\[\^{note.sequence}\]:", body)
    ]
    if not definitions:
        return body + ("\n" if markdown.endswith("\n") else "")
    prefix = f"{body}\n\n" if body else ""
    return prefix + "\n\n".join(definitions) + "\n"


def _has_complete_markdown_notes(markdown: str, notes: list[DocxNote]) -> bool:
    return all(
        re.search(rf"\[\^{note.sequence}\]", markdown)
        and re.search(rf"(?m)^\[\^{note.sequence}\]:", markdown)
        for note in notes
    )


def _normalized_index_map(text: str) -> tuple[str, list[int]]:
    chars: list[str] = []
    index_map: list[int] = []
    previous_space = True
    for index, char in enumerate(text):
        if char.isspace():
            if previous_space:
                continue
            chars.append(" ")
            index_map.append(index)
            previous_space = True
        else:
            chars.append(char)
            index_map.append(index)
            previous_space = False
    if chars and chars[-1] == " ":
        chars.pop()
        index_map.pop()
    return "".join(chars), index_map


def _normalized_offset_for_raw_index(index_map: list[int], raw_index: int) -> int:
    for normalized_index, mapped_raw_index in enumerate(index_map):
        if mapped_raw_index >= raw_index:
            return normalized_index
    return len(index_map)


def _context_suffix(text: str, limit: int = 80) -> str:
    paragraph_text = re.split(r"[\r\n]+", text)[-1]
    normalized = _normalize_spaces(paragraph_text)
    if len(normalized) <= limit:
        return normalized
    suffix = normalized[-limit:]
    first_space = suffix.find(" ")
    return suffix[first_space + 1 :] if first_space >= 0 else suffix


def _context_prefix(text: str, limit: int = 80) -> str:
    normalized = _normalize_spaces(text)
    if len(normalized) <= limit:
        return normalized
    prefix = normalized[:limit]
    last_space = prefix.rfind(" ")
    return prefix[:last_space] if last_space >= 0 else prefix


def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
