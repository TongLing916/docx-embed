from __future__ import annotations

from dataclasses import dataclass
from html import escape
import mimetypes
from pathlib import Path
import re
from typing import Any

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_EMBED_ATTR = f"{{{REL_NS}}}embed"
REL_ID_ATTR = f"{{{REL_NS}}}id"


@dataclass
class RagflowParseOutput:
    markdown: str
    artifacts: dict[str, Path]
    warnings: list[str]


def _parse_ragflow_body_order_docx(
    input_path: str | Path, artifact_dir: str | Path
) -> RagflowParseOutput:
    """Parse DOCX with a RAGFlow-style body-order section reader."""

    source = Path(input_path)
    artifacts = Path(artifact_dir)
    artifacts.mkdir(parents=True, exist_ok=True)
    image_dir = artifacts / "images"
    warnings: list[str] = []

    try:
        document = Document(source)
    except Exception as exc:
        return RagflowParseOutput(
            markdown="",
            artifacts={},
            warnings=[f"RAGFlow-style parser failed to open DOCX: {exc}"],
        )

    sections: list[dict[str, Any]] = []
    last_image: str | None = None
    image_count = 0
    heading_stack: list[tuple[int, str]] = []
    doc_name = re.sub(r"\.[a-zA-Z]+$", "", source.name) or "Untitled Document"

    def flush_last_image() -> None:
        nonlocal last_image
        if last_image is not None:
            sections.append({"text": "", "image": last_image, "table": None, "style": "Image"})
            last_image = None

    for block in document._element.body:
        if block.tag.endswith("p"):
            paragraph = Paragraph(block, document)
            text = _clean_text(paragraph.text)
            style_name = paragraph.style.name if paragraph.style else ""
            heading_level = _heading_level(style_name)
            if heading_level and text:
                heading_stack = _updated_heading_stack(heading_stack, heading_level, text)

            images = []
            for rel_id in _paragraph_image_rel_ids(paragraph):
                image_ref = _write_related_image(document, rel_id, image_dir, image_count + 1, warnings)
                if image_ref is None:
                    continue
                image_count += 1
                images.append(image_ref)

            if text:
                if style_name == "Caption":
                    former_image = None
                    if sections and sections[-1].get("image") and sections[-1].get("style") != "Caption":
                        former_image = str(sections[-1].get("image"))
                        sections.pop()
                    elif last_image is not None:
                        former_image = last_image
                        last_image = None
                    sections.append({"text": text, "image": former_image, "table": None, "style": "Caption"})
                else:
                    flush_last_image()
                    sections.append({"text": text, "image": None, "table": None, "style": style_name})
                    for image_ref in images:
                        sections.append({"text": "", "image": image_ref, "table": None, "style": "Image"})
            elif images:
                last_image = images[-1]
                for image_ref in images[:-1]:
                    sections.append({"text": "", "image": image_ref, "table": None, "style": "Image"})

        elif block.tag.endswith("tbl"):
            flush_last_image()
            table = Table(block, document)
            table_html, image_count = _render_table_html(
                table, document, doc_name, heading_stack, image_dir, image_count, warnings
            )
            sections.append(
                {
                    "text": "",
                    "image": None,
                    "table": table_html,
                    "style": "Table",
                }
            )

    flush_last_image()
    markdown = _render_markdown(sections)
    markdown_path = artifacts / f"{source.stem}.md"
    markdown_path.write_text(markdown, encoding="utf-8")
    return RagflowParseOutput(
        markdown=markdown,
        artifacts={"clean_markdown": markdown_path},
        warnings=warnings,
    )


def parse_ragflow_docx(input_path: str | Path, artifact_dir: str | Path) -> RagflowParseOutput:
    """Parse DOCX with RAGFlow's mammoth-to-markdown export style."""

    source = Path(input_path)
    artifacts = Path(artifact_dir)
    artifacts.mkdir(parents=True, exist_ok=True)
    image_dir = artifacts / "images"
    image_count = 0
    warnings: list[str] = []

    try:
        import mammoth
        from markdownify import markdownify
    except Exception as exc:
        return RagflowParseOutput(
            markdown="",
            artifacts={},
            warnings=[f"RAGFlow markdown export dependencies are unavailable: {exc}"],
        )

    def convert_image(image):
        nonlocal image_count
        image_count += 1
        content_type = getattr(image, "content_type", None)
        suffix = mimetypes.guess_extension(content_type or "") or ".bin"
        filename = f"image_{image_count:03d}{suffix}"
        destination = image_dir / filename
        try:
            image_dir.mkdir(parents=True, exist_ok=True)
            with image.open() as image_file:
                destination.write_bytes(image_file.read())
            return {
                "src": f"images/{filename}",
                "alt": destination.stem,
            }
        except Exception as exc:
            warnings.append(f"RAGFlow markdown export skipped image {filename}: {exc}")
            return {"src": "", "alt": destination.stem}

    try:
        with source.open("rb") as docx_file:
            html = mammoth.convert_to_html(
                docx_file,
                convert_image=mammoth.images.img_element(convert_image),
            ).value
    except Exception as exc:
        return RagflowParseOutput(
            markdown="",
            artifacts={},
            warnings=[f"RAGFlow markdown export failed to convert DOCX: {exc}"],
        )

    markdown = markdownify(html).rstrip() + "\n"
    markdown_path = artifacts / f"{source.stem}.md"
    markdown_path.write_text(markdown, encoding="utf-8")
    return RagflowParseOutput(
        markdown=markdown,
        artifacts={"clean_markdown": markdown_path},
        warnings=warnings,
    )


def _clean_text(text: str) -> str:
    return re.sub(r"\u3000", " ", text).strip()


def _heading_level(style_name: str) -> int | None:
    match = re.search(r"Heading\s*(\d+)", style_name, re.I)
    if not match:
        return None
    level = int(match.group(1))
    if level < 1 or level > 7:
        return None
    return level


def _updated_heading_stack(
    headings: list[tuple[int, str]], level: int, text: str
) -> list[tuple[int, str]]:
    return [heading for heading in headings if heading[0] < level] + [(level, text)]


def _paragraph_image_rel_ids(paragraph: Paragraph) -> list[str]:
    rel_ids: list[str] = []
    for element in paragraph._element.iter():
        if _local_name(element.tag) not in {"blip", "imagedata"}:
            continue
        rel_id = element.attrib.get(REL_EMBED_ATTR) or element.attrib.get(REL_ID_ATTR)
        if rel_id:
            rel_ids.append(rel_id)
    return rel_ids


def _write_related_image(
    document: Document,
    rel_id: str,
    image_dir: Path,
    image_number: int,
    warnings: list[str],
) -> str | None:
    try:
        related_part = document.part.related_parts[rel_id]
    except KeyError:
        warnings.append(f"RAGFlow-style parser skipped missing image relationship: {rel_id}")
        return None
    blob = getattr(related_part, "blob", None)
    if not blob:
        warnings.append(f"RAGFlow-style parser skipped empty image relationship: {rel_id}")
        return None

    content_type = getattr(related_part, "content_type", None)
    suffix = mimetypes.guess_extension(content_type or "") or ".bin"
    filename = f"image_{image_number:03d}{suffix}"
    image_dir.mkdir(parents=True, exist_ok=True)
    destination = image_dir / filename
    destination.write_bytes(blob)
    return f"images/{filename}"


def _render_table_html(
    table: Table,
    document: Document,
    doc_name: str,
    heading_stack: list[tuple[int, str]],
    image_dir: Path,
    image_count: int,
    warnings: list[str],
) -> tuple[str, int]:
    lines = ["<table>"]
    if heading_stack:
        hierarchy = " > ".join([doc_name, *(text for _, text in heading_stack)])
        lines.append(f"<caption>Table Location: {escape(hierarchy)}</caption>")
    for row in table.rows:
        cells = row.cells
        rendered_cells = []
        col_idx = 0
        while col_idx < len(cells):
            cell = cells[col_idx]
            span = 1
            while col_idx + span < len(cells) and cell.text == cells[col_idx + span].text:
                span += 1
            cell_text, image_count = _render_table_cell_html(
                cell, document, image_dir, image_count, warnings
            )
            if span == 1:
                rendered_cells.append(f"<td>{cell_text}</td>")
            else:
                rendered_cells.append(f"<td colspan='{span}'>{cell_text}</td>")
            col_idx += span
        lines.append("<tr>" + "".join(rendered_cells) + "</tr>")
    lines.append("</table>")
    return "\n".join(lines), image_count


def _render_table_cell_html(
    cell: Any,
    document: Document,
    image_dir: Path,
    image_count: int,
    warnings: list[str],
) -> tuple[str, int]:
    rendered_paragraphs: list[str] = []
    for paragraph in cell.paragraphs:
        text = escape(paragraph.text)
        image_refs: list[str] = []
        for rel_id in _paragraph_image_rel_ids(paragraph):
            image_ref = _write_related_image(document, rel_id, image_dir, image_count + 1, warnings)
            if image_ref is None:
                continue
            image_count += 1
            image_refs.append(_image_markdown_reference(image_ref))

        if image_refs:
            separator = "" if not text or text.endswith(" ") else " "
            rendered_paragraphs.append(f"{text}{separator}{' '.join(image_refs)}")
        elif text:
            rendered_paragraphs.append(text)

    if rendered_paragraphs:
        return "<br/>".join(rendered_paragraphs), image_count
    return escape(cell.text), image_count


def _render_markdown(sections: list[dict[str, Any]]) -> str:
    blocks: list[str] = []
    for section in sections:
        image = section.get("image")
        text = section.get("text")
        table = section.get("table")
        if image:
            blocks.append(_image_markdown_reference(str(image)))
        if text:
            blocks.append(str(text))
        if table:
            blocks.append(str(table))
    return "\n\n".join(blocks).rstrip() + "\n"


def _image_markdown_reference(image_ref: str) -> str:
    return f"![{Path(image_ref).stem}]({image_ref})"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
