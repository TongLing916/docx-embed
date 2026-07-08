from __future__ import annotations

from dataclasses import dataclass, field
import json
import os
from pathlib import Path
import shutil
import subprocess
from typing import Any
from urllib import error, request
import uuid
import zipfile
from xml.etree import ElementTree as ET

from edp.docx_notes import normalize_docx_notes_markdown
from edp.ragflow_parser import parse_ragflow_docx


DEFAULT_MINERU_BASE_URL = "http://127.0.0.1:8000"
MINERU_FILE_PARSE_PATH = "/file_parse"
MINERU_BACKEND_PARSERS = {
    "mineru": "pipeline",
    "mineru-pipeline": "pipeline",
    "mineru-vlm-engine": "vlm-engine",
    "mineru-hybrid-engine": "hybrid-engine",
}
DEFAULT_DOCLING_API_BASE = "http://127.0.0.1:5001"
DOCLING_CONVERT_FILE_PATH = "/v1/convert/file"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_ID_ATTR = f"{{{REL_NS}}}id"
REL_EMBED_ATTR = f"{{{REL_NS}}}embed"


@dataclass(frozen=True)
class MainParseResult:
    parser: str
    markdown: str
    artifacts: dict[str, Path] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


def parse_main_document(
    input_path: str | Path, parser_name: str, artifact_dir: str | Path
) -> MainParseResult:
    source = Path(input_path)
    artifacts = Path(artifact_dir)
    if parser_name == "markitdown":
        return _normalize_notes_result(source, _parse_with_markitdown(source, artifacts))
    mineru_backend = MINERU_BACKEND_PARSERS.get(parser_name)
    if mineru_backend is not None:
        return _normalize_notes_result(
            source,
            _parse_with_mineru(
                source,
                artifacts,
                backend=mineru_backend,
                parser_label=parser_name,
            ),
        )
    if parser_name == "docling":
        return _normalize_notes_result(source, _parse_with_docling(source, artifacts))
    if parser_name == "pandoc":
        return _parse_with_pandoc(source, artifacts)
    if parser_name == "ragflow":
        return _normalize_notes_result(source, _parse_with_ragflow(source, artifacts))
    return MainParseResult(
        parser=parser_name,
        markdown="",
        warnings=[f"Unsupported main parser: {parser_name}"],
    )


def _normalize_notes_result(input_path: Path, result: MainParseResult) -> MainParseResult:
    normalized = normalize_docx_notes_markdown(input_path, result.markdown)
    if normalized == result.markdown:
        return result

    clean_markdown = result.artifacts.get("clean_markdown")
    if clean_markdown is not None:
        clean_markdown.write_text(normalized, encoding="utf-8")
    return MainParseResult(
        parser=result.parser,
        markdown=normalized,
        artifacts=result.artifacts,
        warnings=result.warnings,
    )


def _parse_with_ragflow(input_path: Path, artifact_dir: Path) -> MainParseResult:
    result = parse_ragflow_docx(input_path, artifact_dir)
    return MainParseResult(
        parser="ragflow",
        markdown=result.markdown,
        artifacts=result.artifacts,
        warnings=result.warnings,
    )


def _parse_with_pandoc(input_path: Path, artifact_dir: Path) -> MainParseResult:
    """Shell out to ``pandoc input.docx -o main.md --wrap=preserve``.

    Pandoc converts DOCX straight to Markdown via its own reader (no layout
    model). Runs from ``artifact_dir`` with ``--extract-media=.`` so image
    links come out relative to the markdown file rather than absolute or
    CWD-relative. In pipeline mode the clean parent DOCX has its assets
    replaced by sentinels, so pandoc typically extracts no media here.
    """
    if not shutil.which("pandoc"):
        return MainParseResult(
            parser="pandoc",
            markdown="",
            warnings=["pandoc is not installed"],
        )
    artifact_dir.mkdir(parents=True, exist_ok=True)
    markdown_path = artifact_dir / f"{input_path.stem}.md"
    cmd = [
        "pandoc",
        str(input_path.resolve()),
        "-o",
        markdown_path.name,
        "--wrap=preserve",
        "--extract-media=.",
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, cwd=str(artifact_dir))
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()[:1000]
        return MainParseResult(
            parser="pandoc",
            markdown="",
            warnings=[f"pandoc failed (exit {proc.returncode}): {detail}"],
        )
    markdown = markdown_path.read_text(encoding="utf-8") if markdown_path.exists() else ""
    warnings: list[str] = []
    if proc.stderr.strip():
        warnings.append(f"pandoc stderr: {proc.stderr.strip()[:1000]}")
    return MainParseResult(
        parser="pandoc",
        markdown=markdown,
        artifacts={"clean_markdown": markdown_path} if markdown_path.exists() else {},
        warnings=warnings,
    )


def _parse_with_markitdown(input_path: Path, artifact_dir: Path) -> MainParseResult:
    artifact_dir.mkdir(parents=True, exist_ok=True)
    try:
        markdown = _convert_docx_with_markitdown_images(input_path, artifact_dir)
    except Exception as exc:
        return MainParseResult(
            parser="markitdown",
            markdown="",
            warnings=[f"MarkItDown failed to convert main document: {exc}"],
        )

    markdown_path = artifact_dir / f"{input_path.stem}.md"
    markdown_path.write_text(markdown, encoding="utf-8")
    return MainParseResult(
        parser="markitdown",
        markdown=markdown,
        artifacts={"clean_markdown": markdown_path},
    )


def _convert_docx_with_markitdown_images(input_path: Path, artifact_dir: Path) -> str:
    import mammoth
    from markitdown.converters._html_converter import HtmlConverter
    from markitdown.converter_utils.docx.pre_process import pre_process_docx

    image_dir = artifact_dir / "images"
    image_writer = _MarkItDownImageWriter(
        image_dir,
        image_rel_ids=_document_image_rel_ids(input_path),
        skipped_rel_ids=_ole_preview_image_rel_ids(input_path),
    )
    with input_path.open("rb") as stream:
        pre_processed = pre_process_docx(stream)
        html = mammoth.convert_to_html(
            pre_processed,
            convert_image=_markitdown_image_converter(mammoth, image_writer),
        ).value
    markdown = HtmlConverter().convert_string(html).markdown
    return markdown.rstrip() + "\n"


def _markitdown_image_converter(mammoth, image_writer):
    def convert_image(image):
        attributes = {}
        if image.alt_text:
            attributes["alt"] = image.alt_text
        image_attributes = image_writer(image)
        if image_attributes is None:
            return []
        attributes.update(image_attributes)
        return [mammoth.html.element("img", attributes)]

    return convert_image


class _MarkItDownImageWriter:
    def __init__(
        self,
        output_dir: Path,
        *,
        image_rel_ids: list[str],
        skipped_rel_ids: set[str],
    ) -> None:
        self._output_dir = output_dir
        self._image_rel_ids = image_rel_ids
        self._skipped_rel_ids = skipped_rel_ids
        self._source_image_index = 0
        self._image_number = 1

    def __call__(self, image) -> dict[str, str] | None:
        rel_id = self._next_rel_id()
        if rel_id in self._skipped_rel_ids:
            return None
        extension = image.content_type.partition("/")[2] or "bin"
        filename = f"doc_{self._image_number:03d}.{extension}"
        self._output_dir.mkdir(parents=True, exist_ok=True)
        destination = self._output_dir / filename
        with image.open() as source, destination.open("wb") as target:
            shutil.copyfileobj(source, target)
        self._image_number += 1
        return {"src": f"images/{filename}"}

    def _next_rel_id(self) -> str | None:
        if self._source_image_index >= len(self._image_rel_ids):
            return None
        rel_id = self._image_rel_ids[self._source_image_index]
        self._source_image_index += 1
        return rel_id


def _document_image_rel_ids(input_path: Path) -> list[str]:
    root = _document_xml_root(input_path)
    if root is None:
        return []
    rel_ids: list[str] = []
    for element in root.iter():
        if _local_name(element.tag) not in {"blip", "imagedata"}:
            continue
        rel_id = element.attrib.get(REL_EMBED_ATTR) or element.attrib.get(REL_ID_ATTR)
        if rel_id:
            rel_ids.append(rel_id)
    return rel_ids


def _ole_preview_image_rel_ids(input_path: Path) -> set[str]:
    root = _document_xml_root(input_path)
    if root is None:
        return set()
    preview_rel_ids: set[str] = set()
    for run in root.iter(f"{{{WORD_NS}}}r"):
        if not any(_local_name(element.tag) == "OLEObject" for element in run.iter()):
            continue
        for element in run.iter():
            if _local_name(element.tag) not in {"blip", "imagedata"}:
                continue
            rel_id = element.attrib.get(REL_EMBED_ATTR) or element.attrib.get(REL_ID_ATTR)
            if rel_id:
                preview_rel_ids.add(rel_id)
    return preview_rel_ids


def _document_xml_root(input_path: Path) -> ET.Element | None:
    try:
        with zipfile.ZipFile(input_path) as docx:
            return ET.fromstring(docx.read("word/document.xml"))
    except (KeyError, ET.ParseError, zipfile.BadZipFile):
        return None


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _parse_with_mineru(
    input_path: Path,
    artifact_dir: Path,
    *,
    backend: str = "pipeline",
    parser_label: str = "mineru",
) -> MainParseResult:
    artifact_dir.mkdir(parents=True, exist_ok=True)
    authorization = _mineru_authorization()
    if not authorization:
        return MainParseResult(
            parser=parser_label,
            markdown="",
            warnings=["MinerU Authorization key is not configured"],
        )

    endpoint = _mineru_endpoint()
    timeout = float(os.environ.get("MINERU_REQUEST_TIMEOUT") or 600)
    body, content_type = _multipart_body(input_path, _mineru_fields(backend))
    req = request.Request(
        endpoint,
        data=body,
        headers={
            "Authorization": authorization,
            "Accept": "application/zip, application/json",
            "Content-Type": content_type,
        },
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=timeout) as response:
            raw = response.read()
            response_content_type = response.headers.get("Content-Type")
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:1000]
        return MainParseResult(
            parser=parser_label,
            markdown="",
            warnings=[f"MinerU service returned HTTP {exc.code}: {details}"],
        )
    except (error.URLError, TimeoutError, OSError) as exc:
        return MainParseResult(
            parser=parser_label,
            markdown="",
            warnings=[f"MinerU service request failed: {exc}"],
        )

    if _looks_like_zip(raw, response_content_type):
        return _mineru_zip_result(input_path, artifact_dir, raw, parser_label)
    return _mineru_json_result(input_path, artifact_dir, raw, parser_label)


def _parse_with_docling(input_path: Path, artifact_dir: Path) -> MainParseResult:
    artifact_dir.mkdir(parents=True, exist_ok=True)
    authorization = os.environ.get("DOCLING_AUTHORIZATION")
    if not authorization:
        return MainParseResult(
            parser="docling",
            markdown="",
            warnings=["Docling Authorization key is not configured"],
        )

    timeout = float(os.environ.get("DOCLING_REQUEST_TIMEOUT") or 240)
    body, content_type = _multipart_body(input_path, _docling_fields())
    req = request.Request(
        _docling_endpoint(),
        data=body,
        headers={
            "Authorization": authorization,
            "Accept": "application/json, application/zip",
            "Content-Type": content_type,
        },
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=timeout) as response:
            raw = response.read()
            response_content_type = response.headers.get("Content-Type")
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:1000]
        return MainParseResult(
            parser="docling",
            markdown="",
            warnings=[f"Docling service returned HTTP {exc.code}: {details}"],
        )
    except (error.URLError, TimeoutError, OSError) as exc:
        return MainParseResult(
            parser="docling",
            markdown="",
            warnings=[f"Docling service request failed: {exc}"],
        )

    if _looks_like_zip(raw, response_content_type):
        return _docling_zip_result(input_path, artifact_dir, raw)
    return _docling_json_result(input_path, artifact_dir, raw)


def _mineru_authorization() -> str | None:
    return (
        os.environ.get("MINERU_API_KEY")
        or os.environ.get("MINERU_AUTHORIZATION")
        or os.environ.get("DOCUMENT_CONVERTER_AUTHORIZATION")
    )


def _mineru_endpoint() -> str:
    endpoint = os.environ.get("MINERU_API_ENDPOINT")
    if endpoint:
        return endpoint
    base_url = (os.environ.get("MINERU_BASE_URL") or DEFAULT_MINERU_BASE_URL).rstrip("/")
    return f"{base_url}{MINERU_FILE_PARSE_PATH}"


def _mineru_zip_result(
    input_path: Path, artifact_dir: Path, raw: bytes, parser_label: str = "mineru"
) -> MainParseResult:
    zip_path = artifact_dir / f"{input_path.stem}.response.zip"
    zip_dir = artifact_dir / "zip"
    zip_path.write_bytes(raw)
    if zip_dir.exists():
        shutil.rmtree(zip_dir)
    zip_dir.mkdir(parents=True)
    try:
        with zipfile.ZipFile(zip_path) as archive:
            _extract_zip_safely(archive, zip_dir)
    except (ValueError, zipfile.BadZipFile, OSError) as exc:
        return MainParseResult(
            parser=parser_label,
            markdown="",
            artifacts={"parser_zip": zip_path},
            warnings=[f"MinerU zip response could not be unpacked: {exc}"],
        )

    markdown_paths = sorted(zip_dir.rglob("*.md"))
    if not markdown_paths:
        return MainParseResult(
            parser=parser_label,
            markdown="",
            artifacts={"parser_zip": zip_path, "parser_zip_dir": zip_dir},
            warnings=["MinerU zip response did not include Markdown content"],
        )
    markdown_path = markdown_paths[0]
    markdown = markdown_path.read_text(encoding="utf-8").rstrip() + "\n"
    return MainParseResult(
        parser=parser_label,
        markdown=markdown,
        artifacts={
            "parser_zip": zip_path,
            "parser_zip_dir": zip_dir,
            "clean_markdown": markdown_path,
        },
    )


def _docling_endpoint() -> str:
    endpoint = os.environ.get("DOCLING_API_ENDPOINT")
    if endpoint:
        return endpoint
    base_url = (os.environ.get("DOCLING_API_BASE") or DEFAULT_DOCLING_API_BASE).rstrip("/")
    if base_url.endswith(DOCLING_CONVERT_FILE_PATH):
        return base_url
    return f"{base_url}{DOCLING_CONVERT_FILE_PATH}"


def _docling_zip_result(input_path: Path, artifact_dir: Path, raw: bytes) -> MainParseResult:
    zip_path = artifact_dir / f"{input_path.stem}.response.zip"
    zip_dir = artifact_dir / "zip"
    zip_path.write_bytes(raw)
    if zip_dir.exists():
        shutil.rmtree(zip_dir)
    zip_dir.mkdir(parents=True)
    try:
        with zipfile.ZipFile(zip_path) as archive:
            for member in archive.infolist():
                if member.is_dir():
                    continue
                target = _safe_zip_member_path(zip_dir, member.filename)
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source, target.open("wb") as destination:
                    shutil.copyfileobj(source, destination)
    except (ValueError, zipfile.BadZipFile, OSError) as exc:
        return MainParseResult(
            parser="docling",
            markdown="",
            artifacts={"parser_zip": zip_path},
            warnings=[f"Docling zip response could not be unpacked: {exc}"],
        )

    markdown_paths = sorted(zip_dir.rglob("*.md"))
    if not markdown_paths:
        return MainParseResult(
            parser="docling",
            markdown="",
            artifacts={"parser_zip": zip_path, "parser_zip_dir": zip_dir},
            warnings=["Docling zip response did not include Markdown content"],
        )
    markdown_path = _best_markdown_path(markdown_paths, input_path.stem)
    markdown = markdown_path.read_text(encoding="utf-8").rstrip() + "\n"
    return MainParseResult(
        parser="docling",
        markdown=markdown,
        artifacts={
            "parser_zip": zip_path,
            "parser_zip_dir": zip_dir,
            "clean_markdown": markdown_path,
        },
    )


def _docling_json_result(input_path: Path, artifact_dir: Path, raw: bytes) -> MainParseResult:
    response_path = artifact_dir / f"{input_path.stem}.response.json"
    response_path.write_bytes(raw)
    try:
        data = json.loads(raw.decode("utf-8"))
    except json.JSONDecodeError as exc:
        return MainParseResult(
            parser="docling",
            markdown="",
            artifacts={"parser_response": response_path},
            warnings=[f"Docling service returned non-JSON response: {exc}"],
        )

    artifacts: dict[str, Path] = {"parser_response": response_path}
    status = data.get("status")
    if status != "success":
        details = json.dumps(data.get("errors", []), ensure_ascii=False)
        return MainParseResult(
            parser="docling",
            markdown="",
            artifacts=artifacts,
            warnings=[f"Docling service returned {status or 'unknown status'}: {details}"],
        )

    document_data = data.get("document")
    if not isinstance(document_data, dict):
        return MainParseResult(
            parser="docling",
            markdown="",
            artifacts=artifacts,
            warnings=["Docling response did not include a document object"],
        )

    markdown = document_data.get("md_content")
    if not isinstance(markdown, str) or not markdown.strip():
        return MainParseResult(
            parser="docling",
            markdown="",
            artifacts=artifacts,
            warnings=["Docling response did not include Markdown content"],
        )

    markdown = markdown.rstrip() + "\n"
    markdown_path = artifact_dir / f"{input_path.stem}.md"
    markdown_path.write_text(markdown, encoding="utf-8")
    artifacts["clean_markdown"] = markdown_path

    native_json = document_data.get("json_content")
    if native_json not in (None, "", [], {}):
        native_path = artifact_dir / f"{input_path.stem}.docling.json"
        native_path.write_text(json.dumps(native_json, ensure_ascii=False, indent=2), encoding="utf-8")
        artifacts["parser_native_json"] = native_path

    return MainParseResult(parser="docling", markdown=markdown, artifacts=artifacts)


def _mineru_json_result(
    input_path: Path, artifact_dir: Path, raw: bytes, parser_label: str = "mineru"
) -> MainParseResult:
    response_path = artifact_dir / f"{input_path.stem}.response.json"
    response_path.write_bytes(raw)
    try:
        data = json.loads(raw.decode("utf-8"))
    except json.JSONDecodeError as exc:
        return MainParseResult(
            parser=parser_label,
            markdown="",
            artifacts={"parser_response": response_path},
            warnings=[f"MinerU service returned non-JSON response: {exc}"],
        )

    result_data = _extract_mineru_result(data, input_path.stem)
    if data.get("status") not in {None, "completed"}:
        return MainParseResult(
            parser=parser_label,
            markdown="",
            artifacts={"parser_response": response_path},
            warnings=[f"MinerU service returned {data.get('status') or 'unknown status'}"],
        )
    markdown = result_data.get("md_content")
    if not isinstance(markdown, str) or not markdown.strip():
        return MainParseResult(
            parser=parser_label,
            markdown="",
            artifacts={"parser_response": response_path},
            warnings=["MinerU response did not include Markdown content"],
        )
    markdown = markdown.rstrip() + "\n"
    markdown_path = artifact_dir / f"{input_path.stem}.md"
    markdown_path.write_text(markdown, encoding="utf-8")
    return MainParseResult(
        parser=parser_label,
        markdown=markdown,
        artifacts={"parser_response": response_path, "clean_markdown": markdown_path},
    )


def _mineru_fields(backend: str = "pipeline") -> list[tuple[str, str]]:
    fields = [
        ("lang_list", "ch"),
        ("backend", os.environ.get("MINERU_BACKEND") or backend),
        ("parse_method", "auto"),
        ("formula_enable", "true"),
        ("table_enable", "true"),
        ("return_md", "true"),
        ("return_middle_json", "false"),
        ("return_content_list", "false"),
        ("return_images", "true"),
        ("response_format_zip", "true"),
        ("return_original_file", "false"),
    ]
    model_version = os.environ.get("MINERU_MODEL_VERSION")
    if model_version:
        fields.append(("model_version", model_version))
    return fields


def _docling_fields() -> list[tuple[str, str]]:
    return [
        ("to_formats", "md"),
        ("to_formats", "json"),
        ("do_ocr", "true"),
        ("include_images", "true"),
        ("include_page_images", "true"),
        ("table_mode", "accurate"),
        ("image_export_mode", "referenced"),
        ("target_type", "zip"),
    ]


def _multipart_body(input_path: Path, fields: list[tuple[str, str]]) -> tuple[bytes, str]:
    boundary = f"----edp-{uuid.uuid4().hex}"
    chunks: list[bytes] = []
    for name, value in fields:
        chunks.extend(
            [
                f"--{boundary}\r\n".encode(),
                f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode(),
                value.encode(),
                b"\r\n",
            ]
        )
    chunks.extend(
        [
            f"--{boundary}\r\n".encode(),
            (
                f'Content-Disposition: form-data; name="files"; filename="{input_path.name}"\r\n'
            ).encode(),
            b"Content-Type: application/octet-stream\r\n\r\n",
            input_path.read_bytes(),
            b"\r\n",
            f"--{boundary}--\r\n".encode(),
        ]
    )
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def _looks_like_zip(raw: bytes, content_type: str | None) -> bool:
    return raw.startswith(b"PK\x03\x04") or "zip" in (content_type or "").lower()


def _best_markdown_path(markdown_paths: list[Path], stem: str) -> Path:
    return sorted(
        markdown_paths,
        key=lambda path: (
            path.stem != stem,
            len(path.parts),
            str(path),
        ),
    )[0]


def _safe_zip_member_path(root: Path, member_name: str) -> Path:
    target = (root / member_name).resolve()
    root_resolved = root.resolve()
    if target != root_resolved and root_resolved not in target.parents:
        raise ValueError(f"unsafe zip member path: {member_name}")
    return target


def _extract_zip_safely(archive: zipfile.ZipFile, root: Path) -> None:
    for member in archive.infolist():
        if member.is_dir():
            continue
        target = _safe_zip_member_path(root, member.filename)
        target.parent.mkdir(parents=True, exist_ok=True)
        with archive.open(member) as source, target.open("wb") as destination:
            shutil.copyfileobj(source, destination)


def _extract_mineru_result(data: dict[str, Any], stem: str) -> dict[str, Any]:
    results = data.get("results")
    if not isinstance(results, dict) or not results:
        return {}
    if isinstance(results.get(stem), dict):
        return results[stem]
    first = next(iter(results.values()))
    return first if isinstance(first, dict) else {}
