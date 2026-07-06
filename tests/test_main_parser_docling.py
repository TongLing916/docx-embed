from __future__ import annotations

from io import BytesIO
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from edp.main_parser import parse_main_document


def test_docling_parser_reports_missing_authorization(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    monkeypatch.delenv("DOCLING_AUTHORIZATION", raising=False)

    result = parse_main_document(source, "docling", tmp_path / "artifacts")

    assert result.parser == "docling"
    assert result.markdown == ""
    assert result.warnings == ["Docling Authorization key is not configured"]


def test_docling_parser_accepts_json_response_and_writes_artifacts(
    tmp_path: Path, monkeypatch
) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    payload = json.dumps(
        {
            "status": "success",
            "document": {
                "md_content": "# Docling Result\n\nBody\n",
                "json_content": {"schema_name": "docling"},
            },
        }
    ).encode()

    class FakeResponse:
        headers = {"Content-Type": "application/json"}

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def read(self) -> bytes:
            return payload

    captured = {}

    def fake_urlopen(request, timeout):
        captured["timeout"] = timeout
        captured["url"] = request.full_url
        captured["authorization"] = request.headers["Authorization"]
        captured["accept"] = request.headers["Accept"]
        captured["body"] = request.data
        return FakeResponse()

    monkeypatch.setenv("DOCLING_AUTHORIZATION", "docling-token")
    monkeypatch.setenv("DOCLING_API_BASE", "https://docling.example.test/open-api/docling")
    monkeypatch.delenv("DOCLING_API_ENDPOINT", raising=False)
    monkeypatch.setattr("edp.main_parser.request.urlopen", fake_urlopen)

    result = parse_main_document(source, "docling", tmp_path / "artifacts")

    assert result.parser == "docling"
    assert result.markdown == "# Docling Result\n\nBody\n"
    assert result.warnings == []
    assert captured["url"] == "https://docling.example.test/open-api/docling/v1/convert/file"
    assert captured["authorization"] == "docling-token"
    assert captured["accept"] == "application/json, application/zip"
    assert captured["timeout"] == 240
    assert b'name="files"; filename="source.docx"' in captured["body"]
    assert b'name="to_formats"\r\n\r\nmd' in captured["body"]
    assert b'name="to_formats"\r\n\r\njson' in captured["body"]
    assert b'name="target_type"\r\n\r\nzip' in captured["body"]
    assert (tmp_path / "artifacts" / "source.response.json").exists()
    assert (tmp_path / "artifacts" / "source.md").read_text(encoding="utf-8") == result.markdown
    native = tmp_path / "artifacts" / "source.docling.json"
    assert json.loads(native.read_text(encoding="utf-8")) == {"schema_name": "docling"}
    assert result.artifacts["parser_response"] == tmp_path / "artifacts" / "source.response.json"
    assert result.artifacts["clean_markdown"] == tmp_path / "artifacts" / "source.md"
    assert result.artifacts["parser_native_json"] == native


def test_docling_parser_accepts_zip_response_and_writes_artifacts(
    tmp_path: Path, monkeypatch
) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    zip_stream = BytesIO()
    with ZipFile(zip_stream, "w", ZIP_DEFLATED) as archive:
        archive.writestr("document/source.md", "# Docling Zip\n\nBody\n")
        archive.writestr("document/images/page_1.png", b"png")
    payload = zip_stream.getvalue()

    class FakeResponse:
        headers = {"Content-Type": "application/zip"}

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def read(self) -> bytes:
            return payload

    def fake_urlopen(request, timeout):
        return FakeResponse()

    monkeypatch.setenv("DOCLING_AUTHORIZATION", "docling-token")
    monkeypatch.setattr("edp.main_parser.request.urlopen", fake_urlopen)

    result = parse_main_document(source, "docling", tmp_path / "artifacts")

    assert result.parser == "docling"
    assert result.markdown == "# Docling Zip\n\nBody\n"
    assert result.warnings == []
    assert (tmp_path / "artifacts" / "source.response.zip").exists()
    assert (
        tmp_path / "artifacts" / "zip" / "document" / "source.md"
    ).read_text(encoding="utf-8") == result.markdown
    assert (tmp_path / "artifacts" / "zip" / "document" / "images" / "page_1.png").read_bytes() == b"png"
    assert result.artifacts["parser_zip"] == tmp_path / "artifacts" / "source.response.zip"
    assert result.artifacts["parser_zip_dir"] == tmp_path / "artifacts" / "zip"
    assert result.artifacts["clean_markdown"] == tmp_path / "artifacts" / "zip" / "document" / "source.md"


def test_docling_parser_rejects_unsafe_zip_member_paths(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    zip_stream = BytesIO()
    with ZipFile(zip_stream, "w", ZIP_DEFLATED) as archive:
        archive.writestr("../escape.md", "# Escape\n")
    payload = zip_stream.getvalue()

    class FakeResponse:
        headers = {"Content-Type": "application/zip"}

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def read(self) -> bytes:
            return payload

    monkeypatch.setenv("DOCLING_AUTHORIZATION", "docling-token")
    monkeypatch.setattr("edp.main_parser.request.urlopen", lambda request, timeout: FakeResponse())

    result = parse_main_document(source, "docling", tmp_path / "artifacts")

    assert result.parser == "docling"
    assert result.markdown == ""
    assert result.warnings == ["Docling zip response could not be unpacked: unsafe zip member path: ../escape.md"]
    assert not (tmp_path / "escape.md").exists()


def test_docling_parser_reports_service_failure_status(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    payload = json.dumps({"status": "failed", "errors": [{"message": "bad input"}]}).encode()

    class FakeResponse:
        headers = {"Content-Type": "application/json"}

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def read(self) -> bytes:
            return payload

    monkeypatch.setenv("DOCLING_AUTHORIZATION", "docling-token")
    monkeypatch.setattr("edp.main_parser.request.urlopen", lambda request, timeout: FakeResponse())

    result = parse_main_document(source, "docling", tmp_path / "artifacts")

    assert result.parser == "docling"
    assert result.markdown == ""
    assert result.warnings == ['Docling service returned failed: [{"message": "bad input"}]']
    assert result.artifacts["parser_response"] == tmp_path / "artifacts" / "source.response.json"


def test_docling_endpoint_prefers_full_endpoint_override(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    payload = json.dumps(
        {"status": "success", "document": {"md_content": "# Endpoint\n"}}
    ).encode()

    class FakeResponse:
        headers = {"Content-Type": "application/json"}

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def read(self) -> bytes:
            return payload

    captured = {}

    def fake_urlopen(request, timeout):
        captured["url"] = request.full_url
        return FakeResponse()

    monkeypatch.setenv("DOCLING_AUTHORIZATION", "docling-token")
    monkeypatch.setenv("DOCLING_API_BASE", "https://ignored.example.test")
    monkeypatch.setenv("DOCLING_API_ENDPOINT", "https://docling.example.test/custom/convert")
    monkeypatch.setattr("edp.main_parser.request.urlopen", fake_urlopen)

    result = parse_main_document(source, "docling", tmp_path / "artifacts")

    assert result.warnings == []
    assert captured["url"] == "https://docling.example.test/custom/convert"
