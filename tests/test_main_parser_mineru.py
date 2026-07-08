from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from edp.main_parser import parse_main_document, _mineru_endpoint


def test_mineru_parser_accepts_zip_response_and_writes_artifacts(
    tmp_path: Path, monkeypatch
) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")

    zip_stream = BytesIO()
    with ZipFile(zip_stream, "w", ZIP_DEFLATED) as archive:
        archive.writestr("result.md", "# MinerU Result\n\nBody\n")
        archive.writestr("images/page_1.png", b"png")
    payload = zip_stream.getvalue()

    class FakeResponse:
        headers = {"Content-Type": "application/zip"}

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

    monkeypatch.setenv("MINERU_API_KEY", "mineru-token")
    monkeypatch.setenv("MINERU_BASE_URL", "https://mineru.example.test")
    monkeypatch.delenv("MINERU_API_ENDPOINT", raising=False)
    monkeypatch.setattr("edp.main_parser.request.urlopen", fake_urlopen)

    result = parse_main_document(source, "mineru", tmp_path / "artifacts")

    assert result.parser == "mineru"
    assert result.markdown == "# MinerU Result\n\nBody\n"
    assert result.warnings == []
    assert captured["url"] == "https://mineru.example.test/file_parse"
    assert captured["authorization"] == "mineru-token"
    assert captured["accept"] == "application/zip, application/json"
    assert b'name="files"; filename="source.docx"' in captured["body"]
    assert b'name="file"; filename="source.docx"' not in captured["body"]
    assert b'name="return_images"\r\n\r\ntrue' in captured["body"]
    assert b'name="response_format_zip"\r\n\r\ntrue' in captured["body"]
    assert (tmp_path / "artifacts" / "source.response.zip").exists()
    assert (tmp_path / "artifacts" / "zip" / "result.md").read_text(encoding="utf-8") == result.markdown
    assert (tmp_path / "artifacts" / "zip" / "images" / "page_1.png").read_bytes() == b"png"
    assert result.artifacts["parser_zip"] == tmp_path / "artifacts" / "source.response.zip"
    assert result.artifacts["parser_zip_dir"] == tmp_path / "artifacts" / "zip"
    assert result.artifacts["clean_markdown"] == tmp_path / "artifacts" / "zip" / "result.md"


def test_mineru_parser_rejects_unsafe_zip_member_paths(
    tmp_path: Path, monkeypatch
) -> None:
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

    monkeypatch.setenv("MINERU_API_KEY", "mineru-token")
    monkeypatch.setattr("edp.main_parser.request.urlopen", lambda *_args, **_kwargs: FakeResponse())

    result = parse_main_document(source, "mineru-pipeline", tmp_path / "artifacts")

    assert result.markdown == ""
    assert result.warnings == [
        "MinerU zip response could not be unpacked: unsafe zip member path: ../escape.md"
    ]
    assert not (tmp_path / "escape.md").exists()


def test_mineru_vlm_engine_parser_sends_vlm_backend(
    tmp_path: Path, monkeypatch
) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")

    zip_stream = BytesIO()
    with ZipFile(zip_stream, "w", ZIP_DEFLATED) as archive:
        archive.writestr("result.md", "# VLM Result\n")
    payload = zip_stream.getvalue()

    class FakeResponse:
        headers = {"Content-Type": "application/zip"}

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def read(self) -> bytes:
            return payload

    captured = {}

    def fake_urlopen(request, timeout):
        captured["body"] = request.data
        return FakeResponse()

    monkeypatch.setenv("MINERU_API_KEY", "mineru-token")
    monkeypatch.delenv("MINERU_BACKEND", raising=False)
    monkeypatch.setattr("edp.main_parser.request.urlopen", fake_urlopen)

    result = parse_main_document(source, "mineru-vlm-engine", tmp_path / "artifacts")

    assert result.parser == "mineru-vlm-engine"
    assert b'name="backend"\r\n\r\nvlm-engine' in captured["body"]


def test_mineru_hybrid_engine_missing_authorization_uses_variant_label(
    tmp_path: Path, monkeypatch
) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    monkeypatch.delenv("MINERU_AUTHORIZATION", raising=False)
    monkeypatch.delenv("MINERU_API_KEY", raising=False)
    monkeypatch.delenv("DOCUMENT_CONVERTER_AUTHORIZATION", raising=False)

    result = parse_main_document(source, "mineru-hybrid-engine", tmp_path / "artifacts")

    assert result.parser == "mineru-hybrid-engine"
    assert result.markdown == ""
    assert result.warnings == ["MinerU Authorization key is not configured"]


def test_mineru_parser_reports_missing_authorization(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    monkeypatch.delenv("MINERU_AUTHORIZATION", raising=False)
    monkeypatch.delenv("MINERU_API_KEY", raising=False)
    monkeypatch.delenv("DOCUMENT_CONVERTER_AUTHORIZATION", raising=False)
    monkeypatch.delenv("MINERU_API_ENDPOINT", raising=False)

    result = parse_main_document(source, "mineru", tmp_path / "artifacts")

    assert result.parser == "mineru"
    assert result.markdown == ""
    assert result.warnings == ["MinerU Authorization key is not configured"]


def test_mineru_endpoint_defaults_to_local_deployment(monkeypatch) -> None:
    monkeypatch.delenv("MINERU_API_ENDPOINT", raising=False)
    monkeypatch.delenv("MINERU_BASE_URL", raising=False)

    assert _mineru_endpoint() == "http://127.0.0.1:8000/file_parse"


def test_mineru_fields_allow_backend_and_model_version_overrides(monkeypatch) -> None:
    from edp.main_parser import _mineru_fields

    monkeypatch.setenv("MINERU_BACKEND", "vlm-sglang-engine")
    monkeypatch.setenv("MINERU_MODEL_VERSION", "vlm")

    fields = _mineru_fields("hybrid-engine")

    assert ("backend", "vlm-sglang-engine") in fields
    assert ("model_version", "vlm") in fields


def test_legacy_mineru_parser_defaults_to_pipeline_backend(monkeypatch) -> None:
    from edp.main_parser import _mineru_fields

    monkeypatch.delenv("MINERU_BACKEND", raising=False)
    monkeypatch.delenv("MINERU_MODEL_VERSION", raising=False)

    fields = _mineru_fields("pipeline")

    assert ("backend", "pipeline") in fields
