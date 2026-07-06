from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class EmbeddedObject:
    ref: str
    filename: str
    type: str
    path: Path
    rel_id: str
    source_path: str
    position: int
    content_type: str | None = None
    kind: str = "attachment"
    resource_id: str | None = None
    relationship_type: str | None = None
    original_filename: str | None = None
    detected_mime: str | None = None
    extension: str = ""
    size_bytes: int = 0
    sha256: str = ""
    anchor: dict[str, object] = field(default_factory=dict)
    parse_policy: dict[str, object] = field(default_factory=dict)
    parse_status: dict[str, object] = field(default_factory=dict)
    risk: dict[str, object] = field(
        default_factory=lambda: {"risk_level": "unassessed", "flags": []}
    )
    related_parts: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class ExtractionResult:
    input_docx: Path
    work_dir: Path
    raw_original_path: Path
    embedded_dir: Path
    media_dir: Path | None = None
    objects: list[EmbeddedObject] = field(default_factory=list)
    images: list[EmbeddedObject] = field(default_factory=list)
    charts: list[EmbeddedObject] = field(default_factory=list)
    diagrams: list[EmbeddedObject] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class ParsedTable:
    sheet_name: str
    csv_path: Path
    json_path: Path
    row_count: int


@dataclass(frozen=True)
class ParsedPackage:
    ref: str
    package_dir: Path
    content_path: Path
    tables: list[ParsedTable] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    type: str = "attachment"


@dataclass(frozen=True)
class DocumentPackage:
    package_dir: Path
    content_path: Path
    manifest_path: Path
    warnings: list[str] = field(default_factory=list)
