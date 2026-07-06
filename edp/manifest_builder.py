from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any


PIPELINE_VERSION = "0.1.0"


def build_manifest(
    package_dir: str | Path,
    status: str,
    warnings: list[str],
    content_map: dict[str, Any],
) -> dict[str, Any]:
    package_path = Path(package_dir)
    content_path = package_path / "structured" / "content.md"
    embedded_objects = content_map.get("embedded_objects", [])
    images = content_map.get("images", [])
    charts = content_map.get("charts", [])
    diagrams = content_map.get("diagrams", [])

    return {
        "doc_id": package_path.name,
        "original_file": "raw/original.docx",
        "parse_timestamp": datetime.now(UTC).isoformat(),
        "pipeline_version": PIPELINE_VERSION,
        "parse_status": status,
        "content_stats": {
            "embedded_object_count": len(embedded_objects),
            "image_count": len(images),
            "chart_count": len(charts),
            "diagram_count": len(diagrams),
            "content_md_bytes": content_path.stat().st_size if content_path.exists() else 0,
        },
        "parse_warnings": warnings,
        "content_map": content_map,
    }
