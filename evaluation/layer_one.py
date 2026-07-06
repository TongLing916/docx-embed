from __future__ import annotations

from pathlib import Path


def embedded_object_recall(package_dir: str | Path, expected_count: int) -> float:
    manifest_path = Path(package_dir) / "manifest.json"
    if expected_count == 0:
        return 1.0
    if not manifest_path.exists():
        return 0.0

    import json

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    actual = len(manifest.get("content_map", {}).get("embedded_objects", []))
    return min(actual / expected_count, 1.0)
