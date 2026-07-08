from __future__ import annotations

from pathlib import PurePosixPath, PureWindowsPath


EXTERNAL_LINK_PREFIXES = ("http://", "https://", "data:", "mailto:")


def is_external_or_absolute_path(path: str) -> bool:
    """Return true for links that must not be treated as package-relative files."""

    lowered = path.lower()
    return (
        lowered.startswith(EXTERNAL_LINK_PREFIXES)
        or path.startswith("#")
        or PurePosixPath(path).is_absolute()
        or PureWindowsPath(path).is_absolute()
    )
