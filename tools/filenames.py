"""Safe Windows-compatible output naming and automatic versioning."""

from __future__ import annotations

import re
from pathlib import Path


_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}


def sanitize_filename(filename: str, fallback: str = "Reportus_Report") -> str:
    """Return a portable filename while preserving a supplied extension."""

    candidate = _INVALID.sub("_", filename).strip().rstrip(". ")
    if not candidate:
        candidate = fallback
    path = Path(candidate)
    if path.stem.upper() in _RESERVED:
        candidate = f"_{candidate}"
    if len(candidate) > 180:
        suffix = path.suffix[:20]
        candidate = candidate[: 180 - len(suffix)].rstrip(". ") + suffix
    return candidate


def versioned_output_path(directory: Path, filename: str) -> Path:
    """Choose a non-existing output path without overwriting prior reports."""

    safe_name = sanitize_filename(filename)
    first = directory / safe_name
    if not first.exists():
        return first
    path = Path(safe_name)
    for version in range(2, 10_000):
        candidate = directory / f"{path.stem}_v{version}{path.suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError("Unable to choose an available report filename.")
