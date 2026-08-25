"""Temporary report-session storage with privacy-first cleanup."""

from __future__ import annotations

import shutil
import tempfile
import uuid
from dataclasses import dataclass
from pathlib import Path

from config.settings import settings


def default_session_root() -> Path:
    """Return a narrow application-owned directory in the OS temp area."""

    return Path(tempfile.gettempdir()) / settings.session_root_name / "sessions"


@dataclass(slots=True)
class ReportSession:
    """Own all non-final files for one report run."""

    session_id: str
    path: Path
    _closed: bool = False

    @classmethod
    def create(cls, root: Path | None = None) -> "ReportSession":
        session_root = root or default_session_root()
        session_root.mkdir(parents=True, exist_ok=True)
        session_id = uuid.uuid4().hex
        path = session_root / session_id
        path.mkdir(mode=0o700)
        for child in ("uploads", "working", "preview"):
            (path / child).mkdir()
        return cls(session_id=session_id, path=path)

    @property
    def uploads(self) -> Path:
        return self.path / "uploads"

    @property
    def working(self) -> Path:
        return self.path / "working"

    @property
    def preview(self) -> Path:
        return self.path / "preview"

    @property
    def closed(self) -> bool:
        return self._closed

    def cleanup(self) -> None:
        """Irrecoverably remove the application-owned session directory."""

        if self._closed:
            return
        if self.path.parent == self.path or self.path.name != self.session_id:
            raise RuntimeError("refusing to clean an invalid session path")
        shutil.rmtree(self.path, ignore_errors=False)
        self._closed = True

    def __enter__(self) -> "ReportSession":
        return self

    def __exit__(self, exc_type: object, exc: object, traceback: object) -> None:
        self.cleanup()


def purge_abandoned_sessions(root: Path | None = None) -> int:
    """Delete crash leftovers from the application-owned session root."""

    session_root = root or default_session_root()
    if not session_root.exists():
        return 0
    removed = 0
    for path in session_root.iterdir():
        if path.is_dir() and len(path.name) == 32 and all(char in "0123456789abcdef" for char in path.name):
            shutil.rmtree(path)
            removed += 1
    return removed
