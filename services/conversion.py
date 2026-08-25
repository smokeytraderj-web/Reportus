"""Convert report intermediates through locally installed office software."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


class ConversionError(RuntimeError):
    pass


def convert_pptx_to_pdf(source: Path, destination: Path) -> Path:
    """Convert a PPTX to PDF without silently accepting stale output."""

    if not source.is_file():
        raise ConversionError("The presentation intermediate is missing.")
    office = shutil.which("libreoffice") or shutil.which("soffice")
    if office is None:
        raise ConversionError("PowerPoint or LibreOffice is required for PDF conversion.")
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="reportus-convert-") as temporary:
        temp_dir = Path(temporary)
        profile_uri = (temp_dir / "office-profile").resolve().as_uri()
        completed = subprocess.run(
            [
                office,
                f"-env:UserInstallation={profile_uri}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(temp_dir),
                str(source),
            ],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
        produced = temp_dir / f"{source.stem}.pdf"
        if completed.returncode != 0 or not produced.is_file() or produced.stat().st_size < 100:
            raise ConversionError("Office conversion did not produce a valid PDF.")
        shutil.copy2(produced, destination)
    return destination
