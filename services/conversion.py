"""Convert report intermediates through locally installed office software."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from config.settings import PROJECT_ROOT, settings


class ConversionError(RuntimeError):
    pass


def convert_pptx_to_pdf(source: Path, destination: Path) -> Path:
    """Convert a PPTX to PDF without silently accepting stale output."""

    if not source.is_file():
        raise ConversionError("The presentation intermediate is missing.")
    office = shutil.which("libreoffice") or shutil.which("soffice")
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="reporticles-convert-") as temporary:
        temp_dir = Path(temporary)
        produced = temp_dir / "converted.pdf"
        if office is not None:
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
            libreoffice_output = temp_dir / f"{source.stem}.pdf"
            if libreoffice_output.is_file():
                shutil.move(libreoffice_output, produced)
        else:
            powershell = shutil.which("pwsh") or shutil.which("powershell")
            script = PROJECT_ROOT / "services" / "windows" / "export_powerpoint_pdf.ps1"
            if os.name != "nt" or powershell is None or not script.is_file():
                raise ConversionError("Microsoft PowerPoint or LibreOffice is required for PDF conversion.")
            try:
                completed = subprocess.run(
                    [
                        powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
                        "-InputPath", str(source.resolve()), "-OutputPath", str(produced.resolve()),
                    ],
                    capture_output=True,
                    text=True,
                    timeout=settings.normal_report_timeout_seconds,
                    check=False,
                )
            except subprocess.TimeoutExpired as exc:
                raise ConversionError("PowerPoint PDF conversion exceeded the seven-minute limit.") from exc
        if completed.returncode != 0 or not produced.is_file() or produced.stat().st_size < 100:
            raise ConversionError("Office conversion did not produce a valid PDF.")
        shutil.copy2(produced, destination)
    return destination
