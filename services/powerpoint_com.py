"""Windows PowerPoint COM adapter for the supplied GSWM template skill."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
from pathlib import Path

from config.settings import PROJECT_ROOT, settings
from quality.output_qa import OutputInspector


class PowerPointBuildError(RuntimeError):
    pass


SKILL_ROOT = PROJECT_ROOT / "skills" / "powerpoint-deck-builder"


def _powershell() -> str:
    executable = shutil.which("pwsh") or shutil.which("powershell")
    if os.name != "nt" or executable is None:
        raise PowerPointBuildError("PowerPoint Deck generation requires Windows with Microsoft PowerPoint installed.")
    return executable


def _run(script: Path, arguments: list[str]) -> None:
    if not script.is_file():
        raise PowerPointBuildError("A required PowerPoint build resource is missing.")
    command = [_powershell(), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), *arguments]
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=settings.normal_report_timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise PowerPointBuildError("PowerPoint generation exceeded the seven-minute limit.") from exc
    if completed.returncode != 0:
        raise PowerPointBuildError("PowerPoint could not build the presentation from the approved template.")


def build_powerpoint_deck(content: dict[str, object], output_path: Path,
                          preview_path: Path, *, template_path: Path | None = None) -> tuple[Path, Path]:
    template = template_path or SKILL_ROOT / "GSWM_template.pptx"
    if not template.is_file() or template.suffix.lower() != ".pptx":
        raise PowerPointBuildError("Use a valid .pptx template.")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    preview_path.parent.mkdir(parents=True, exist_ok=True)
    content_path = output_path.parent / "content.json"
    content_path.write_text(json.dumps(content, indent=2, ensure_ascii=False), encoding="utf-8")
    _run(
        SKILL_ROOT / "builddeck.ps1",
        ["-ContentPath", str(content_path), "-OutputPath", str(output_path), "-TemplatePath", str(template)],
    )
    deck_qa = OutputInspector().inspect(output_path)
    if not deck_qa.approved:
        raise PowerPointBuildError("The generated presentation failed integrity checks.")
    _run(
        PROJECT_ROOT / "services" / "windows" / "export_powerpoint_pdf.ps1",
        ["-InputPath", str(output_path), "-OutputPath", str(preview_path)],
    )
    preview_qa = OutputInspector().inspect(preview_path)
    if not preview_qa.approved or preview_qa.page_or_sheet_count != deck_qa.page_or_sheet_count:
        raise PowerPointBuildError("The presentation preview failed page-count checks.")
    return output_path, preview_path
