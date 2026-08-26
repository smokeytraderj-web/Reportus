"""Report intake contracts used by the desktop shell."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from config.settings import PROJECT_ROOT


@dataclass(frozen=True, slots=True)
class UploadRequirement:
    """One skill-defined upload slot."""

    requirement_id: str
    label: str
    description: str
    accepted_extensions: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class ReportField:
    """One concise user-supplied value required by a report generator."""

    field_id: str
    label: str
    placeholder: str
    required: bool = True
    default: str = ""
    multiline: bool = False


@dataclass(frozen=True, slots=True)
class ReportWorkflow:
    """UI-facing intake contract for one report skill."""

    skill_id: str
    title: str
    subtitle: str
    fields: tuple[ReportField, ...]
    required_uploads: tuple[UploadRequirement, ...]
    optional_uploads: tuple[UploadRequirement, ...]


def _requirements(items: list[dict[str, str]]) -> tuple[UploadRequirement, ...]:
    return tuple(
        UploadRequirement(
            requirement_id=item["id"],
            label=item["label"],
            description=item["description"],
            accepted_extensions=tuple(extension.lower() for extension in item["accepted_extensions"]),
        )
        for item in items
    )


def _fields(items: list[dict[str, object]]) -> tuple[ReportField, ...]:
    return tuple(
        ReportField(
            field_id=str(item["id"]),
            label=str(item["label"]),
            placeholder=str(item.get("placeholder", "")),
            required=bool(item.get("required", True)),
            default=str(item.get("default", "")),
            multiline=bool(item.get("multiline", False)),
        )
        for item in items
    )


def load_workflows(path: Path | None = None) -> tuple[ReportWorkflow, ...]:
    """Load the small UI manifest without loading any skill instructions."""

    manifest = path or PROJECT_ROOT / "config" / "report_workflows.json"
    payload = json.loads(manifest.read_text(encoding="utf-8"))
    return tuple(
        ReportWorkflow(
            skill_id=item["skill_id"],
            title=item["title"],
            subtitle=item["subtitle"],
            fields=_fields(item.get("fields", [])),
            required_uploads=_requirements(item.get("required_uploads", [])),
            optional_uploads=_requirements(item.get("optional_uploads", [])),
        )
        for item in payload["workflows"]
    )
