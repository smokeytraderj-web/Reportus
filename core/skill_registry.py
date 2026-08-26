"""Discover and validate Reporticles skills without loading every skill body."""

from __future__ import annotations

import json
import re
from pathlib import Path

from config.settings import settings
from core.models import SkillDefinition, SkillKind, SkillValidationIssue


_FRONTMATTER = re.compile(r"\A---\s*\n(?P<body>.*?)\n---\s*\n", re.DOTALL)
_FIELD = re.compile(r"^(?P<key>[a-zA-Z0-9_-]+):\s*(?P<value>.*)$")


def _frontmatter_fields(skill_md: Path) -> dict[str, str]:
    """Read only the metadata needed for routing from a SKILL.md file."""

    text = skill_md.read_text(encoding="utf-8")
    match = _FRONTMATTER.match(text)
    if not match:
        raise ValueError("missing YAML frontmatter")

    fields: dict[str, str] = {}
    active_key: str | None = None
    folded: list[str] = []

    def finish_folded() -> None:
        nonlocal active_key, folded
        if active_key is not None:
            fields[active_key] = " ".join(part.strip() for part in folded).strip()
        active_key = None
        folded = []

    for line in match.group("body").splitlines():
        field_match = _FIELD.match(line)
        if field_match and not line.startswith((" ", "\t")):
            finish_folded()
            key = field_match.group("key")
            value = field_match.group("value").strip()
            if value in {">", "|"}:
                active_key = key
            else:
                fields[key] = value.strip("'\"")
        elif active_key is not None and line.strip():
            folded.append(line)
    finish_folded()
    return fields


class SkillRegistry:
    """Expose validated menu skills and on-demand skill instructions."""

    def __init__(self, skills_root: Path | None = None, manifest: Path | None = None):
        self.skills_root = skills_root or settings.skills_root
        self.manifest = manifest or settings.skill_manifest
        self._definitions: dict[str, SkillDefinition] = {}
        self._issues: list[SkillValidationIssue] = []

    @property
    def issues(self) -> tuple[SkillValidationIssue, ...]:
        return tuple(self._issues)

    def discover(self) -> tuple[SkillDefinition, ...]:
        """Validate configured metadata without retaining SKILL.md bodies."""

        self._definitions.clear()
        self._issues.clear()
        payload = json.loads(self.manifest.read_text(encoding="utf-8"))

        for item in payload.get("skills", []):
            if not item.get("enabled", False):
                continue
            skill_id = item["id"]
            skill_path = self.skills_root / skill_id
            skill_md = skill_path / "SKILL.md"
            if not skill_md.is_file():
                self._issues.append(SkillValidationIssue(skill_id, "SKILL.md is missing"))
                continue
            try:
                fields = _frontmatter_fields(skill_md)
                name = fields["name"]
                description = fields["description"]
                kind = SkillKind(item["kind"])
            except (KeyError, ValueError, json.JSONDecodeError) as exc:
                self._issues.append(SkillValidationIssue(skill_id, str(exc)))
                continue
            if name != skill_id:
                self._issues.append(
                    SkillValidationIssue(skill_id, f"frontmatter name is {name!r}")
                )
                continue
            self._definitions[skill_id] = SkillDefinition(
                skill_id=skill_id,
                name=name,
                label=item["label"],
                description=description,
                kind=kind,
                path=skill_path,
                menu_order=item.get("menu_order"),
            )

        return tuple(self._definitions.values())

    def menu_skills(self) -> tuple[SkillDefinition, ...]:
        """Return report skills in stable display order."""

        if not self._definitions:
            self.discover()
        reports = (
            skill for skill in self._definitions.values() if skill.kind is SkillKind.REPORT
        )
        return tuple(sorted(reports, key=lambda skill: skill.menu_order or 999))

    def load_instructions(self, skill_id: str) -> str:
        """Load one selected skill body on demand."""

        if not self._definitions:
            self.discover()
        try:
            skill = self._definitions[skill_id]
        except KeyError as exc:
            raise KeyError(f"unknown or invalid skill: {skill_id}") from exc
        return (skill.path / "SKILL.md").read_text(encoding="utf-8")
