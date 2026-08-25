"""Typed models shared across Reportus subsystems."""

from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path


class SkillKind(StrEnum):
    """How a skill is exposed by Reportus."""

    REPORT = "report"
    CAPABILITY = "capability"


@dataclass(frozen=True, slots=True)
class SkillDefinition:
    """Validated metadata for one lazily loaded skill."""

    skill_id: str
    name: str
    label: str
    description: str
    kind: SkillKind
    path: Path
    menu_order: int | None = None


@dataclass(frozen=True, slots=True)
class SkillValidationIssue:
    """A blocking reason a skill cannot be activated."""

    skill_id: str
    message: str
