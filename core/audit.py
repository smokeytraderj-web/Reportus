"""Temporary, privacy-approved evidence shown beside a prepared report."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class AuditSource:
    """One approved input and the report areas that rely on it."""

    name: str
    role: str


@dataclass(frozen=True, slots=True)
class AuditCitation:
    """A section-level source locator supplied by a deterministic parser or model."""

    section: str
    locator: str


@dataclass(frozen=True, slots=True)
class ReportAudit:
    """Ephemeral review metadata that is never written into the final artifact."""

    report_type: str
    sections: tuple[str, ...]
    sources: tuple[AuditSource, ...]
    citations: tuple[AuditCitation, ...] = ()

    @classmethod
    def from_staged_inputs(
        cls,
        *,
        report_type: str,
        sections: tuple[str, ...],
        staged: dict[str, tuple[Path, ...]],
        slot_labels: dict[str, str],
        citations: tuple[AuditCitation, ...] = (),
    ) -> "ReportAudit":
        sources = tuple(
            AuditSource(path.name, slot_labels.get(slot, slot.replace("_", " ").title()))
            for slot, paths in staged.items()
            for path in paths
        )
        return cls(report_type, sections, sources, citations)
