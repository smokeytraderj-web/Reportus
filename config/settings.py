"""Environment, model, and system configuration."""

from dataclasses import dataclass
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


@dataclass(frozen=True, slots=True)
class Settings:
    """Application paths and conservative runtime defaults."""

    project_root: Path = PROJECT_ROOT
    skills_root: Path = PROJECT_ROOT / "skills"
    skill_manifest: Path = PROJECT_ROOT / "config" / "skills.json"
    session_root_name: str = "Reportus"
    normal_report_timeout_seconds: int = 7 * 60
    transient_retry_count: int = 1


settings = Settings()
