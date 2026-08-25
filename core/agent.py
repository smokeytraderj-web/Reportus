"""Central router and supervisor for the agent system."""

from core.skill_registry import SkillRegistry


class ReportusAgent:
    """Thin coordinator that activates exactly one report skill per session."""

    def __init__(self, registry: SkillRegistry | None = None):
        self.registry = registry or SkillRegistry()

    def available_reports(self) -> tuple[tuple[str, str], ...]:
        """Return stable menu identifiers and labels."""

        return tuple((skill.skill_id, skill.label) for skill in self.registry.menu_skills())
