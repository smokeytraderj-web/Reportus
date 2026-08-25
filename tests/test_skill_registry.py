"""Tests for skill discovery and lazy instruction loading."""

import unittest

from core.skill_registry import SkillRegistry


class SkillRegistryTests(unittest.TestCase):
    def test_initial_menu_order(self) -> None:
        registry = SkillRegistry()

        self.assertEqual(
            [skill.label for skill in registry.menu_skills()],
            ["Client Deck", "Excel to PDF", "Excel Workbook", "PowerPoint Deck"],
        )
        self.assertEqual(registry.issues, ())

    def test_support_capability_is_not_in_menu(self) -> None:
        registry = SkillRegistry()

        menu_ids = {skill.skill_id for skill in registry.menu_skills()}

        self.assertNotIn("ycharts-performance-charts", menu_ids)

    def test_selected_skill_loads_on_demand(self) -> None:
        registry = SkillRegistry()

        instructions = registry.load_instructions("client-deck-builder")

        self.assertIn("name: client-deck-builder", instructions)
        self.assertNotIn("name: template-pdf-report", instructions)


if __name__ == "__main__":
    unittest.main()
