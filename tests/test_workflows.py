"""Tests for report workflow configuration."""

import unittest

from core.workflows import load_workflows


class WorkflowTests(unittest.TestCase):
    def test_initial_workflows_match_menu(self) -> None:
        workflows = load_workflows()

        self.assertEqual(
            [workflow.title for workflow in workflows],
            ["Client Deck", "Excel to PDF", "Custom Excel Workbook", "PowerPoint Deck"],
        )

    def test_only_custom_excel_can_run_without_an_upload(self) -> None:
        for workflow in load_workflows():
            with self.subTest(workflow=workflow.skill_id):
                if workflow.skill_id == "excel-workbook-builder":
                    self.assertFalse(workflow.required_uploads)
                    self.assertTrue(workflow.optional_uploads)
                else:
                    self.assertTrue(workflow.required_uploads)

    def test_upload_ids_are_unique_within_workflow(self) -> None:
        for workflow in load_workflows():
            ids = [
                requirement.requirement_id
                for requirement in workflow.required_uploads + workflow.optional_uploads
            ]
            with self.subTest(workflow=workflow.skill_id):
                self.assertEqual(len(ids), len(set(ids)))

    def test_field_ids_are_unique_within_workflow(self) -> None:
        for workflow in load_workflows():
            ids = [field.field_id for field in workflow.fields]
            with self.subTest(workflow=workflow.skill_id):
                self.assertEqual(len(ids), len(set(ids)))

    def test_initial_workflows_collect_report_titles(self) -> None:
        for workflow in load_workflows():
            with self.subTest(workflow=workflow.skill_id):
                self.assertTrue(workflow.fields)

    def test_every_upload_declares_file_types(self) -> None:
        for workflow in load_workflows():
            for requirement in workflow.required_uploads + workflow.optional_uploads:
                with self.subTest(workflow=workflow.skill_id, slot=requirement.requirement_id):
                    self.assertTrue(requirement.accepted_extensions)

    def test_client_deck_requires_holdings_and_attribution_uploads(self) -> None:
        workflow = next(
            item for item in load_workflows() if item.skill_id == "client-deck-builder"
        )
        requirements = {item.requirement_id: item for item in workflow.required_uploads}

        self.assertIn("holdings", requirements)
        self.assertIn("attribution", requirements)
        self.assertEqual(requirements["holdings"].accepted_extensions, (".xlsx", ".xlsm"))

    def test_custom_excel_collects_a_multiline_request(self) -> None:
        workflow = next(
            item for item in load_workflows() if item.skill_id == "excel-workbook-builder"
        )
        request = next(field for field in workflow.fields if field.field_id == "workbook_request")

        self.assertTrue(request.multiline)
        self.assertIn("YCharts", request.placeholder)
        optional = {item.requirement_id: item for item in workflow.optional_uploads}
        self.assertEqual(
            optional["ycharts_reference"].accepted_extensions, (".xlsx", ".xlsm")
        )


if __name__ == "__main__":
    unittest.main()
