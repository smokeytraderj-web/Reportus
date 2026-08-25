"""Tests for deterministic report-input validation."""

import tempfile
import unittest
from pathlib import Path

from core.workflows import ReportWorkflow, UploadRequirement
from validation.inputs import InputValidator


class InputValidatorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.requirement = UploadRequirement("data", "Source data", "Data", (".csv",))
        self.workflow = ReportWorkflow("test", "Test", "Test", (), (self.requirement,), ())

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_valid_csv_is_approved(self) -> None:
        source = self.root / "data.csv"
        source.write_text("Symbol,Value\nIVV,100.00\n", encoding="utf-8")

        result = InputValidator().validate(self.workflow, {"data": (source,)})

        self.assertTrue(result.approved)

    def test_wrong_file_type_is_rejected(self) -> None:
        source = self.root / "data.pdf"
        source.write_bytes(b"not a pdf")

        result = InputValidator().validate(self.workflow, {"data": (source,)})

        self.assertFalse(result.approved)
        self.assertEqual(result.errors[0].code, "file_type_not_allowed")

    def test_missing_required_slot_is_rejected(self) -> None:
        result = InputValidator().validate(self.workflow, {})

        self.assertFalse(result.approved)
        self.assertEqual(result.errors[0].code, "required_upload_missing")

    def test_empty_file_is_rejected(self) -> None:
        source = self.root / "data.csv"
        source.touch()

        result = InputValidator().validate(self.workflow, {"data": (source,)})

        self.assertFalse(result.approved)
        self.assertEqual(result.errors[0].code, "file_empty")

    def test_json_structure_is_supported(self) -> None:
        source = self.root / "content.json"
        source.write_text('{"slides": []}', encoding="utf-8")
        requirement = UploadRequirement("data", "Source data", "Data", (".json",))
        workflow = ReportWorkflow("test", "Test", "Test", (), (requirement,), ())

        result = InputValidator().validate(workflow, {"data": (source,)})

        self.assertTrue(result.approved)


if __name__ == "__main__":
    unittest.main()
