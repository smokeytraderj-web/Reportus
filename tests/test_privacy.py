"""Synthetic tests for the global privacy gate."""

import tempfile
import unittest
from pathlib import Path

from security.privacy import PrivacyScanner, ProhibitedCategory


class PrivacyScannerTests(unittest.TestCase):
    def test_typed_revision_uses_same_privacy_policy(self) -> None:
        scanner = PrivacyScanner()

        approved = scanner.scan_text("Shorten the second bullet on the risk slide.")
        rejected = scanner.scan_text("Send it to analyst@example.com")

        self.assertTrue(approved.approved)
        self.assertFalse(rejected.approved)
        self.assertEqual(
            rejected.findings[0].category, ProhibitedCategory.EMAIL_ADDRESS
        )

    def setUp(self) -> None:
        self.scanner = PrivacyScanner()
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def write(self, name: str, text: str) -> Path:
        path = self.root / name
        path.write_text(text, encoding="utf-8")
        return path

    def test_clean_financial_data_is_approved(self) -> None:
        path = self.write(
            "clean.csv",
            "Client,Symbol,Value,Weight\nRon Bloom,IVV,5046150.25,47.44%\n",
        )

        result = self.scanner.scan_files([path])

        self.assertTrue(result.approved)
        self.assertEqual(result.findings, ())

    def test_any_finding_rejects_complete_batch(self) -> None:
        clean = self.write("clean.csv", "Symbol,Value\nIVV,1000\n")
        unsafe = self.write(
            "unsafe.txt",
            "Client email: client@example.com\nAccount ending in 7220\n",
        )

        result = self.scanner.scan_files([clean, unsafe])

        self.assertFalse(result.approved)
        self.assertEqual(
            {finding.category for finding in result.findings},
            {ProhibitedCategory.EMAIL_ADDRESS, ProhibitedCategory.ACCOUNT_IDENTIFIER},
        )

    def test_finding_does_not_repeat_sensitive_value(self) -> None:
        sensitive_value = "client@example.com"
        path = self.write("unsafe.txt", f"Email: {sensitive_value}\n")

        result = self.scanner.scan_files([path])
        serialized = repr(result)

        self.assertFalse(result.approved)
        self.assertNotIn(sensitive_value, serialized)

    def test_uninspectable_format_fails_closed(self) -> None:
        path = self.write("legacy.xls", "not a real workbook")

        result = self.scanner.scan_files([path])

        self.assertFalse(result.approved)
        self.assertIn("not safely inspectable", result.errors[0])

    def test_account_label_without_identifier_is_allowed(self) -> None:
        path = self.write("clean.txt", "Account type: Traditional IRA\n")

        result = self.scanner.scan_files([path])

        self.assertTrue(result.approved)

    def test_table_headers_and_masked_identifiers_are_rejected(self) -> None:
        path = self.write(
            "table.txt",
            "Account Holder Registration Account DOB\nClient One IRA … 7220 1/1/1950\n",
        )

        result = self.scanner.scan_files([path])

        self.assertFalse(result.approved)
        self.assertEqual(
            {finding.category for finding in result.findings},
            {ProhibitedCategory.ACCOUNT_IDENTIFIER, ProhibitedCategory.DATE_OF_BIRTH},
        )


if __name__ == "__main__":
    unittest.main()
