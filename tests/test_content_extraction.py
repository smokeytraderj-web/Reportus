"""Tests for bounded local extraction with source provenance."""

import tempfile
import unittest
from pathlib import Path

from extraction.content import ContentExtractor, ExtractionError


class ContentExtractorTests(unittest.TestCase):
    def test_csv_rows_keep_source_locators(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "data.csv"
            path.write_text("Symbol,Return\nIVV,5.25%\n", encoding="utf-8")

            fragments = ContentExtractor().extract([path])

            self.assertEqual(fragments[1].locator, "row 2")
            self.assertEqual(fragments[1].text, "IVV | 5.25%")

    def test_size_limit_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "data.txt"
            path.write_text("123456", encoding="utf-8")

            with self.assertRaisesRegex(ExtractionError, "too large"):
                ContentExtractor(max_characters_per_file=5).extract([path])


if __name__ == "__main__":
    unittest.main()
