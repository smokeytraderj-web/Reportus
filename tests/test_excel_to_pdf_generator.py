"""Contract tests for the Excel-to-PDF stock-review generator."""

import datetime as dt
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from generators.excel_to_pdf import (
    ExcelToPDFError,
    ReviewDefinition,
    StockReviewConfig,
    build_stock_review_html,
    find_chromium,
    load_stock_reviews,
)


def _workbook(path: Path, *, note: str = "Buy - Strong earnings outlook.") -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "August"
    sheet.append([None, "As-Of Date", dt.date(2026, 8, 25)])
    sheet.append([None, "Default Rec. Date", dt.date(2026, 6, 1)])
    sheet.append([None, "S&P 500 Total Return (benchmark)", None, None, None, .0712])
    sheet.append([None, "Technology"])
    sheet.append([None, "Example Holdings", "EXM", "data", None, .1234, .0522, note, 125.678, "YCharts"])
    sheet.append([None, "Financials"])
    sheet.append([None, "Sample Bank", "SBK", "data", None, -.0123, -.0835, "Hold - Awaiting clarity.", 82.2, "Internal"])
    workbook.save(path)


def _config() -> StockReviewConfig:
    return StockReviewConfig(
        client_name="Sample Household",
        report_title="Recommendations Review",
        period_label="2026 Review",
        reviews=(ReviewDefinition("August", "Technology and Financials"),),
        source_label="YCharts workbook, uploaded August 25, 2026",
    )


class ExcelToPDFGeneratorTests(unittest.TestCase):
    def test_loads_strict_contract_and_builds_deterministic_html(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "report.xlsx"
            _workbook(source)

            reviews = load_stock_reviews(source, _config())
            document = build_stock_review_html(reviews, _config())

            self.assertEqual(len(reviews), 1)
            self.assertEqual(sum(len(section.rows) for section in reviews[0].sections), 2)
            self.assertEqual(document.count('<div class="page'), 4)
            self.assertIn("+12.34%", document)
            self.assertIn("$125.68", document)
            self.assertIn("Source: YCharts", document)
            self.assertIn("Gottfried &amp; Somberg Wealth Management", document)
            self.assertNotIn("LLC", document)

    def test_rejects_missing_rating_instead_of_inventing_one(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "report.xlsx"
            _workbook(source, note="Strong earnings outlook.")

            with self.assertRaisesRegex(ExcelToPDFError, "explicit rating"):
                load_stock_reviews(source, _config())

    def test_escapes_workbook_text(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "report.xlsx"
            _workbook(source, note="Buy - A&B <script>alert(1)</script>")

            document = build_stock_review_html(load_stock_reviews(source, _config()), _config())

            self.assertIn("A&amp;B &lt;script&gt;alert(1)&lt;/script&gt;", document)
            self.assertNotIn("<script>alert", document)

    def test_browser_error_is_clear(self) -> None:
        with self.assertRaisesRegex(ExcelToPDFError, "Chrome or Chromium"):
            find_chromium(Path("/definitely/not/a/browser"))


if __name__ == "__main__":
    unittest.main()
