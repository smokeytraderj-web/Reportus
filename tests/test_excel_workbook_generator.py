"""Synthetic validation for the clean GSWM holdings workbook."""

import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from generators.excel_workbook import (
    HoldingsWorkbookConfig,
    WorkbookBuildError,
    build_holdings_snapshot,
    build_holdings_workbook,
    load_holdings,
)
from quality.output_qa import OutputInspector


class ExcelWorkbookGeneratorTests(unittest.TestCase):
    def test_builds_styled_workbook_and_matching_snapshot(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "holdings.csv"
            source.write_text(
                "Description,Symbol,Quantity,Price,Value,% of Assets\n"
                "iShares Core S&P 500 ETF,IVV,100,675.25,67525,67.525%\n"
                "AT&T Inc,T,200,29.75,5950,5.95%\n",
                encoding="utf-8",
            )
            rows = load_holdings(source)
            config = HoldingsWorkbookConfig("Holdings Summary", "Synthetic test data")
            output = root / "holdings.xlsx"
            preview = root / "holdings.pdf"

            build_holdings_workbook(rows, config, output)
            build_holdings_snapshot(rows, config, preview)

            self.assertTrue(OutputInspector().inspect(output).approved)
            self.assertTrue(OutputInspector().inspect(preview).approved)
            workbook = load_workbook(output, data_only=False, rich_text=True)
            try:
                sheet = workbook["Holdings Summary"]
                self.assertFalse(sheet.sheet_view.showGridLines)
                self.assertEqual(sheet.freeze_panes, "A6")
                self.assertEqual(sheet["D6"].number_format, '$#,##0.00;[Red]($#,##0.00);-')
                self.assertEqual(sheet["E8"].value, "=SUM(E6:E7)")
                self.assertNotIn("LLC", sheet["A1"].value)
                self.assertIn("AT&T", str(sheet["A7"].value))
            finally:
                workbook.close()

    def test_rejects_ambiguous_percentage_scale(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "holdings.csv"
            source.write_text(
                "Description,Symbol,Quantity,Price,Value,% of Assets\n"
                "Sample,SAM,1,10,10,25\n",
                encoding="utf-8",
            )
            with self.assertRaisesRegex(WorkbookBuildError, "ambiguous"):
                load_holdings(source)

    def test_does_not_copy_preheader_account_text(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "holdings.csv"
            source.write_text(
                "Account 123456,,,,,\n"
                "Description,Symbol,Quantity,Price,Value,% of Assets\n"
                "Sample,SAM,1,10,10,100%\n",
                encoding="utf-8",
            )
            rows = load_holdings(source)
            self.assertEqual(rows[0].company, "Sample")


if __name__ == "__main__":
    unittest.main()
