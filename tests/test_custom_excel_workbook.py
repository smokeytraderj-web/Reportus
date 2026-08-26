"""Validation and rendering tests for custom branded workbooks."""

import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from generators.custom_excel_workbook import (
    CustomWorkbookError,
    build_custom_workbook,
    build_custom_workbook_preview,
    normalize_custom_workbook_payload,
)
from quality.output_qa import OutputInspector
from services.ycharts_catalog import YChartsCatalog, YChartsMetric


def _catalog() -> YChartsCatalog:
    return YChartsCatalog((
        YChartsMetric(
            "1 Month Total Returns (Monthly)", "total_monthly_return", "YCP",
            "metrics", ("Company",),
        ),
    ))


def _payload() -> dict[str, object]:
    return {
        "workbook_subtitle": "Live performance analysis",
        "sheets": [{
            "name": "Performance",
            "title": "Performance Comparison",
            "subtitle": "Monthly total return history",
            "source_note": "Live YCharts formulas",
            "columns": [
                {"header": "Date", "format": "date", "width": 16},
                {"header": "MSFT", "format": "percent", "width": 16},
            ],
            "rows": [{"cells": [
                {
                    "kind": "ycharts", "function": "YCDS", "security": "MSFT",
                    "metric_code": "total_monthly_return", "start_date": "2026-01-01",
                    "end_date": "2026-08-26", "metric_name": "Monthly dates",
                },
                {
                    "kind": "ycharts", "function": "YCS", "security": "MSFT",
                    "metric_code": "total_monthly_return", "start_date": "2026-01-01",
                    "end_date": "2026-08-26", "metric_name": "Monthly total return",
                },
            ]}],
            "charts": [{
                "type": "line", "title": "MSFT Monthly Total Return",
                "category_column": 0, "series_columns": [1], "max_rows": 24,
            }],
        }],
    }


class CustomExcelWorkbookTests(unittest.TestCase):
    def test_builds_branded_workbook_with_validated_ycharts_series(self) -> None:
        catalog = _catalog()
        plan = normalize_custom_workbook_payload(
            _payload(), report_title="Performance Review", catalog=catalog
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook_path = root / "custom.xlsx"
            preview_path = root / "preview.pdf"

            build_custom_workbook(plan, "Chart MSFT monthly total return for 2026.", workbook_path)
            build_custom_workbook_preview(plan, "Chart MSFT monthly total return for 2026.", preview_path)

            self.assertTrue(OutputInspector().inspect(workbook_path).approved)
            self.assertTrue(OutputInspector().inspect(preview_path).approved)
            workbook = load_workbook(workbook_path, data_only=False)
            try:
                sheet = workbook["Performance"]
                self.assertEqual(
                    sheet["A6"].value,
                    '=YCDS("MSFT","total_monthly_return","2026-01-01","2026-08-26")',
                )
                self.assertEqual(
                    sheet["B6"].value,
                    '=YCS("MSFT","total_monthly_return","2026-01-01","2026-08-26")',
                )
                self.assertEqual(sheet["A1"].fill.fgColor.rgb, "000A1224")
                self.assertEqual(len(sheet._charts), 1)
            finally:
                workbook.close()

    def test_rejects_invented_ycharts_code(self) -> None:
        payload = _payload()
        payload["sheets"][0]["rows"][0]["cells"][1]["metric_code"] = "invented_metric"

        with self.assertRaisesRegex(CustomWorkbookError, "not in the supplied YCharts reference"):
            normalize_custom_workbook_payload(
                payload, report_title="Performance Review", catalog=_catalog()
            )

    def test_rejects_unsafe_excel_formula(self) -> None:
        payload = _payload()
        payload["sheets"][0]["rows"][0]["cells"][1] = {
            "kind": "excel_formula", "formula": '=WEBSERVICE("https://example.com")'
        }

        with self.assertRaisesRegex(CustomWorkbookError, "unsupported function"):
            normalize_custom_workbook_payload(
                payload, report_title="Performance Review", catalog=_catalog()
            )


if __name__ == "__main__":
    unittest.main()
