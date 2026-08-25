"""Tests for isolated report execution and final-output retention."""

import datetime as dt
import json
import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from pypdf import PdfWriter

from services.report_runner import ReportRunRequest, ReportRunner


def _make_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Review"
    sheet.append([None, "As-Of Date", dt.date(2026, 8, 25)])
    sheet.append([None, "Default Rec. Date", dt.date(2026, 8, 1)])
    sheet.append([None, "S&P 500 Total Return (benchmark)", None, None, None, .02])
    sheet.append([None, "Core"])
    sheet.append([None, "Sample", "SAM", "data", None, .04, .02, "Buy - Test", 10.0, "YCharts"])
    workbook.save(path)


def _make_template(path: Path) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=792, height=612)
    with path.open("wb") as stream:
        writer.write(stream)


class ReportRunnerTests(unittest.TestCase):
    def test_prepares_and_finalizes_powerpoint_json_package(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "content.json"
            source.write_text(json.dumps({
                "slides": [{"title": "The portfolio remains balanced", "bullets": ["Risk remains within the agreed range."]}]
            }), encoding="utf-8")

            def fake_builder(content, artifact: Path, preview: Path, *, template_path=None):
                self.assertEqual(content["title"], "Portfolio Review")
                template = Path("skills/powerpoint-deck-builder/GSWM_template.pptx")
                shutil.copy2(template, artifact)
                writer = PdfWriter()
                writer.add_blank_page(width=960, height=540)
                writer.add_blank_page(width=960, height=540)
                with preview.open("wb") as stream:
                    writer.write(stream)
                return artifact, preview

            runner = ReportRunner(
                session_root=root / "sessions", powerpoint_builder=fake_builder
            )
            request = ReportRunRequest(
                "powerpoint-deck-builder",
                {"content": (source,)},
                {"report_title": "Portfolio Review"},
            )

            prepared = runner.prepare(request)

            self.assertEqual(prepared.artifact_path.suffix, ".pptx")
            self.assertEqual(prepared.preview_path.suffix, ".pdf")
            result = runner.finalize(prepared, root / "final")
            self.assertEqual(result.output_path.name, "Portfolio Review.pptx")

    def test_prepares_and_finalizes_excel_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "holdings.csv"
            source.write_text(
                "Description,Symbol,Quantity,Price,Value,% of Assets\n"
                "Sample Holding,SAM,10,20,200,100%\n",
                encoding="utf-8",
            )
            request = ReportRunRequest(
                "excel-workbook-builder",
                {"source_data": (source,)},
                {"report_title": "Holdings Summary", "source_label": "Synthetic data"},
            )
            runner = ReportRunner(session_root=root / "sessions")

            prepared = runner.prepare(request)

            self.assertEqual(prepared.artifact_path.suffix, ".xlsx")
            self.assertEqual(prepared.preview_path.suffix, ".pdf")
            self.assertTrue(prepared.artifact_path.is_file())
            self.assertTrue(prepared.preview_path.is_file())
            result = runner.finalize(prepared, root / "final")
            self.assertEqual(result.output_path.name, "Holdings Summary.xlsx")
            self.assertTrue(result.output_path.is_file())

    def test_prepare_keeps_preview_until_finalize(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = root / "source.xlsx"
            template = root / "template.pdf"
            sessions = root / "sessions"
            _make_workbook(workbook)
            _make_template(template)

            def fake_builder(source: Path, destination: Path, config) -> Path:
                writer = PdfWriter()
                writer.add_blank_page(width=792, height=612)
                with destination.open("wb") as stream:
                    writer.write(stream)
                return destination

            request = ReportRunRequest(
                "template-pdf-report",
                {"spreadsheet": (workbook,), "template": (template,)},
                {"client_name": "Sample", "period_label": "2026", "report_title": "Review", "source_label": "Test"},
            )
            runner = ReportRunner(session_root=sessions, pdf_builder=fake_builder)

            prepared = runner.prepare(request)

            self.assertTrue(prepared.preview_path.is_file())
            self.assertTrue(prepared.session.path.is_dir())
            result = runner.finalize(prepared, root / "final")
            self.assertTrue(result.output_path.is_file())
            self.assertFalse(prepared.session.path.exists())

    def test_retains_only_verified_final_output(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = root / "source.xlsx"
            template = root / "template.pdf"
            output = root / "final"
            sessions = root / "sessions"
            _make_workbook(workbook)
            _make_template(template)

            def fake_builder(source: Path, destination: Path, config) -> Path:
                self.assertTrue(str(source).startswith(str(sessions)))
                writer = PdfWriter()
                writer.add_blank_page(width=792, height=612)
                with destination.open("wb") as stream:
                    writer.write(stream)
                return destination

            request = ReportRunRequest(
                skill_id="template-pdf-report",
                selections={"spreadsheet": (workbook,), "template": (template,)},
                options={
                    "client_name": "Sample Household",
                    "period_label": "2026 Review",
                    "report_title": "Recommendations Review",
                    "source_label": "Synthetic test data",
                },
            )
            result = ReportRunner(session_root=sessions, pdf_builder=fake_builder).run(request, output)

            self.assertTrue(result.output_path.is_file())
            self.assertEqual(result.page_or_sheet_count, 1)
            self.assertEqual(list(sessions.iterdir()), [])

    def test_never_overwrites_existing_final(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = root / "source.xlsx"
            template = root / "template.pdf"
            output = root / "final"
            sessions = root / "sessions"
            _make_workbook(workbook)
            _make_template(template)

            def fake_builder(source: Path, destination: Path, config) -> Path:
                writer = PdfWriter()
                writer.add_blank_page(width=792, height=612)
                with destination.open("wb") as stream:
                    writer.write(stream)
                return destination

            request = ReportRunRequest(
                "template-pdf-report",
                {"spreadsheet": (workbook,), "template": (template,)},
                {"client_name": "Sample", "period_label": "2026", "report_title": "Review", "source_label": "Test"},
            )
            runner = ReportRunner(session_root=sessions, pdf_builder=fake_builder)
            first = runner.run(request, output)
            second = runner.run(request, output)

            self.assertEqual(first.output_path.name, "Sample - Review.pdf")
            self.assertEqual(second.output_path.name, "Sample - Review_v2.pdf")


if __name__ == "__main__":
    unittest.main()
