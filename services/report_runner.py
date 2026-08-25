"""Run a complete report job inside an isolated, disposable local session."""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from pypdf import PdfReader

from core.session import ReportSession
from generators.excel_to_pdf import (
    ReviewDefinition,
    StockReviewConfig,
    build_stock_review_pdf,
)
from quality.output_qa import OutputInspector
from tools.filenames import sanitize_filename, versioned_output_path


class ReportRunError(RuntimeError):
    """A report could not safely reach final output."""


@dataclass(frozen=True, slots=True)
class ReportRunRequest:
    skill_id: str
    selections: dict[str, tuple[Path, ...]]
    options: dict[str, str]
    custom_prompt: str = ""


@dataclass(frozen=True, slots=True)
class GenerationResult:
    output_path: Path
    page_or_sheet_count: int | None


@dataclass(frozen=True, slots=True)
class PreparedReport:
    session: ReportSession
    preview_path: Path
    suggested_filename: str
    page_or_sheet_count: int | None


PDFBuilder = Callable[[Path, Path, StockReviewConfig], Path]


def _single_file(selections: dict[str, tuple[Path, ...]], slot: str) -> Path:
    paths = selections.get(slot, ())
    if len(paths) != 1:
        raise ReportRunError(f"Exactly one {slot.replace('_', ' ')} file is required.")
    return paths[0]


def _validate_landscape_letter_template(path: Path) -> None:
    """Fail when a reference cannot use the installed fixed-page report layout."""

    try:
        reader = PdfReader(path)
        if not reader.pages:
            raise ReportRunError("The reference PDF has no pages.")
        for number, page in enumerate(reader.pages, start=1):
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            if abs(width - 792) > 2 or abs(height - 612) > 2:
                raise ReportRunError(
                    f"Reference PDF page {number} must be US Letter landscape (11 × 8.5 inches)."
                )
    except ReportRunError:
        raise
    except Exception as exc:
        raise ReportRunError("The reference PDF could not be verified.") from exc


def _visible_sheet_definitions(path: Path) -> tuple[ReviewDefinition, ...]:
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        visible = tuple(
            ReviewDefinition(sheet.title, sheet.title)
            for sheet in workbook.worksheets
            if sheet.sheet_state == "visible"
        )
    finally:
        workbook.close()
    if not visible:
        raise ReportRunError("The workbook has no visible report sheets.")
    return visible


def _stage_inputs(session: ReportSession, selections: dict[str, tuple[Path, ...]]) -> dict[str, tuple[Path, ...]]:
    staged: dict[str, tuple[Path, ...]] = {}
    for slot, paths in selections.items():
        slot_dir = session.uploads / sanitize_filename(slot)
        slot_dir.mkdir(parents=True, exist_ok=True)
        copies: list[Path] = []
        for index, source in enumerate(paths, start=1):
            if not source.is_file():
                raise ReportRunError("A selected source file is no longer available.")
            name = source.name if index == 1 else f"{source.stem}_{index}{source.suffix}"
            destination = slot_dir / sanitize_filename(name)
            shutil.copy2(source, destination)
            copies.append(destination)
        staged[slot] = tuple(copies)
    return staged


class ReportRunner:
    """Dispatch a report request and retain only its approved final artifact."""

    def __init__(self, *, session_root: Path | None = None, pdf_builder: PDFBuilder | None = None):
        self.session_root = session_root
        self.pdf_builder = pdf_builder or build_stock_review_pdf

    def prepare(self, request: ReportRunRequest) -> PreparedReport:
        if request.custom_prompt.strip():
            raise ReportRunError(
                "Custom-section generation is not enabled yet; remove it before generating this report."
            )
        if request.skill_id != "template-pdf-report":
            raise ReportRunError("This report generator is not connected yet.")
        session = ReportSession.create(self.session_root)
        try:
            staged = _stage_inputs(session, request.selections)
            workbook = _single_file(staged, "spreadsheet")
            template = _single_file(staged, "template")
            _validate_landscape_letter_template(template)

            required = ("client_name", "period_label", "report_title", "source_label")
            missing = [key.replace("_", " ") for key in required if not request.options.get(key, "").strip()]
            if missing:
                raise ReportRunError("Missing report detail(s): " + ", ".join(missing))
            config = StockReviewConfig(
                client_name=request.options["client_name"].strip(),
                period_label=request.options["period_label"].strip(),
                report_title=request.options["report_title"].strip(),
                source_label=request.options["source_label"].strip(),
                reviews=_visible_sheet_definitions(workbook),
            )
            preview_pdf = session.preview / "report.pdf"
            self.pdf_builder(workbook, preview_pdf, config)
            qa = OutputInspector().inspect(preview_pdf)
            if not qa.approved:
                raise ReportRunError("The generated report failed final integrity checks.")

            filename = sanitize_filename(f"{config.client_name} - {config.report_title}.pdf")
            return PreparedReport(session, preview_pdf, filename, qa.page_or_sheet_count)
        except Exception:
            session.cleanup()
            raise

    def finalize(self, prepared: PreparedReport, output_directory: Path) -> GenerationResult:
        """Copy an approved preview to final storage, verify it, then purge the session."""

        preview_qa = OutputInspector().inspect(prepared.preview_path)
        if not preview_qa.approved:
            raise ReportRunError("The report preview is no longer available or valid.")
        output_directory.mkdir(parents=True, exist_ok=True)
        final_path = versioned_output_path(output_directory, prepared.suggested_filename)
        shutil.copy2(prepared.preview_path, final_path)
        final_qa = OutputInspector().inspect(final_path)
        if not final_qa.approved:
            final_path.unlink(missing_ok=True)
            raise ReportRunError("The saved report failed final integrity checks.")
        prepared.session.cleanup()
        return GenerationResult(final_path, final_qa.page_or_sheet_count)

    def cancel(self, prepared: PreparedReport) -> None:
        prepared.session.cleanup()

    def run(self, request: ReportRunRequest, output_directory: Path) -> GenerationResult:
        """Non-UI convenience method that prepares and immediately finalizes."""

        prepared = self.prepare(request)
        try:
            return self.finalize(prepared, output_directory)
        except Exception:
            if not prepared.session.closed:
                prepared.session.cleanup()
            raise
