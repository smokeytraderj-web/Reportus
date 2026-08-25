"""Run a complete report job inside an isolated, disposable local session."""

from __future__ import annotations

import shutil
import datetime as dt
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from pypdf import PdfReader

from core.session import ReportSession
from core.structured_agent import StructuredAgent
from extraction.content import ContentExtractor
from generators.excel_workbook import (
    HoldingsWorkbookConfig,
    build_holdings_snapshot,
    build_holdings_workbook,
    load_holdings,
)
from generators.client_deck import ClientDeckData, build_client_deck
from generators.client_deck_content import (
    client_deck_schema,
    normalize_client_deck_payload,
)
from generators.excel_to_pdf import (
    ReviewDefinition,
    StockReviewConfig,
    build_stock_review_pdf,
)
from generators.powerpoint_content import (
    deck_content_schema,
    normalize_deck_payload,
    validate_deck_content,
)
from providers.base import StructuredProvider, StructuredRequest
from providers.registry import provider_from_environment
from quality.output_qa import OutputInspector
from quality.grounding import verify_numeric_grounding
from security.privacy import PrivacyScanner
from services.powerpoint_com import build_powerpoint_deck
from services.conversion import convert_pptx_to_pdf
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
    artifact_path: Path
    preview_path: Path
    suggested_filename: str
    page_or_sheet_count: int | None


PDFBuilder = Callable[[Path, Path, StockReviewConfig], Path]
WorkbookBuilder = Callable[[tuple, HoldingsWorkbookConfig, Path], Path]
PowerPointBuilder = Callable[..., tuple[Path, Path]]
ClientDeckBuilder = Callable[[ClientDeckData, Path], Path]
PresentationConverter = Callable[[Path, Path], Path]


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

    def __init__(
        self,
        *,
        session_root: Path | None = None,
        pdf_builder: PDFBuilder | None = None,
        workbook_builder: WorkbookBuilder | None = None,
        workbook_preview_builder: WorkbookBuilder | None = None,
        powerpoint_builder: PowerPointBuilder | None = None,
        provider: StructuredProvider | None = None,
        client_deck_builder: ClientDeckBuilder | None = None,
        presentation_converter: PresentationConverter | None = None,
    ):
        self.session_root = session_root
        self.pdf_builder = pdf_builder or build_stock_review_pdf
        self.workbook_builder = workbook_builder or build_holdings_workbook
        self.workbook_preview_builder = workbook_preview_builder or build_holdings_snapshot
        self.powerpoint_builder = powerpoint_builder or build_powerpoint_deck
        self.provider = provider or provider_from_environment()
        self.client_deck_builder = client_deck_builder or build_client_deck
        self.presentation_converter = presentation_converter or convert_pptx_to_pdf

    def prepare(self, request: ReportRunRequest) -> PreparedReport:
        source_paths = [path for paths in request.selections.values() for path in paths]
        privacy = PrivacyScanner().scan_files(source_paths)
        if not privacy.approved:
            categories = sorted({finding.category.value for finding in privacy.findings})
            details = []
            if categories:
                details.append("remove " + ", ".join(categories))
            if privacy.errors:
                details.append("replace files that could not be inspected")
            raise ReportRunError("Privacy check failed: " + "; ".join(details) + ".")
        if request.custom_prompt.strip():
            raise ReportRunError(
                "Custom-section generation is not enabled yet; remove it before generating this report."
            )
        if request.skill_id not in {
            "client-deck-builder", "template-pdf-report", "excel-workbook-builder",
            "powerpoint-deck-builder",
        }:
            raise ReportRunError("This report generator is not connected yet.")
        session = ReportSession.create(self.session_root)
        try:
            staged = _stage_inputs(session, request.selections)
            if request.skill_id == "client-deck-builder":
                required_options = ("client_name", "period", "as_of", "report_date")
                missing = [key.replace("_", " ") for key in required_options if not request.options.get(key, "").strip()]
                if missing:
                    raise ReportRunError("Missing report detail(s): " + ", ".join(missing))
                try:
                    report_date = dt.date.fromisoformat(request.options["report_date"].strip())
                except ValueError as exc:
                    raise ReportRunError("Report date must use YYYY-MM-DD format.") from exc
                source_paths = tuple(
                    path for slot_paths in staged.values() for path in slot_paths
                )
                fragments = ContentExtractor().extract(list(source_paths))
                included_slots = ", ".join(sorted(staged))
                structured_request = StructuredRequest(
                    task_name="client-deck",
                    instructions=(
                        "Map the approved portfolio-review sources into the Client Deck JSON contract. "
                        "Copy values exactly; do not calculate missing fields or infer investment conclusions. "
                        "Returns and contributions must be decimal numbers (0.052 means 5.20%); sector exposures "
                        "must be percentage points (5.20 means 5.20%). Include optional_sections only when the "
                        f"corresponding real source is present. Uploaded sections: {included_slots}. "
                        "Every sources value must name the source file and page, row, or slide locator."
                    ),
                    json_schema=client_deck_schema(),
                    sources=fragments,
                    privacy_approved=True,
                )
                deck_data = StructuredAgent(self.provider).run(
                    structured_request,
                    lambda payload: self._validated_client_deck(
                        payload, fragments, request.options
                    ),
                )
                intermediate = session.working / "client-deck.pptx"
                preview = session.preview / "report.pdf"
                self.client_deck_builder(deck_data, intermediate)
                self.presentation_converter(intermediate, preview)
                deck_qa = OutputInspector().inspect(intermediate)
                preview_qa = OutputInspector().inspect(preview)
                if (
                    not deck_qa.approved
                    or not preview_qa.approved
                    or deck_qa.page_or_sheet_count != preview_qa.page_or_sheet_count
                ):
                    raise ReportRunError("The Client Deck or its PDF preview failed integrity checks.")
                client_slug = request.options["client_name"].strip().replace(" ", "_")
                filename = sanitize_filename(
                    f"{client_slug}_{report_date.isoformat()}_PortfolioReview.pdf"
                )
                return PreparedReport(
                    session, preview, preview, filename, preview_qa.page_or_sheet_count
                )

            if request.skill_id == "powerpoint-deck-builder":
                sources = staged.get("content", ())
                if not sources:
                    raise ReportRunError("At least one deck content file is required.")
                report_title = request.options.get("report_title", "").strip()
                if not report_title:
                    raise ReportRunError("Presentation title is required.")
                references = staged.get("reference", ())
                if len(references) > 1:
                    raise ReportRunError("Use at most one PowerPoint reference template.")
                template = references[0] if references else None
                if template is not None and template.suffix.lower() != ".pptx":
                    raise ReportRunError("The PowerPoint reference must be a .pptx file.")
                images = staged.get("charts", ())
                json_sources = [path for path in sources if path.suffix.lower() == ".json"]
                if json_sources:
                    if len(sources) != 1:
                        raise ReportRunError("Use one JSON package by itself, or use source documents without JSON.")
                    content = validate_deck_content(
                        json_sources[0], report_title=report_title, image_paths=images
                    )
                else:
                    fragments = ContentExtractor().extract(list(sources))
                    structured_request = StructuredRequest(
                        task_name="powerpoint-deck",
                        instructions=(
                            "Create a concise, client-ready financial presentation from the supplied sources. "
                            "Each slide must advance a clear narrative and use only facts present in the sources. "
                            "Use source filenames and page/row/slide locators in each footer_right line. "
                            "Do not add external facts, recommendations, dates, values, or conclusions."
                        ),
                        json_schema=deck_content_schema(),
                        sources=fragments,
                        privacy_approved=True,
                    )
                    content = StructuredAgent(self.provider).run(
                        structured_request,
                        lambda payload: self._validated_powerpoint_content(
                            payload, fragments, report_title, images
                        ),
                    )
                artifact = session.working / "report.pptx"
                preview = session.preview / "report.pdf"
                self.powerpoint_builder(content, artifact, preview, template_path=template)
                artifact_qa = OutputInspector().inspect(artifact)
                preview_qa = OutputInspector().inspect(preview)
                if (
                    not artifact_qa.approved
                    or not preview_qa.approved
                    or artifact_qa.page_or_sheet_count != preview_qa.page_or_sheet_count
                ):
                    raise ReportRunError("The presentation or its preview failed integrity checks.")
                return PreparedReport(
                    session,
                    artifact,
                    preview,
                    sanitize_filename(f"{report_title}.pptx"),
                    artifact_qa.page_or_sheet_count,
                )

            if request.skill_id == "excel-workbook-builder":
                if staged.get("style_reference"):
                    raise ReportRunError("Custom workbook style references are not connected yet.")
                source = _single_file(staged, "source_data")
                report_title = request.options.get("report_title", "").strip()
                source_label = request.options.get("source_label", "").strip()
                if not report_title or not source_label:
                    raise ReportRunError("Workbook title and source label are required.")
                holdings = load_holdings(source)
                workbook_config = HoldingsWorkbookConfig(report_title, source_label)
                artifact = session.working / "report.xlsx"
                preview = session.preview / "report.pdf"
                self.workbook_builder(holdings, workbook_config, artifact)
                self.workbook_preview_builder(holdings, workbook_config, preview)
                artifact_qa = OutputInspector().inspect(artifact)
                preview_qa = OutputInspector().inspect(preview)
                if not artifact_qa.approved or not preview_qa.approved:
                    raise ReportRunError("The generated workbook or preview failed integrity checks.")
                filename = sanitize_filename(f"{report_title}.xlsx")
                return PreparedReport(
                    session, artifact, preview, filename, artifact_qa.page_or_sheet_count
                )

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
            return PreparedReport(
                session, preview_pdf, preview_pdf, filename, qa.page_or_sheet_count
            )
        except Exception:
            session.cleanup()
            raise

    def finalize(self, prepared: PreparedReport, output_directory: Path) -> GenerationResult:
        """Copy an approved preview to final storage, verify it, then purge the session."""

        artifact_qa = OutputInspector().inspect(prepared.artifact_path)
        if not artifact_qa.approved:
            raise ReportRunError("The prepared report is no longer available or valid.")
        output_directory.mkdir(parents=True, exist_ok=True)
        final_path = versioned_output_path(output_directory, prepared.suggested_filename)
        shutil.copy2(prepared.artifact_path, final_path)
        final_qa = OutputInspector().inspect(final_path)
        if not final_qa.approved:
            final_path.unlink(missing_ok=True)
            raise ReportRunError("The saved report failed final integrity checks.")
        prepared.session.cleanup()
        return GenerationResult(final_path, final_qa.page_or_sheet_count)

    def cancel(self, prepared: PreparedReport) -> None:
        prepared.session.cleanup()

    @staticmethod
    def _validated_client_deck(payload: dict[str, object], fragments, options) -> ClientDeckData:
        verify_numeric_grounding(
            payload,
            fragments,
            excluded_keys=frozenset({"sources", "earnings_years"}),
        )
        return normalize_client_deck_payload(
            payload,
            client_name=options["client_name"],
            period=options["period"],
            as_of=options["as_of"],
        )

    @staticmethod
    def _validated_powerpoint_content(payload, fragments, report_title, images):
        verify_numeric_grounding(
            payload,
            fragments,
            excluded_keys=frozenset({"footer_right", "footer_left", "header_right"}),
        )
        return normalize_deck_payload(
            payload, report_title=report_title, image_paths=images
        )

    def run(self, request: ReportRunRequest, output_directory: Path) -> GenerationResult:
        """Non-UI convenience method that prepares and immediately finalizes."""

        prepared = self.prepare(request)
        try:
            return self.finalize(prepared, output_directory)
        except Exception:
            if not prepared.session.closed:
                prepared.session.cleanup()
            raise
