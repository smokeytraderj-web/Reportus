"""Run a complete report job inside an isolated, disposable local session."""

from __future__ import annotations

import shutil
import datetime as dt
import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from pypdf import PdfReader

from core.audit import AuditCitation, ReportAudit
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
from generators.portfolio_pdf import (
    build_portfolio_workbook_pdf,
    is_portfolio_workbook,
)
from generators.powerpoint_content import (
    deck_content_schema,
    normalize_deck_payload,
    validate_deck_content,
)
from providers.base import SourceFragment, StructuredProvider, StructuredRequest
from providers.registry import provider_from_environment
from quality.output_qa import OutputInspector
from quality.grounding import verify_numeric_grounding
from security.privacy import PrivacyScanner
from security.portal_capture import PortalCaptureError, prepare_client_deck_portal_captures
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


@dataclass(slots=True)
class PreparedReport:
    session: ReportSession
    artifact_path: Path
    preview_path: Path
    suggested_filename: str
    page_or_sheet_count: int | None
    audit: ReportAudit | None
    revision_context: "RevisionContext | None" = None

    def clear_temporary_context(self) -> None:
        """Forget review evidence and editable content after the session ends."""

        self.audit = None
        self.revision_context = None


@dataclass(slots=True)
class RevisionContext:
    """Only the current editable state; previous report versions are never retained."""

    skill_id: str
    content: dict[str, object]
    source_paths: tuple[Path, ...]
    template_path: Path | None


PDFBuilder = Callable[[Path, Path, StockReviewConfig], Path]
PortfolioPDFBuilder = Callable[..., Path]
WorkbookBuilder = Callable[[tuple, HoldingsWorkbookConfig, Path], Path]
PowerPointBuilder = Callable[..., tuple[Path, Path]]
ClientDeckBuilder = Callable[[ClientDeckData, Path], Path]
PresentationConverter = Callable[[Path, Path], Path]


_NUMERIC_TOKEN = re.compile(r"(?<![A-Za-z])[-+]?\d[\d,.]*(?:%|x)?", re.IGNORECASE)


def _numeric_signature(value: object, path: str = "report") -> dict[str, tuple[str, ...]]:
    """Keep every displayed number anchored to the same structured field."""

    if isinstance(value, dict):
        result: dict[str, tuple[str, ...]] = {}
        for key, item in value.items():
            result.update(_numeric_signature(item, f"{path}.{key}"))
        return result
    if isinstance(value, list):
        result = {}
        for index, item in enumerate(value):
            result.update(_numeric_signature(item, f"{path}[{index}]"))
        return result
    tokens = tuple(_NUMERIC_TOKEN.findall(str(value)))
    return {path: tokens} if tokens else {}


def _single_file(selections: dict[str, tuple[Path, ...]], slot: str) -> Path:
    paths = selections.get(slot, ())
    if len(paths) != 1:
        raise ReportRunError(f"Exactly one {slot.replace('_', ' ')} file is required.")
    return paths[0]


def _validate_reference_pdf(path: Path) -> None:
    """Verify that a visual reference is readable without constraining its page shape."""

    try:
        reader = PdfReader(path)
        if reader.is_encrypted:
            raise ReportRunError("Encrypted reference PDFs are not supported.")
        if not reader.pages:
            raise ReportRunError("The reference PDF has no pages.")
        for number, page in enumerate(reader.pages, start=1):
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            if width <= 0 or height <= 0:
                raise ReportRunError(
                    f"Reference PDF page {number} has an invalid page size."
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
        portfolio_pdf_builder: PortfolioPDFBuilder | None = None,
        workbook_builder: WorkbookBuilder | None = None,
        workbook_preview_builder: WorkbookBuilder | None = None,
        powerpoint_builder: PowerPointBuilder | None = None,
        provider: StructuredProvider | None = None,
        client_deck_builder: ClientDeckBuilder | None = None,
        presentation_converter: PresentationConverter | None = None,
    ):
        self.session_root = session_root
        self.pdf_builder = pdf_builder or build_stock_review_pdf
        self.portfolio_pdf_builder = portfolio_pdf_builder or build_portfolio_workbook_pdf
        self.workbook_builder = workbook_builder or build_holdings_workbook
        self.workbook_preview_builder = workbook_preview_builder or build_holdings_snapshot
        self.powerpoint_builder = powerpoint_builder or build_powerpoint_deck
        self.provider = provider or provider_from_environment()
        self.client_deck_builder = client_deck_builder or build_client_deck
        self.presentation_converter = presentation_converter or convert_pptx_to_pdf

    def prepare(self, request: ReportRunRequest) -> PreparedReport:
        prepared_selections = None
        if request.skill_id == "client-deck-builder":
            try:
                prepared_selections = prepare_client_deck_portal_captures(request.selections)
            except PortalCaptureError as exc:
                raise ReportRunError(str(exc)) from exc
        effective_selections = (
            prepared_selections.selections if prepared_selections is not None else request.selections
        )
        source_paths = [path for paths in effective_selections.values() for path in paths]
        try:
            privacy = PrivacyScanner().scan_files(source_paths)
        except Exception:
            if prepared_selections is not None:
                prepared_selections.cleanup()
            raise
        if not privacy.approved:
            if prepared_selections is not None:
                prepared_selections.cleanup()
            categories = sorted({finding.category.value for finding in privacy.findings})
            details = []
            if categories:
                details.append("remove " + ", ".join(categories))
            if privacy.errors:
                details.append("replace files that could not be inspected")
            raise ReportRunError("Privacy check failed: " + "; ".join(details) + ".")
        if request.custom_prompt.strip():
            if prepared_selections is not None:
                prepared_selections.cleanup()
            raise ReportRunError(
                "Custom-section generation is not enabled yet; remove it before generating this report."
            )
        if request.skill_id not in {
            "client-deck-builder", "template-pdf-report", "excel-workbook-builder",
            "powerpoint-deck-builder",
        }:
            if prepared_selections is not None:
                prepared_selections.cleanup()
            raise ReportRunError("This report generator is not connected yet.")
        try:
            session = ReportSession.create(self.session_root)
        except Exception:
            if prepared_selections is not None:
                prepared_selections.cleanup()
            raise
        try:
            try:
                staged = _stage_inputs(session, effective_selections)
            finally:
                if prepared_selections is not None:
                    prepared_selections.cleanup()
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
                        "Use the Client 360 holdings workbook for overall allocation and portfolio sector "
                        "exposure. Use the J.P. Morgan attribution report for contributors and detractors. "
                        "Use the Riskalyze source for risk metrics, and the market report for sector "
                        "performance and earnings expectations. "
                        "For risk_metrics, include Portfolio total, Risk, Historical loss, Historical loss %, "
                        "Historical gain, Historical gain %, Annual dividend, Max drawdown, Annual range midpoint, "
                        "Expense ratio, and Portfolio costs when those values are shown. "
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
                section_names = (
                    "Overall Asset Allocation",
                    "Risk Snapshot",
                    "Sector YTD Performance",
                    "Equity Sector Exposure",
                    "Attribution Report",
                    "S&P 500 Earnings Expectations",
                ) + tuple(
                    {"rmd": "RMD Report", "529": "Portfolio Summary - 529 Accounts", "annuity": "Annuity Review"}[key]
                    for key in ("rmd", "529", "annuity")
                    if key in deck_data.optional_sections
                )
                audit = ReportAudit.from_staged_inputs(
                    report_type="Client Deck",
                    sections=section_names,
                    staged=staged,
                    slot_labels={
                        "risk_snapshot": "Allocation and risk data",
                        "holdings": "Current holdings and allocation",
                        "attribution": "Performance attribution",
                        "market_report": "Market and earnings data",
                        "rmd": "RMD supporting data",
                        "529": "529 supporting data",
                        "annuity": "Annuity supporting data",
                        "custom": "Custom-section support",
                    },
                    citations=tuple(
                        AuditCitation(section.replace("_", " ").title(), locator)
                        for section, locator in deck_data.sources.items()
                    ),
                )
                return PreparedReport(
                    session, preview, preview, filename, preview_qa.page_or_sheet_count, audit
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
                audit = ReportAudit.from_staged_inputs(
                    report_type="PowerPoint Deck",
                    sections=tuple(
                        str(slide.get("title", f"Slide {number}"))
                        for number, slide in enumerate(content["slides"], start=1)
                    ),
                    staged=staged,
                    slot_labels={
                        "content": "Presentation content",
                        "charts": "Chart or image",
                        "reference": "PowerPoint reference",
                        "custom": "Custom-section support",
                    },
                    citations=tuple(
                        AuditCitation(
                            str(slide.get("title", f"Slide {number}")),
                            str(slide["footer_right"]),
                        )
                        for number, slide in enumerate(content["slides"], start=1)
                        if slide.get("footer_right")
                    ),
                )
                return PreparedReport(
                    session,
                    artifact,
                    preview,
                    sanitize_filename(f"{report_title}.pptx"),
                    artifact_qa.page_or_sheet_count,
                    audit,
                    RevisionContext(
                        "powerpoint-deck-builder",
                        content,
                        tuple(sources),
                        template,
                    ),
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
                audit = ReportAudit.from_staged_inputs(
                    report_type="Excel Workbook",
                    sections=(workbook_config.sheet_name,),
                    staged=staged,
                    slot_labels={
                        "source_data": "Workbook source data",
                        "style_reference": "Style reference",
                        "custom": "Custom-section support",
                    },
                )
                return PreparedReport(
                    session, artifact, preview, filename, artifact_qa.page_or_sheet_count, audit
                )

            workbook = _single_file(staged, "spreadsheet")
            template = _single_file(staged, "template")
            _validate_reference_pdf(template)

            required = ("client_name", "period_label", "report_title", "source_label")
            missing = [key.replace("_", " ") for key in required if not request.options.get(key, "").strip()]
            if missing:
                raise ReportRunError("Missing report detail(s): " + ", ".join(missing))

            if is_portfolio_workbook(workbook):
                preview_pdf = session.preview / "report.pdf"
                self.portfolio_pdf_builder(
                    workbook,
                    template,
                    preview_pdf,
                    client_name=request.options["client_name"].strip(),
                    period_label=request.options["period_label"].strip(),
                    report_title=request.options["report_title"].strip(),
                    source_label=request.options["source_label"].strip(),
                )
                qa = OutputInspector().inspect(preview_pdf)
                if not qa.approved:
                    raise ReportRunError("The generated portfolio report failed integrity checks.")
                filename = sanitize_filename(
                    f"{request.options['client_name'].strip()} - "
                    f"{request.options['report_title'].strip()}.pdf"
                )
                audit = ReportAudit.from_staged_inputs(
                    report_type="Portfolio PDF",
                    sections=(
                        "Portfolio at a glance",
                        "Target allocation",
                        "Risk, assumptions and implementation",
                        "Recent research checks",
                        "Sources and data audit",
                    ),
                    staged=staged,
                    slot_labels={
                        "spreadsheet": "Portfolio data and source audit",
                        "template": "Approved PDF visual reference",
                        "custom": "Custom-section support",
                    },
                    citations=(
                        AuditCitation("Target allocation", f"{workbook.name} · worksheet Portfolio"),
                        AuditCitation("Recent research checks", f"{workbook.name} · worksheet Research & Assumptions"),
                        AuditCitation("Sources and data audit", f"{workbook.name} · worksheet Sources"),
                    ),
                )
                return PreparedReport(
                    session, preview_pdf, preview_pdf, filename, qa.page_or_sheet_count, audit
                )

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
            audit = ReportAudit.from_staged_inputs(
                report_type="Excel to PDF",
                sections=tuple(item.theme for item in config.reviews),
                staged=staged,
                slot_labels={
                    "spreadsheet": "Report data",
                    "template": "Approved PDF layout",
                    "custom": "Custom-section support",
                },
                citations=tuple(
                    AuditCitation(item.theme, f"{workbook.name} · worksheet {item.sheet_name}")
                    for item in config.reviews
                ),
            )
            return PreparedReport(
                session, preview_pdf, preview_pdf, filename, qa.page_or_sheet_count, audit
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
        prepared.clear_temporary_context()
        return GenerationResult(final_path, final_qa.page_or_sheet_count)

    def cancel(self, prepared: PreparedReport) -> None:
        prepared.session.cleanup()
        prepared.clear_temporary_context()

    def revise(self, prepared: PreparedReport, prompt: str) -> PreparedReport:
        """Replace one working presentation slide after privacy and grounding checks."""

        if prepared.session.closed or not prepared.session.path.is_dir():
            raise ReportRunError("The report working session is no longer available.")
        privacy = PrivacyScanner().scan_text(prompt, source="revision request")
        if not privacy.approved:
            categories = sorted({finding.category.value for finding in privacy.findings})
            if categories:
                raise ReportRunError(
                    "Revision privacy check failed: remove " + ", ".join(categories) + "."
                )
            raise ReportRunError("Enter one inspectable revision request.")
        context = prepared.revision_context
        if context is None or context.skill_id != "powerpoint-deck-builder":
            raise ReportRunError(
                "Chat revisions are currently available for PowerPoint Decks. "
                "For this report type, cancel and upload the corrected source data."
            )
        if prepared.audit is None:
            raise ReportRunError("The temporary Data & Sources audit is unavailable.")
        if not all("bullets" in slide for slide in context.content.get("slides", [])):
            raise ReportRunError(
                "Chat revisions currently support text-only PowerPoint slides. "
                "Edit the structured JSON source for table or chart slides."
            )

        fragments = ContentExtractor().extract(list(context.source_paths))
        current_fragment = SourceFragment(
            "current-report.json",
            "current approved working draft",
            json.dumps(context.content, indent=2, ensure_ascii=False),
        )
        report_title = str(context.content["title"])
        request = StructuredRequest(
            task_name="powerpoint-text-revision",
            instructions=(
                "Apply exactly one requested editorial change to exactly one slide in the current report. "
                "Do not add, remove, or reorder slides. Do not change any number, date, ticker, rating, "
                "source footer, or report title. Preserve every unchanged slide byte-for-byte in meaning "
                "and wording. The user's request is: " + json.dumps(prompt)
            ),
            json_schema=deck_content_schema(),
            sources=fragments + (current_fragment,),
            privacy_approved=True,
        )
        revised = StructuredAgent(self.provider).run(
            request,
            lambda payload: self._validated_powerpoint_revision(
                payload, context.content, fragments, report_title
            ),
        )

        candidate_artifact = prepared.session.working / "revised-report.pptx"
        candidate_preview = prepared.session.preview / "revised-report.pdf"
        try:
            self.powerpoint_builder(
                revised,
                candidate_artifact,
                candidate_preview,
                template_path=context.template_path,
            )
            artifact_qa = OutputInspector().inspect(candidate_artifact)
            preview_qa = OutputInspector().inspect(candidate_preview)
            if (
                not artifact_qa.approved
                or not preview_qa.approved
                or artifact_qa.page_or_sheet_count != preview_qa.page_or_sheet_count
            ):
                raise ReportRunError("The revised presentation failed integrity checks.")
            os.replace(candidate_artifact, prepared.artifact_path)
            os.replace(candidate_preview, prepared.preview_path)
        finally:
            candidate_artifact.unlink(missing_ok=True)
            candidate_preview.unlink(missing_ok=True)

        context.content = revised
        prepared.page_or_sheet_count = artifact_qa.page_or_sheet_count
        prepared.audit = ReportAudit(
            report_type=prepared.audit.report_type,
            sections=tuple(str(slide["title"]) for slide in revised["slides"]),
            sources=prepared.audit.sources,
            citations=tuple(
                AuditCitation(str(slide["title"]), str(slide["footer_right"]))
                for slide in revised["slides"]
                if slide.get("footer_right")
            ),
        )
        return prepared

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

    @staticmethod
    def _validated_powerpoint_revision(payload, current, fragments, report_title):
        revised = normalize_deck_payload(payload, report_title=report_title)
        old_slides = current["slides"]
        new_slides = revised["slides"]
        if len(old_slides) != len(new_slides):
            raise ReportRunError("A revision cannot add or remove slides.")
        changed = [
            index for index, (old, new) in enumerate(zip(old_slides, new_slides))
            if old != new
        ]
        if len(changed) != 1:
            raise ReportRunError("A revision must change exactly one slide.")
        for old, new in zip(old_slides, new_slides):
            if old.get("footer_right") != new.get("footer_right"):
                raise ReportRunError("A revision cannot change source citations.")
        if _numeric_signature(current) != _numeric_signature(revised):
            raise ReportRunError("A revision cannot change report numbers through chat.")
        verify_numeric_grounding(
            revised,
            fragments,
            excluded_keys=frozenset({"footer_right", "footer_left", "header_right"}),
        )
        return revised

    def run(self, request: ReportRunRequest, output_directory: Path) -> GenerationResult:
        """Non-UI convenience method that prepares and immediately finalizes."""

        prepared = self.prepare(request)
        try:
            return self.finalize(prepared, output_directory)
        except Exception:
            if not prepared.session.closed:
                prepared.session.cleanup()
            raise
