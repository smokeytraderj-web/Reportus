"""Reportus application pages."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

try:
    from PySide6.QtPdf import QPdfDocument
    from PySide6.QtPdfWidgets import QPdfView
except ImportError:  # pragma: no cover - depends on the packaged Qt build
    QPdfDocument = None
    QPdfView = None

from core.audit import ReportAudit
from core.workflows import ReportWorkflow
from security.privacy import PrivacyScanResult, PrivacyScanner
from validation.inputs import InputValidator
from ui.theme import DANGER, SUCCESS
from ui.widgets import ReportCard, UploadBox


class HomePage(QWidget):
    workflow_selected = Signal(str)

    def __init__(self, workflows: tuple[ReportWorkflow, ...], parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(52, 42, 52, 42)
        layout.setSpacing(10)

        eyebrow = QLabel("REPORT WORKSPACE")
        eyebrow.setStyleSheet("color: #B49A58; font-size: 11px; font-weight: 700;")
        title = QLabel("What would you like to build?")
        title.setObjectName("PageTitle")
        subtitle = QLabel("Choose a report type. Reportus will guide the required inputs and checks.")
        subtitle.setObjectName("Muted")

        layout.addWidget(eyebrow)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addSpacing(25)

        grid = QGridLayout()
        grid.setHorizontalSpacing(18)
        grid.setVerticalSpacing(18)
        for index, workflow in enumerate(workflows, start=1):
            card = ReportCard(workflow.skill_id, workflow.title, workflow.subtitle, index)
            card.selected.connect(self.workflow_selected)
            grid.addWidget(card, (index - 1) // 2, (index - 1) % 2)
        layout.addLayout(grid)
        layout.addStretch()


class IntakePage(QWidget):
    back_requested = Signal()
    review_ready = Signal(object, object, str, object)

    def __init__(
        self,
        scanner: PrivacyScanner | None = None,
        validator: InputValidator | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self.scanner = scanner or PrivacyScanner()
        self.validator = validator or InputValidator()
        self.workflow: ReportWorkflow | None = None
        self.upload_boxes: dict[str, UploadBox] = {}
        self.optional_checks: dict[str, QCheckBox] = {}
        self.field_inputs: dict[str, QLineEdit] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(42, 28, 42, 32)
        root.setSpacing(14)

        nav = QHBoxLayout()
        back = QPushButton("←  Reports")
        back.setObjectName("SecondaryButton")
        back.clicked.connect(self.back_requested)
        nav.addWidget(back)
        nav.addStretch()
        root.addLayout(nav)

        self.title = QLabel()
        self.title.setObjectName("PageTitle")
        self.subtitle = QLabel()
        self.subtitle.setObjectName("Muted")
        root.addWidget(self.title)
        root.addWidget(self.subtitle)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.content = QWidget()
        self.content_layout = QVBoxLayout(self.content)
        self.content_layout.setContentsMargins(0, 14, 8, 14)
        self.content_layout.setSpacing(12)
        scroll.setWidget(self.content)
        root.addWidget(scroll, 1)

        footer = QHBoxLayout()
        self.status = QLabel("Add the required files to continue.")
        self.status.setObjectName("Muted")
        self.review_button = QPushButton("Review inputs")
        self.review_button.setObjectName("PrimaryButton")
        self.review_button.clicked.connect(self.validate_and_review)
        footer.addWidget(self.status, 1)
        footer.addWidget(self.review_button)
        root.addLayout(footer)

    def set_workflow(self, workflow: ReportWorkflow) -> None:
        self.workflow = workflow
        self.title.setText(workflow.title)
        self.subtitle.setText(workflow.subtitle)
        self.status.setText("Add the required files to continue.")
        self.status.setStyleSheet("")
        self._clear_content()

        if workflow.fields:
            details_title = QLabel("Report details")
            details_title.setObjectName("SectionTitle")
            self.content_layout.addWidget(details_title)
            details_grid = QGridLayout()
            details_grid.setHorizontalSpacing(18)
            details_grid.setVerticalSpacing(9)
            for index, field in enumerate(workflow.fields):
                label = QLabel(field.label)
                label.setObjectName("Muted")
                editor = QLineEdit(field.default)
                editor.setPlaceholderText(field.placeholder)
                self.field_inputs[field.field_id] = editor
                row, column = divmod(index, 2)
                cell = QVBoxLayout()
                cell.setSpacing(4)
                cell.addWidget(label)
                cell.addWidget(editor)
                details_grid.addLayout(cell, row, column)
            self.content_layout.addLayout(details_grid)
            self.content_layout.addSpacing(10)

        required_title = QLabel("Required uploads")
        required_title.setObjectName("SectionTitle")
        self.content_layout.addWidget(required_title)
        for requirement in workflow.required_uploads:
            self._add_upload_box(requirement.requirement_id, requirement.label, requirement.description)

        if workflow.optional_uploads:
            self.content_layout.addSpacing(10)
            optional_title = QLabel("Optional sections")
            optional_title.setObjectName("SectionTitle")
            self.content_layout.addWidget(optional_title)
            optional_hint = QLabel("Select a section to add its supporting files.")
            optional_hint.setObjectName("Muted")
            self.content_layout.addWidget(optional_hint)
            for requirement in workflow.optional_uploads:
                check = QCheckBox(requirement.label)
                self.optional_checks[requirement.requirement_id] = check
                box = UploadBox(requirement.requirement_id, requirement.label, requirement.description)
                box.setVisible(False)
                check.toggled.connect(box.setVisible)
                self.upload_boxes[requirement.requirement_id] = box
                self.content_layout.addWidget(check)
                self.content_layout.addWidget(box)

        self.content_layout.addSpacing(10)
        custom_check = QCheckBox("Other / Custom Section")
        self.custom_prompt = QTextEdit()
        self.custom_prompt.setPlaceholderText("Describe the section or custom research you want...")
        self.custom_prompt.setFixedHeight(82)
        self.custom_prompt.setVisible(False)
        custom_box = UploadBox("custom", "Supporting files", "Add any files needed for the custom request.")
        custom_box.setVisible(False)
        custom_check.toggled.connect(self.custom_prompt.setVisible)
        custom_check.toggled.connect(custom_box.setVisible)
        self.custom_check = custom_check
        self.upload_boxes["custom"] = custom_box
        self.content_layout.addWidget(custom_check)
        self.content_layout.addWidget(self.custom_prompt)
        self.content_layout.addWidget(custom_box)
        self.content_layout.addStretch()

    def _add_upload_box(self, slot_id: str, label: str, description: str) -> None:
        box = UploadBox(slot_id, label, description)
        self.upload_boxes[slot_id] = box
        self.content_layout.addWidget(box)

    def _clear_content(self) -> None:
        self.upload_boxes.clear()
        self.optional_checks.clear()
        self.field_inputs.clear()
        while self.content_layout.count():
            item = self.content_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def validate_and_review(self) -> None:
        if self.workflow is None:
            return
        missing_fields = [
            field.label
            for field in self.workflow.fields
            if field.required and not self.field_inputs[field.field_id].text().strip()
        ]
        missing = [
            item.label
            for item in self.workflow.required_uploads
            if not self.upload_boxes[item.requirement_id].paths
        ]
        missing = missing_fields + missing
        selected_optional_missing = [
            self.optional_checks[item.requirement_id].text()
            for item in self.workflow.optional_uploads
            if self.optional_checks[item.requirement_id].isChecked()
            and not self.upload_boxes[item.requirement_id].paths
        ]
        if self.custom_check.isChecked() and (
            not self.custom_prompt.toPlainText().strip() or not self.upload_boxes["custom"].paths
        ):
            selected_optional_missing.append("Other / Custom Section description and files")
        missing.extend(selected_optional_missing)
        if missing:
            self._show_error("Missing: " + ", ".join(missing))
            return

        paths = [path for box in self.upload_boxes.values() for path in box.paths]
        result = self.scanner.scan_files(paths)
        if not result.approved:
            self._show_privacy_failure(result)
            return

        selections = {
            slot_id: tuple(box.paths)
            for slot_id, box in self.upload_boxes.items()
            if box.paths
        }
        validation = self.validator.validate(self.workflow, selections)
        if not validation.approved:
            self._show_error(" | ".join(issue.message for issue in validation.errors))
            return

        self.status.setText("Privacy and data-structure checks passed.")
        self.status.setStyleSheet(f"color: {SUCCESS};")
        options = {
            field_id: editor.text().strip()
            for field_id, editor in self.field_inputs.items()
        }
        self.review_ready.emit(
            self.workflow, selections, self.custom_prompt.toPlainText().strip(), options
        )

    def _show_error(self, message: str) -> None:
        self.status.setText(message)
        self.status.setStyleSheet(f"color: {DANGER};")

    def _show_privacy_failure(self, result: PrivacyScanResult) -> None:
        categories = sorted({finding.category.value for finding in result.findings})
        parts = []
        if categories:
            parts.append("Remove: " + ", ".join(categories))
        if result.errors:
            parts.append("Inspection blocked: " + "; ".join(result.errors))
        self._show_error(" | ".join(parts))


class ReviewPage(QWidget):
    back_requested = Signal()
    generate_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(52, 36, 52, 42)
        layout.setSpacing(14)

        back = QPushButton("←  Edit inputs")
        back.setObjectName("SecondaryButton")
        back.clicked.connect(self.back_requested)
        layout.addWidget(back)

        title = QLabel("Review before generating")
        title.setObjectName("PageTitle")
        hint = QLabel("Confirm the selected sections and files. Nothing has been sent externally.")
        hint.setObjectName("Muted")
        layout.addWidget(title)
        layout.addWidget(hint)

        self.panel = QFrame()
        self.panel.setObjectName("Panel")
        self.summary_layout = QVBoxLayout(self.panel)
        self.summary_layout.setContentsMargins(24, 22, 24, 22)
        self.summary_layout.setSpacing(10)
        layout.addWidget(self.panel)
        layout.addStretch()

        action_row = QHBoxLayout()
        self.note = QLabel("Reportus will build, verify, and save only the final report.")
        self.note.setObjectName("Muted")
        generate = QPushButton("Generate report")
        generate.setObjectName("PrimaryButton")
        generate.clicked.connect(self.generate_requested)
        action_row.addWidget(self.note)
        action_row.addStretch()
        action_row.addWidget(generate)
        layout.addLayout(action_row)

    def set_summary(
        self,
        workflow: ReportWorkflow,
        selections: dict[str, tuple[Path, ...]],
        custom_prompt: str,
        options: dict[str, str],
    ) -> None:
        while self.summary_layout.count():
            item = self.summary_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        report = QLabel(workflow.title)
        report.setObjectName("SectionTitle")
        self.summary_layout.addWidget(report)
        labels = {field.field_id: field.label for field in workflow.fields}
        for field_id, value in options.items():
            row = QLabel(f"{labels.get(field_id, field_id)}:  {value}")
            row.setObjectName("Muted")
            self.summary_layout.addWidget(row)
        for slot_id, paths in selections.items():
            label = next(
                (
                    item.label
                    for item in workflow.required_uploads + workflow.optional_uploads
                    if item.requirement_id == slot_id
                ),
                "Other / Custom Section",
            )
            row = QLabel(f"{label}:  {len(paths)} file{'s' if len(paths) != 1 else ''}")
            row.setObjectName("Muted")
            self.summary_layout.addWidget(row)
        if custom_prompt:
            custom = QLabel(f"Custom request: {custom_prompt}")
            custom.setWordWrap(True)
            custom.setObjectName("Muted")
            self.summary_layout.addWidget(custom)


class PreviewPage(QWidget):
    """Full-document preview with a deliberately slim revision rail."""

    finalize_requested = Signal()
    cancel_requested = Signal()
    revision_requested = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.document = None
        root = QVBoxLayout(self)
        root.setContentsMargins(32, 24, 32, 28)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("Review report")
        title.setObjectName("PageTitle")
        self.status = QLabel("Prepared locally · not finalized")
        self.status.setStyleSheet("color: #B49A58; font-size: 10px; font-weight: 700;")
        header.addWidget(title)
        header.addStretch()
        header.addWidget(self.status)
        root.addLayout(header)

        splitter = QSplitter()
        if QPdfView is not None and QPdfDocument is not None:
            self.viewer = QPdfView()
            self.viewer.setPageMode(QPdfView.PageMode.MultiPage)
            self.viewer.setZoomMode(QPdfView.ZoomMode.FitToWidth)
            splitter.addWidget(self.viewer)
        else:
            self.viewer = QLabel("PDF preview is unavailable in this Qt installation.")
            self.viewer.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.viewer.setObjectName("Panel")
            splitter.addWidget(self.viewer)

        rail = QFrame()
        rail.setObjectName("Panel")
        rail.setMinimumWidth(250)
        rail.setMaximumWidth(315)
        rail_layout = QVBoxLayout(rail)
        rail_layout.setContentsMargins(18, 18, 18, 18)
        rail_layout.setSpacing(10)
        rail_tabs = QTabWidget()
        revision_tab = QWidget()
        revision_layout = QVBoxLayout(revision_tab)
        revision_layout.setContentsMargins(4, 10, 4, 4)
        revision_layout.setSpacing(10)
        rail_hint = QLabel("Describe one change at a time.")
        rail_hint.setObjectName("Muted")
        rail_hint.setWordWrap(True)
        self.revision = QTextEdit()
        self.revision.setPlaceholderText("e.g., shorten the notes on page 4…")
        self.revision.setFixedHeight(110)
        self.apply_revision = QPushButton("Apply revision")
        self.apply_revision.setObjectName("SecondaryButton")
        self.apply_revision.clicked.connect(self._request_revision)
        revision_layout.addWidget(rail_hint)
        revision_layout.addWidget(self.revision)
        revision_layout.addWidget(self.apply_revision)
        revision_layout.addStretch()
        rail_tabs.addTab(revision_tab, "Revision")

        audit_tab = QWidget()
        audit_layout = QVBoxLayout(audit_tab)
        audit_layout.setContentsMargins(4, 10, 4, 4)
        audit_layout.setSpacing(8)
        audit_intro = QLabel("Temporary build evidence. Not included in the final report.")
        audit_intro.setObjectName("Muted")
        audit_intro.setWordWrap(True)
        self.audit_text = QLabel()
        self.audit_text.setObjectName("Muted")
        self.audit_text.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.audit_text.setTextFormat(Qt.TextFormat.RichText)
        self.audit_text.setWordWrap(True)
        audit_scroll = QScrollArea()
        audit_scroll.setWidgetResizable(True)
        audit_scroll.setFrameShape(QFrame.Shape.NoFrame)
        audit_scroll.setWidget(self.audit_text)
        audit_layout.addWidget(audit_intro)
        audit_layout.addWidget(audit_scroll, 1)
        rail_tabs.addTab(audit_tab, "Data & Sources")
        rail_layout.addWidget(rail_tabs, 1)
        splitter.addWidget(rail)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)
        root.addWidget(splitter, 1)

        actions = QHBoxLayout()
        self.cancel_button = QPushButton("Cancel report")
        self.cancel_button.setObjectName("SecondaryButton")
        self.cancel_button.clicked.connect(self.cancel_requested)
        self.finalize_button = QPushButton("Finalize Report")
        self.finalize_button.setObjectName("PrimaryButton")
        self.finalize_button.clicked.connect(self.finalize_requested)
        actions.addWidget(self.cancel_button)
        actions.addStretch()
        actions.addWidget(self.finalize_button)
        root.addLayout(actions)

    def set_report(
        self, path: Path, page_count: int | None, audit: ReportAudit
    ) -> None:
        count = f" · {page_count} page(s)" if page_count else ""
        self.status.setText(f"Prepared locally{count} · not finalized")
        self.revision.clear()
        self._set_audit(audit)
        if QPdfDocument is not None and QPdfView is not None:
            if self.document is not None:
                self.document.close()
                self.document.deleteLater()
            self.document = QPdfDocument(self)
            self.document.load(str(path))
            self.viewer.setDocument(self.document)

    def _set_audit(self, audit: ReportAudit) -> None:
        import html

        lines = [f"<b>{html.escape(audit.report_type)}</b>", "<br><b>Sections</b>"]
        lines.extend(f"<br>• {html.escape(section)}" for section in audit.sections)
        lines.append("<br><br><b>Approved inputs</b>")
        lines.extend(
            f"<br>• {html.escape(source.name)}<br>&nbsp;&nbsp;{html.escape(source.role)}"
            for source in audit.sources
        )
        if audit.citations:
            lines.append("<br><br><b>Section citations</b>")
            lines.extend(
                f"<br>• {html.escape(item.section)}<br>&nbsp;&nbsp;{html.escape(item.locator)}"
                for item in audit.citations
            )
        self.audit_text.setText("".join(lines))

    def _request_revision(self) -> None:
        prompt = self.revision.toPlainText().strip()
        if prompt:
            self.revision_requested.emit(prompt)

    def set_revision_busy(self, busy: bool) -> None:
        self.revision.setEnabled(not busy)
        self.apply_revision.setEnabled(not busy)
        self.cancel_button.setEnabled(not busy)
        self.finalize_button.setEnabled(not busy)
        self.apply_revision.setText("Applying…" if busy else "Apply revision")
