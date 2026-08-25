"""Reportus application pages."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.workflows import ReportWorkflow
from security.privacy import PrivacyScanResult, PrivacyScanner
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
    review_ready = Signal(object, object, str)

    def __init__(self, scanner: PrivacyScanner | None = None, parent=None):
        super().__init__(parent)
        self.scanner = scanner or PrivacyScanner()
        self.workflow: ReportWorkflow | None = None
        self.upload_boxes: dict[str, UploadBox] = {}
        self.optional_checks: dict[str, QCheckBox] = {}

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
        while self.content_layout.count():
            item = self.content_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def validate_and_review(self) -> None:
        if self.workflow is None:
            return
        missing = [
            item.label
            for item in self.workflow.required_uploads
            if not self.upload_boxes[item.requirement_id].paths
        ]
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

        self.status.setText("Privacy and file inspection passed.")
        self.status.setStyleSheet(f"color: {SUCCESS};")
        selections = {
            slot_id: tuple(box.paths)
            for slot_id, box in self.upload_boxes.items()
            if box.paths
        }
        self.review_ready.emit(self.workflow, selections, self.custom_prompt.toPlainText().strip())

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
        self.note = QLabel("Generation execution will be connected in the next report phase.")
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
    ) -> None:
        while self.summary_layout.count():
            item = self.summary_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        report = QLabel(workflow.title)
        report.setObjectName("SectionTitle")
        self.summary_layout.addWidget(report)
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
