"""Main Reportus window and page navigation."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings, QThread, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressDialog,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from core.workflows import ReportWorkflow, load_workflows
from services.report_runner import (
    GenerationResult,
    PreparedReport,
    ReportRunRequest,
    ReportRunner,
)
from ui.pages import HomePage, IntakePage, PreviewPage, ReviewPage


class GenerationWorker(QThread):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, runner: ReportRunner, request: ReportRunRequest, parent=None):
        super().__init__(parent)
        self.runner = runner
        self.request = request

    def run(self) -> None:
        try:
            self.completed.emit(self.runner.prepare(self.request))
        except Exception as exc:
            self.failed.emit(str(exc) or "Report generation failed.")


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Reportus")
        self.resize(1180, 780)
        self.setMinimumSize(940, 660)
        self.workflows = {item.skill_id: item for item in load_workflows()}
        self.current_request: ReportRunRequest | None = None
        self.prepared_report: PreparedReport | None = None
        self.runner = ReportRunner()
        self.worker: GenerationWorker | None = None
        self.progress: QProgressDialog | None = None
        self.preferences = QSettings("Reportus", "Reportus")

        root = QWidget()
        root.setObjectName("AppRoot")
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        topbar = QFrame()
        topbar.setObjectName("TopBar")
        topbar.setFixedHeight(76)
        topbar_layout = QHBoxLayout(topbar)
        topbar_layout.setContentsMargins(34, 12, 34, 12)
        brand = QVBoxLayout()
        brand.setSpacing(1)
        wordmark = QLabel("Reportus")
        wordmark.setObjectName("Wordmark")
        firm = QLabel("Gottfried & Somberg Wealth Management")
        firm.setObjectName("FirmName")
        brand.addWidget(wordmark)
        brand.addWidget(firm)
        topbar_layout.addLayout(brand)
        topbar_layout.addStretch()
        privacy = QLabel("LOCAL PRIVACY CHECK")
        privacy.setStyleSheet("color: #277A57; font-size: 10px; font-weight: 700;")
        topbar_layout.addWidget(privacy)
        root_layout.addWidget(topbar)

        self.stack = QStackedWidget()
        self.home = HomePage(tuple(self.workflows.values()))
        self.intake = IntakePage()
        self.review = ReviewPage()
        self.preview = PreviewPage()
        self.stack.addWidget(self.home)
        self.stack.addWidget(self.intake)
        self.stack.addWidget(self.review)
        self.stack.addWidget(self.preview)
        root_layout.addWidget(self.stack, 1)
        self.setCentralWidget(root)

        self.home.workflow_selected.connect(self.open_workflow)
        self.intake.back_requested.connect(lambda: self.stack.setCurrentWidget(self.home))
        self.intake.review_ready.connect(self.open_review)
        self.review.back_requested.connect(lambda: self.stack.setCurrentWidget(self.intake))
        self.review.generate_requested.connect(self.start_generation)
        self.preview.finalize_requested.connect(self.finalize_report)
        self.preview.cancel_requested.connect(self.cancel_prepared_report)
        self.preview.revision_requested.connect(self.request_revision)

    def open_workflow(self, skill_id: str) -> None:
        workflow = self.workflows[skill_id]
        self.intake.set_workflow(workflow)
        self.stack.setCurrentWidget(self.intake)

    def open_review(
        self,
        workflow: ReportWorkflow,
        selections: dict[str, tuple[Path, ...]],
        custom_prompt: str,
        options: dict[str, str],
    ) -> None:
        self.current_request = ReportRunRequest(
            skill_id=workflow.skill_id,
            selections=selections,
            options=options,
            custom_prompt=custom_prompt,
        )
        self.review.set_summary(workflow, selections, custom_prompt, options)
        self.stack.setCurrentWidget(self.review)

    def _output_directory(self) -> Path | None:
        saved = self.preferences.value("output_directory", "", type=str)
        if saved and Path(saved).is_dir():
            return Path(saved)
        selected = QFileDialog.getExistingDirectory(
            self,
            "Choose the Reportus output folder",
            str(Path.home() / "Documents"),
        )
        if not selected:
            return None
        self.preferences.setValue("output_directory", selected)
        return Path(selected)

    def start_generation(self) -> None:
        if self.current_request is None or self.worker is not None:
            return
        self.progress = QProgressDialog("Building and checking your report…", "", 0, 0, self)
        self.progress.setWindowTitle("Reportus")
        self.progress.setCancelButton(None)
        self.progress.setMinimumDuration(0)
        self.progress.setAutoClose(False)
        self.progress.show()

        self.worker = GenerationWorker(self.runner, self.current_request, self)
        self.worker.completed.connect(self._generation_completed)
        self.worker.failed.connect(self._generation_failed)
        self.worker.finished.connect(self._worker_finished)
        self.worker.start()

    def _generation_completed(self, result: PreparedReport) -> None:
        if self.progress is not None:
            self.progress.close()
        self.prepared_report = result
        self.preview.set_report(result.preview_path, result.page_or_sheet_count)
        self.stack.setCurrentWidget(self.preview)

    def _generation_failed(self, message: str) -> None:
        if self.progress is not None:
            self.progress.close()
        QMessageBox.warning(self, "Report not generated", message)

    def _worker_finished(self) -> None:
        if self.worker is not None:
            self.worker.deleteLater()
        self.worker = None
        self.progress = None

    def finalize_report(self) -> None:
        if self.prepared_report is None:
            return
        output_directory = self._output_directory()
        if output_directory is None:
            return
        try:
            result: GenerationResult = self.runner.finalize(self.prepared_report, output_directory)
        except Exception as exc:
            QMessageBox.warning(self, "Report not finalized", str(exc) or "Finalization failed.")
            return
        self.prepared_report = None
        self.current_request = None
        QMessageBox.information(self, "Report finalized", f"Saved to:\n{result.output_path}")
        self.stack.setCurrentWidget(self.home)

    def cancel_prepared_report(self) -> None:
        if self.prepared_report is not None:
            self.runner.cancel(self.prepared_report)
        self.prepared_report = None
        self.current_request = None
        self.stack.setCurrentWidget(self.home)

    def request_revision(self, prompt: str) -> None:
        QMessageBox.information(
            self,
            "Revision queued for the next build slice",
            "The preview and revision workspace are ready. Revision execution is not connected yet.",
        )

    def closeEvent(self, event) -> None:
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.information(
                self,
                "Report generation in progress",
                "Wait for generation to finish, then finalize or cancel the report.",
            )
            event.ignore()
            return
        if self.prepared_report is not None and not self.prepared_report.session.closed:
            self.runner.cancel(self.prepared_report)
        super().closeEvent(event)
