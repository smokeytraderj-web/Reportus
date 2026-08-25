"""Main Reportus window and page navigation."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from core.workflows import ReportWorkflow, load_workflows
from ui.pages import HomePage, IntakePage, ReviewPage


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Reportus")
        self.resize(1180, 780)
        self.setMinimumSize(940, 660)
        self.workflows = {item.skill_id: item for item in load_workflows()}

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
        self.stack.addWidget(self.home)
        self.stack.addWidget(self.intake)
        self.stack.addWidget(self.review)
        root_layout.addWidget(self.stack, 1)
        self.setCentralWidget(root)

        self.home.workflow_selected.connect(self.open_workflow)
        self.intake.back_requested.connect(lambda: self.stack.setCurrentWidget(self.home))
        self.intake.review_ready.connect(self.open_review)
        self.review.back_requested.connect(lambda: self.stack.setCurrentWidget(self.intake))
        self.review.generate_requested.connect(self.show_generation_placeholder)

    def open_workflow(self, skill_id: str) -> None:
        workflow = self.workflows[skill_id]
        self.intake.set_workflow(workflow)
        self.stack.setCurrentWidget(self.intake)

    def open_review(self, workflow: ReportWorkflow, selections: object, custom_prompt: str) -> None:
        self.review.set_summary(workflow, selections, custom_prompt)
        self.stack.setCurrentWidget(self.review)

    def show_generation_placeholder(self) -> None:
        QMessageBox.information(
            self,
            "Reportus foundation ready",
            "The intake and privacy flow is complete. Report execution is connected in the next phase.",
        )
