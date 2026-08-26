"""Reusable Reporticles interface controls."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
)


FILE_FILTER = (
    "Supported files (*.xlsx *.xlsm *.xls *.csv *.tsv *.pdf *.pptx *.ppt "
    "*.docx *.doc *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp);;All files (*.*)"
)


class UploadBox(QFrame):
    """Compact multi-file drop zone with an accessible file picker."""

    files_changed = Signal(str, list)

    def __init__(self, slot_id: str, label: str, description: str, parent=None):
        super().__init__(parent)
        self.slot_id = slot_id
        self.paths: list[Path] = []
        self.setObjectName("UploadBox")
        self.setProperty("active", False)
        self.setAcceptDrops(True)
        self.setMinimumHeight(108)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 15, 18, 15)
        layout.setSpacing(5)

        self.title_row = QHBoxLayout()
        title = QLabel(label)
        title.setObjectName("SectionTitle")
        self.pick_button = QPushButton("Choose files")
        self.pick_button.setObjectName("SecondaryButton")
        self.pick_button.clicked.connect(self.choose_files)
        self.title_row.addWidget(title)
        self.title_row.addStretch()
        self.title_row.addWidget(self.pick_button)

        hint = QLabel(description)
        hint.setObjectName("Muted")
        hint.setWordWrap(True)
        self.file_label = QLabel("Drop files here or choose files")
        self.file_label.setObjectName("Muted")

        layout.addLayout(self.title_row)
        layout.addWidget(hint)
        layout.addWidget(self.file_label)

    def add_action(self, label: str, callback) -> QPushButton:
        """Add a workflow-specific action beside the standard file picker."""

        button = QPushButton(label)
        button.setObjectName("SecondaryButton")
        button.clicked.connect(callback)
        self.title_row.insertWidget(self.title_row.count() - 1, button)
        return button

    def choose_files(self) -> None:
        selected, _ = QFileDialog.getOpenFileNames(self, "Select report files", "", FILE_FILTER)
        if selected:
            self.set_files([Path(item) for item in selected])

    def set_files(self, paths: list[Path]) -> None:
        unique = list(dict.fromkeys(path.resolve() for path in paths if path.is_file()))
        self.paths = unique
        self.setProperty("active", bool(unique))
        self.style().unpolish(self)
        self.style().polish(self)
        if not unique:
            self.file_label.setText("Drop files here or choose files")
        elif len(unique) == 1:
            self.file_label.setText(unique[0].name)
        else:
            self.file_label.setText(f"{len(unique)} files selected")
        self.files_changed.emit(self.slot_id, unique)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.isLocalFile()]
        self.set_files(paths)
        event.acceptProposedAction()


class ReportCard(QFrame):
    """Home-screen card for one validated report workflow."""

    selected = Signal(str)

    def __init__(self, skill_id: str, title: str, subtitle: str, index: int, parent=None):
        super().__init__(parent)
        self.skill_id = skill_id
        self.setObjectName("ReportCard")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(180)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(9)

        number = QLabel(f"0{index}")
        number.setStyleSheet("color: #B49A58; font-size: 12px; font-weight: 700;")
        title_label = QLabel(title)
        title_label.setObjectName("SectionTitle")
        subtitle_label = QLabel(subtitle)
        subtitle_label.setObjectName("Muted")
        subtitle_label.setWordWrap(True)
        action = QPushButton("Start report  →")
        action.setObjectName("CardAction")
        action.clicked.connect(lambda: self.selected.emit(self.skill_id))

        layout.addWidget(number)
        layout.addWidget(title_label)
        layout.addWidget(subtitle_label)
        layout.addStretch()
        layout.addWidget(action)

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.selected.emit(self.skill_id)
        super().mouseReleaseEvent(event)
