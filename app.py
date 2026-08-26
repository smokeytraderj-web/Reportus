"""Launch the Reporticles Windows desktop application."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from PySide6.QtWidgets import QApplication

from core.agent import ReporticlesAgent
from core.session import purge_abandoned_sessions
from ui.main_window import MainWindow
from ui.theme import APP_STYLESHEET


def main() -> int:
    if "--smoke-test" in sys.argv:
        reports = ReporticlesAgent().available_reports()
        if len(reports) != 4:
            raise RuntimeError("The packaged skill registry is incomplete.")
        print("Reporticles package smoke test passed.")
        return 0
    purge_abandoned_sessions()
    purge_abandoned_sessions(Path(tempfile.gettempdir()) / "Reportus" / "sessions")
    application = QApplication(sys.argv)
    application.setApplicationName("Reporticles")
    application.setOrganizationName("Gottfried & Somberg Wealth Management")
    application.setStyleSheet(APP_STYLESHEET)
    window = MainWindow()
    window.show()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
