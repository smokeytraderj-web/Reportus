"""Launch the Reportus Windows desktop application."""

from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from core.agent import ReportusAgent
from core.session import purge_abandoned_sessions
from ui.main_window import MainWindow
from ui.theme import APP_STYLESHEET


def main() -> int:
    if "--smoke-test" in sys.argv:
        reports = ReportusAgent().available_reports()
        if len(reports) != 4:
            raise RuntimeError("The packaged skill registry is incomplete.")
        print("Reportus package smoke test passed.")
        return 0
    purge_abandoned_sessions()
    application = QApplication(sys.argv)
    application.setApplicationName("Reportus")
    application.setOrganizationName("Gottfried & Somberg Wealth Management")
    application.setStyleSheet(APP_STYLESHEET)
    window = MainWindow()
    window.show()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
