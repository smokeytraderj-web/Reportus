"""Restrained Gottfried & Somberg Wealth Management desktop theme."""

NAVY = "#10233F"
NAVY_HOVER = "#183759"
GOLD = "#B49A58"
INK = "#17202C"
MUTED = "#697586"
SURFACE = "#FFFFFF"
CANVAS = "#F5F6F8"
BORDER = "#DDE2E8"
SUCCESS = "#277A57"
DANGER = "#A83B3B"


APP_STYLESHEET = f"""
QWidget {{
    color: {INK};
    font-family: "Segoe UI";
    font-size: 14px;
}}
QMainWindow, QWidget#AppRoot {{ background: {CANVAS}; }}
QFrame#TopBar {{
    background: {SURFACE};
    border-bottom: 1px solid {BORDER};
}}
QLabel#Wordmark {{
    color: {NAVY};
    font-family: Georgia;
    font-size: 25px;
    font-weight: 600;
}}
QLabel#FirmName {{ color: {MUTED}; font-size: 11px; }}
QLabel#PageTitle {{
    color: {NAVY};
    font-family: Georgia;
    font-size: 30px;
    font-weight: 600;
}}
QLabel#SectionTitle {{ color: {NAVY}; font-size: 16px; font-weight: 600; }}
QLabel#Muted {{ color: {MUTED}; }}
QFrame#ReportCard, QFrame#Panel, QFrame#UploadBox {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 12px;
}}
QFrame#ReportCard:hover {{ border: 1px solid {GOLD}; }}
QFrame#UploadBox[active="true"] {{ border: 1px solid {GOLD}; background: #FCFAF5; }}
QPushButton {{
    border: 0;
    border-radius: 8px;
    padding: 10px 16px;
    font-weight: 600;
}}
QPushButton#PrimaryButton {{ background: {NAVY}; color: white; }}
QPushButton#PrimaryButton:hover {{ background: {NAVY_HOVER}; }}
QPushButton#PrimaryButton:disabled {{ background: #AEB6C1; }}
QPushButton#SecondaryButton {{
    background: {SURFACE};
    color: {NAVY};
    border: 1px solid {BORDER};
}}
QPushButton#CardAction {{
    background: transparent;
    color: {NAVY};
    text-align: left;
    padding: 0;
}}
QLineEdit, QTextEdit {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 8px;
    padding: 9px;
}}
QLineEdit:focus, QTextEdit:focus {{ border: 1px solid {GOLD}; }}
QTabWidget::pane {{ border: 0; background: {SURFACE}; }}
QTabBar::tab {{
    color: {MUTED};
    background: transparent;
    border: 0;
    border-bottom: 2px solid transparent;
    padding: 8px 7px;
    font-size: 11px;
    font-weight: 600;
}}
QTabBar::tab:selected {{ color: {NAVY}; border-bottom: 2px solid {GOLD}; }}
QScrollArea {{ background: {SURFACE}; border: 0; }}
QCheckBox {{ spacing: 8px; }}
QProgressBar {{
    background: #E8EBEF;
    border: 0;
    border-radius: 3px;
    max-height: 6px;
}}
QProgressBar::chunk {{ background: {GOLD}; border-radius: 3px; }}
"""
