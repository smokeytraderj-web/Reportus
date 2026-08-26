"""Gottfried & Somberg-inspired visual system for the Reporticles desktop app."""

NAVY = "#1B2A4A"
NAVY_DARK = "#0A1224"
NAVY_MID = "#2B3D63"
NAVY_HOVER = "#324970"
GOLD = "#BFA054"
GOLD_LIGHT = "#D6C189"
GOLD_WASH = "#F6F0E2"
INK = "#1B2A4A"
MUTED = "#657188"
WARM_MUTED = "#C9B88A"
SURFACE = "#FFFFFF"
CANVAS = "#F6F4EF"
BORDER = "#DDD8CC"
SUCCESS = "#277A57"
DANGER = "#A83B3B"


APP_STYLESHEET = f"""
QWidget {{
    color: {INK};
    font-family: "Segoe UI";
    font-size: 14px;
}}
QMainWindow, QWidget#AppRoot, QStackedWidget {{ background: {CANVAS}; }}
QFrame#TopBar {{
    background: {NAVY};
    border: 0;
    border-bottom: 2px solid {GOLD};
}}
QFrame#BrandMark {{
    background: transparent;
    border: 1px solid {GOLD};
    border-radius: 24px;
}}
QLabel#BrandInitials {{
    color: {GOLD};
    background: transparent;
    font-family: Georgia;
    font-size: 19px;
    font-weight: 700;
}}
QLabel#Wordmark {{
    color: {SURFACE};
    background: transparent;
    font-family: Georgia;
    font-size: 22px;
    font-weight: 600;
}}
QLabel#FirmName {{
    color: {WARM_MUTED};
    background: transparent;
    font-size: 10px;
    font-weight: 600;
}}
QLabel#PrivacyBadge {{
    color: {GOLD_LIGHT};
    background: {NAVY};
    border: 1px solid {NAVY_MID};
    border-radius: 14px;
    padding: 7px 12px;
    font-size: 9px;
    font-weight: 700;
}}
QFrame#HomeHero {{
    background: {NAVY};
    border: 0;
    border-radius: 3px;
}}
QLabel#Eyebrow {{
    color: {GOLD_LIGHT};
    background: transparent;
    font-size: 10px;
    font-weight: 700;
}}
QLabel#SectionEyebrow, QLabel#StatusLabel {{
    color: {GOLD};
    font-size: 10px;
    font-weight: 700;
}}
QLabel#HeroTitle {{
    color: {SURFACE};
    background: transparent;
    font-family: Georgia;
    font-size: 34px;
    font-weight: 600;
}}
QLabel#HeroSubtitle {{
    color: #E3E7EE;
    background: transparent;
    font-size: 14px;
}}
QFrame#GoldRule {{
    background: {GOLD};
    border: 0;
    min-height: 2px;
    max-height: 2px;
}}
QLabel#PageTitle {{
    color: {NAVY};
    font-family: Georgia;
    font-size: 31px;
    font-weight: 600;
}}
QLabel#SectionTitle {{
    color: {NAVY};
    font-family: Georgia;
    font-size: 17px;
    font-weight: 600;
}}
QLabel#FieldLabel {{
    color: {NAVY_MID};
    font-size: 11px;
    font-weight: 700;
}}
QLabel#Muted {{ color: {MUTED}; }}
QLabel#FileLabel {{ color: {NAVY}; font-size: 12px; font-weight: 600; }}
QLabel#CardNumber {{
    color: {GOLD};
    font-family: Georgia;
    font-size: 17px;
    font-weight: 700;
}}
QFrame#ReportCard {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-top: 3px solid {GOLD};
    border-radius: 3px;
}}
QFrame#ReportCard:hover {{
    background: #FFFEFA;
    border: 1px solid {GOLD};
    border-top: 3px solid {GOLD};
}}
QFrame#Panel {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 3px;
}}
QFrame#UploadBox {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-left: 3px solid {NAVY_MID};
    border-radius: 3px;
}}
QFrame#UploadBox[active="true"] {{
    border: 1px solid {GOLD};
    border-left: 3px solid {GOLD};
    background: {GOLD_WASH};
}}
QPushButton {{
    border: 0;
    border-radius: 3px;
    padding: 10px 17px;
    font-weight: 700;
}}
QPushButton#PrimaryButton {{ background: {GOLD}; color: {NAVY_DARK}; }}
QPushButton#PrimaryButton:hover {{ background: {GOLD_LIGHT}; }}
QPushButton#PrimaryButton:pressed {{ background: #A98C45; }}
QPushButton#PrimaryButton:disabled {{ background: #CAC4B6; color: #747064; }}
QPushButton#SecondaryButton {{
    background: {SURFACE};
    color: {NAVY};
    border: 1px solid {BORDER};
}}
QPushButton#SecondaryButton:hover {{
    color: {NAVY_DARK};
    border: 1px solid {GOLD};
    background: {GOLD_WASH};
}}
QPushButton#CardAction {{
    background: transparent;
    color: {NAVY};
    text-align: left;
    padding: 2px 0;
    font-size: 12px;
    font-weight: 700;
}}
QPushButton#CardAction:hover {{ color: {GOLD}; }}
QLineEdit, QTextEdit {{
    background: {SURFACE};
    color: {INK};
    border: 1px solid {BORDER};
    border-radius: 3px;
    padding: 10px;
    selection-background-color: {NAVY};
    selection-color: {SURFACE};
}}
QLineEdit:hover, QTextEdit:hover {{ border: 1px solid #C5BFAF; }}
QLineEdit:focus, QTextEdit:focus {{
    border: 1px solid {GOLD};
    background: #FFFEFA;
}}
QTabWidget::pane {{
    border: 1px solid {BORDER};
    border-radius: 2px;
    background: {SURFACE};
}}
QTabBar::tab {{
    color: {MUTED};
    background: transparent;
    border: 0;
    border-bottom: 2px solid transparent;
    padding: 9px 8px;
    font-size: 10px;
    font-weight: 700;
}}
QTabBar::tab:selected {{ color: {NAVY}; border-bottom: 2px solid {GOLD}; }}
QScrollArea {{ background: transparent; border: 0; }}
QScrollArea > QWidget > QWidget {{ background: transparent; }}
QScrollBar:vertical {{
    background: transparent;
    width: 9px;
    margin: 2px;
}}
QScrollBar::handle:vertical {{
    background: #C9C2B2;
    border-radius: 4px;
    min-height: 28px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QCheckBox {{ spacing: 9px; color: {NAVY}; font-weight: 600; }}
QCheckBox::indicator {{ width: 16px; height: 16px; }}
QProgressBar {{
    background: #E4E0D7;
    border: 0;
    border-radius: 2px;
    max-height: 5px;
}}
QProgressBar::chunk {{ background: {GOLD}; border-radius: 2px; }}
QSplitter::handle {{ background: {BORDER}; width: 1px; }}
QToolTip {{
    background: {NAVY_DARK};
    color: {SURFACE};
    border: 1px solid {GOLD};
    padding: 6px;
}}
"""
