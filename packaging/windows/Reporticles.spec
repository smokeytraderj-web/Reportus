# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller onedir build for fast startup inside the Reporticles installer."""

from pathlib import Path
from PyInstaller.utils.hooks import collect_all


PROJECT_ROOT = Path(SPECPATH).parents[1]
playwright_datas, playwright_binaries, playwright_hiddenimports = collect_all("playwright")

datas = [
    (str(PROJECT_ROOT / "config" / "skills.json"), "config"),
    (str(PROJECT_ROOT / "config" / "report_workflows.json"), "config"),
    (str(PROJECT_ROOT / "skills"), "skills"),
    (str(PROJECT_ROOT / "services" / "windows"), "services/windows"),
    (str(PROJECT_ROOT / "templates"), "templates"),
]

a = Analysis(
    [str(PROJECT_ROOT / "app.py")],
    pathex=[str(PROJECT_ROOT)],
    binaries=playwright_binaries,
    datas=datas + playwright_datas,
    hiddenimports=["PySide6.QtPdf", "PySide6.QtPdfWidgets", *playwright_hiddenimports],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Reporticles",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Reporticles",
)
