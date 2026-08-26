# Reporticles

A minimal Windows desktop application for building polished financial reports from clean uploaded data and explicitly requested external research.

## Current build

- Four skill-driven intake flows with local privacy and structural validation.
- Excel-to-PDF stock-review generation using the supplied GSWM template skill.
- Custom Excel Workbook generation from a plain-language request, with locally validated YCharts formulas, native charts, GSWM styling, and a matching PDF review preview.
- PowerPoint Deck generation from a strict JSON content package through the supplied Windows PowerPoint/GSWM-template workflow.
- Provider-neutral structured synthesis with a disabled-by-default, free local Ollama option and numeric source-grounding checks.
- Client Deck PPTX/PDF generation engine with synthetic end-to-end coverage.
- Clear closest-match Riskalyze Current Portfolio capture through a dedicated local browser profile, with ambiguous households blocked and account-level details removed before report synthesis.
- Isolated report sessions, embedded PDF review, a temporary Data & Sources audit, explicit finalization, automatic filename versioning, and verified cleanup.
- One-change, text-only PowerPoint revisions with privacy, citation, numeric-grounding, and atomic-replacement guards.
- General document-to-deck and custom Excel synthesis are connected. External research and chat revisions for deterministic Excel reports remain in development.

## Development

```powershell
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m playwright install chromium
python app.py
```

Chromium is used for the dedicated Riskalyze sign-in and capture window. Google Chrome is required for the Excel-to-PDF renderer. PowerPoint or LibreOffice is required for Client Deck PDF conversion.

See [AI provider configuration](docs/AI_PROVIDERS.md) for the optional local Ollama development setup.

Run the non-UI test suite with:

```powershell
python -m unittest discover -s tests -v
```

## Windows installer build

Builds are intentionally created on Windows so the packaged Qt and PowerPoint integrations match the deployment platform. Install Python 3 and Inno Setup 6, then run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_windows.ps1
```

The script creates an isolated build environment, runs the tests, builds the fast-starting onedir application, runs a packaged-resource smoke test, and writes the installer to `dist\installer`. Use `-SkipInstaller` when only the unpackaged application folder is needed.

The repository also includes a manually triggered **Windows Installer** GitHub Actions workflow. It builds on a clean Windows runner and provides the installer plus its SHA-256 checksum as a private workflow artifact.
