# Reportus

A minimal Windows desktop application for building polished financial reports from clean uploaded data and explicitly requested external research.

## Current build

- Four skill-driven intake flows with local privacy and structural validation.
- Excel-to-PDF stock-review generation using the supplied GSWM template skill.
- Excel Workbook generation in the current clean/minimal GSWM holdings style, with a matching PDF review snapshot.
- PowerPoint Deck generation from a strict JSON content package through the supplied Windows PowerPoint/GSWM-template workflow.
- Provider-neutral structured synthesis with a disabled-by-default, free local Ollama option and numeric source-grounding checks.
- Client Deck PPTX/PDF generation engine with synthetic end-to-end coverage.
- Isolated report sessions, embedded PDF review, explicit finalization, automatic filename versioning, and verified cleanup.
- General document-to-deck synthesis, custom sections, external research, and revision execution remain in development.

## Development

```powershell
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Google Chrome is required for the Excel-to-PDF renderer. PowerPoint or LibreOffice is required for Client Deck PDF conversion.

See [AI provider configuration](docs/AI_PROVIDERS.md) for the optional local Ollama development setup.

Run the non-UI test suite with:

```powershell
python -m unittest discover -s tests -v
```
