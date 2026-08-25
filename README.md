# Reportus

A minimal Windows desktop application for building polished financial reports from clean uploaded data and explicitly requested external research.

## Current build

- Four skill-driven intake flows with local privacy and structural validation.
- Excel-to-PDF stock-review generation using the supplied GSWM template skill.
- Client Deck PPTX/PDF generation engine with synthetic end-to-end coverage.
- Isolated report sessions, embedded PDF review, explicit finalization, automatic filename versioning, and verified cleanup.
- Excel Workbook, PowerPoint Deck, custom sections, external research, and revision execution remain in development.

## Development

```powershell
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Google Chrome is required for the Excel-to-PDF renderer. PowerPoint or LibreOffice is required for Client Deck PDF conversion.

Run the non-UI test suite with:

```powershell
python -m unittest discover -s tests -v
```
