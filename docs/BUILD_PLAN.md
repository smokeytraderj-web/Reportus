# Reportus Build Plan

## Phase 1 - Foundation

- Establish the project instructions and architecture.
- Add typed configuration, session state, and result models.
- Create a skill registry that exposes only validated report skills.
- Import and audit the five supplied skill packages.
- Resolve inconsistent, user-specific, missing-resource, and environment-specific skill instructions.

Exit condition: all four initial menu functions resolve to a validated skill and YCharts resolves as a supporting capability.

## Phase 2 - Safety and Data Intake

- Build the local privacy scanner for supported Office files, PDFs, CSVs, and images.
- Add encrypted-file rejection.
- Implement skill-defined required and optional upload slots.
- Add schema, type, date, completeness, reconciliation, and ambiguity checks.
- Create temporary session storage and cleanup on finalize, cancel, close, and startup.

Exit condition: prohibited or malformed test fixtures cannot reach report execution, and no working files remain after session cleanup.

## Phase 3 - Desktop Shell

- Build the PySide6 Reportus window and restrained firm styling.
- Add four report cards: Client Deck, Excel to PDF, Excel Workbook, and PowerPoint Deck.
- Add drag-and-drop, file selection, optional sections, and Other / Custom Section.
- Add validation feedback, the pre-generation review, and a minimal progress screen.

Exit condition: a user can complete the entire intake and approval flow without using a terminal.

## Phase 4 - Report Execution

- Adapt Client Deck to the Reportus session and output contracts.
- Adapt Excel to PDF to the same contracts.
- Adapt Excel Workbook and PowerPoint Deck.
- Add YCharts as an internal capability with a reliable non-browser fallback.
- Add provider-independent AI and research adapters.

Exit condition: each menu function generates its declared deliverable from clean representative data.

## Phase 5 - Review, Revision, and QA

- Embed the document preview.
- Add the slim revision-chat panel without persistent history.
- Add the Data & Sources review.
- Render and visually inspect deliverables.
- Add structural checks for row counts, totals, page bounds, sources, output integrity, and stale content.
- Implement Finalize Report, automatic filename versioning, and verified cleanup.

Exit condition: finalized outputs open correctly, meet their skill-defined QA rules, and require only minor manual adjustment.

## Phase 6 - Windows Packaging

- Package Reportus for a standard Windows business laptop.
- Build an installer, desktop shortcut, uninstaller, and configurable output-folder setting.
- Test clean install, upgrade, generation, cancellation, crash cleanup, and uninstall.

Exit condition: Reportus installs and runs without a development environment.

## Deferred

- Coworker accounts and permissions
- Central company API key and protected backend
- Admin skill-import interface
- Automatic updates
- Multi-device synchronization
- Commercial or external-firm use

