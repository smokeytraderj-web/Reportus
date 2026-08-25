# Reportus Project Instructions

## Product

Reportus is a single-user Windows desktop application for creating polished internal and client-facing financial reports. It transforms approved uploaded data and explicitly requested external research into report-specific deliverables.

The product must feel minimal, sleek, professional, and fast. Assume users understand financial material and do not need educational explanations.

The displayed firm name is **Gottfried & Somberg Wealth Management**. Do not append "LLC" or use an abbreviation unless a report skill explicitly requires different source-faithful wording.

## Priorities

Apply these priorities in order:

1. Protect client information.
2. Never fabricate, guess, or silently repair data.
3. Preserve numerical and source accuracy.
4. Produce clean, client-ready presentation quality.
5. Keep normal report generation within 5-7 minutes when practical.
6. Keep the interface minimal and easy to operate.

If accuracy and speed conflict, choose accuracy. If required information is uncertain, stop and ask the user.

## Global Privacy Gate

Run a local privacy inspection before any skill, AI provider, external service, parser, or report tool receives uploaded content.

Allow full client first and last names. Reject the complete upload session if any source contains:

- full or partial account numbers or account identifiers
- Social Security numbers, tax identification numbers, or equivalent government identifiers
- dates of birth
- street addresses
- phone numbers
- email addresses
- login credentials, secrets, or API keys
- bank-routing information

Do not redact and continue automatically. Explain the prohibited categories found without repeating their sensitive values, then require clean replacement files.

Reject encrypted or password-protected inputs and request unlocked clean copies.

Global privacy rules override every skill and every custom user prompt.

## Data Integrity

- Treat uploaded data as authoritative unless the selected skill or user explicitly requests external data.
- Never invent a missing value, citation, holding, calculation, date, chart, or narrative conclusion.
- Validate required files, sheets, columns, types, date ranges, and totals before generation.
- Show concise, actionable upload errors and allow replacement files.
- Ask the user to resolve likely inconsistencies.
- When uploaded and external values conflict, show both sources and pause for a decision.
- Round financial values and percentages to two decimals unless the selected skill specifies otherwise.
- Show a pre-generation review of included sections, dates, major totals, and assumptions.
- Provide a compact in-app Data & Sources audit before finalization. Do not add the full audit to the final report.
- Cite externally retrieved facts, charts, news, and sentiment in the final deliverable.

## Skill System

Skills are the primary source of report-specific behavior. `CLAUDE.md` is the source of global product, privacy, and quality rules.

- Load only the selected skill and the resources it directly requires.
- Never load all skills into one model context.
- Global rules always override skill instructions.
- A skill defines its required and optional uploads, validation contract, workflow, disclosures, filename convention, output type, and QA rules.
- If a skill leaves its output format or another material decision ambiguous, ask the user.
- Expose only validated skills in the application menu.
- Before activation, validate skill metadata, resources, allowed inputs, output rules, privacy compatibility, and a representative test run.
- Only the project owner may add or modify skills. An admin import interface is deferred.

Initial menu functions:

1. Client Deck
2. Excel to PDF
3. Excel Workbook
4. PowerPoint Deck

YCharts performance charts are a supporting capability invoked by report skills or explicit custom requests, not a primary menu function.

## User Workflow

Use this default sequence:

1. Select a report type from the main menu.
2. Show the skill-defined required upload checklist.
3. Show optional section tiles separately.
4. Provide an Other / Custom Section option requiring a description and supporting files.
5. Accept multiple files through drag-and-drop or the Windows file picker.
6. Run the privacy gate locally.
7. Validate report-specific data.
8. Gather only explicitly requested external data.
9. Present the pre-generation review and request approval.
10. Generate with a minimal progress bar and short stage labels.
11. Retry one transient generation failure automatically.
12. Show the result in an embedded full-document viewer with a slim revision-chat panel.
13. Keep AI-written commentary editable until approval.
14. Finalize only when the user presses **Finalize Report**.

Do not preserve revision-chat history or intermediate report versions.

## Files and Session Lifecycle

- Use a configurable default local output folder and provide Save As when needed.
- Let each skill determine the final deliverable type and base filename.
- Sanitize filenames for Windows.
- Never overwrite an existing final deliverable; create `_v2`, `_v3`, and later versions automatically.
- Retain uploads only inside the active temporary report session.
- On finalization, verify that the saved deliverable opens successfully, then delete uploads, extracted content, intermediate documents, temporary charts, and revision context.
- Delete the same working data when the user cancels or closes an unfinished report.
- On startup, purge temporary data left by an unexpected crash. Unfinished reports are not recoverable.
- Privacy-safe operational logs may include report type, timestamps, duration, stages, and error codes. Never log client names, source content, financial values, report text, or sensitive data.

## External Services and AI

- Internet access may be required during report generation.
- External research is opt-in and driven by the selected skill or a custom user request.
- Support market and historical data, stock-basket performance, chart images, financial news, and sentiment sources when requested.
- Keep the model provider replaceable; do not couple business logic to one AI vendor.
- Use deterministic Python for privacy inspection, parsing, calculations, charts, formatting, file handling, and QA whenever possible.
- Use AI only where language, visual interpretation, synthesis, or flexible reasoning is genuinely needed.
- Development should support free or local models with synthetic data. Production provider and key management remain configurable.
- Version one is single-user. Coworker accounts, centralized key management, a protected backend, and admin skill imports are later-phase decisions.

## Presentation Standards

- Apply a restrained Gottfried & Somberg Wealth Management visual system globally while allowing each report type to differ in layout.
- Prefer clean neutral surfaces, navy structure and typography, and restrained gold accents.
- Follow the selected skill's exact templates, fonts, disclosures, headers, footers, and pagination rules.
- Inspect rendered output visually before presenting it.
- Reject clipped text, overflow, broken tables, low-resolution charts, inconsistent spacing, incorrect pagination, stale template wording, and missing sources.
- Count and reconcile source rows against rendered rows when applicable.
- Deliverables should normally need no more than minor wording or layout adjustments.

## Desktop Application

- Target Windows business laptops without requiring a dedicated GPU.
- Build a native-feeling desktop UI with PySide6 unless a verified constraint requires another framework.
- Plan for distribution through a standard Windows installer with a desktop shortcut and uninstaller.
- Keep technical implementation details out of user-facing screens.
- Prefer short labels, spacious layouts, and clear blocking errors over dense controls or long explanations.

## Engineering Rules

- Keep orchestration thin and explicit.
- Separate privacy, validation, skill loading, external retrieval, generation, QA, and cleanup into testable components.
- Use typed data models at subsystem boundaries.
- Keep deterministic calculations out of model prompts.
- Record source provenance alongside normalized values throughout an active session.
- Add tests for privacy rejection, malformed inputs, reconciliation, filename versioning, cleanup after success/cancel/crash, skill validation, and output QA.
- Do not claim a workflow is complete until a representative end-to-end test produces and verifies the expected deliverable.

