---
name: client-deck-builder
description: >
  Build a GSWM client Portfolio Review PDF (Riskalyze snapshot, sector performance,
  attribution, S&P earnings outlook, plus optional RMD / 529 / annuity sections) from
  data the user provides, using python-pptx to build the branded deck then converting to
  PDF for delivery. Use this skill when someone asks for a client portfolio review, a
  quarterly/annual client deck, or references the Bloom-template GSWM Portfolio Review
  format by name. This is a different pipeline from ../template-pdf-report/SKILL.md (which
  is HTML-to-Chrome, and used for the Mudry & Rinaldi stock-picks recommendations review) —
  this one is python-pptx end to end, built from the shared helpers in style.py and the
  blank-template builder in build_template.py, and is not tied to Windows/PowerPoint COM
  automation, so it can run anywhere Python + python-pptx are available.
---

# GSWM client Portfolio Review deck

Builds professional portfolio review decks for wealth management clients, using the
Bloom Summer 2026 deck (`reference/template_pdf.pdf`) as the gold-standard visual and
structural reference. Every finished deck is delivered as a **PDF only** — the PPTX is an
intermediate build artifact, not the deliverable.

## ⚠️ Critical security rule — check before accepting any data

**Do not include sensitive data in uploads.** Before accepting data files from the user,
confirm:

- ❌ No account numbers (any specific account IDs)
- ❌ No Social Security Numbers or TINs
- ❌ No full street addresses
- ❌ No phone numbers or email addresses
- ❌ No specific account/brokerage details

What's fine to include: client names (first, or first + last), account type/term names
("Roth IRA", "IRA", "Brokerage", "Bucket 1/2/3", etc.), holdings, percentages, values,
performance data, and generic references to account purposes.

If sensitive data shows up in an uploaded file, **stop immediately** and ask the user to
clean and re-upload it before proceeding. This protects account security while keeping
client context clear.

## Workflow

### Step 1 — Verify data security
Before accepting any data files, remind the user to confirm account numbers, SSNs/TINs,
addresses, and phone/email are all removed, and client names are generic or first-name-only
if needed. If sensitive data is found in what's provided, stop and ask for a cleaned
re-upload before doing anything else.

### Step 2 — Gather core reports (required)
Ask the user to provide:

1. **Riskalyze Snapshot** — risk score and portfolio metrics, asset allocation
   percentages, max drawdown and annual range data.
2. **Attribution Report** — top performing/detracting holdings, performance contribution
   by security, date range (typically YTD or a specific period).
3. **JPM Report** (J.P. Morgan market/economic report) — sector performance data, market
   commentary, earnings expectations / S&P 500 outlook.

Accept file paths, direct data/numbers, or links to pull from.

### Step 3 — Client-specific additions (optional)
Ask whether anything else should be added: RMD (Required Minimum Distribution)
calculations, a 529 account portfolio summary, annuity/insurance product reviews, custom
market commentary, or specific asset-class deep-dives. Only include a section if real data
for it exists — never add a placeholder RMD/529/annuity section on spec.

### Step 4 — Confirm before building
Before generating anything, confirm with the user: build the PPTX with all gathered data,
apply the template layout, add GS branding, convert to PDF, and save to the output
location. Wait for explicit confirmation before proceeding.

### Step 5 — Build & deliver
1. Extract data from the provided sources.
2. Generate the PPTX using `build_template.py`'s structure and the shared helpers in
   `style.py` (navy/gold, Garamond headlines, Times New Roman body — see style.py for the
   full palette and helper functions: tables, donut/bar charts, diff-bar charts, TOC slide,
   content header/footer).
3. Embed the GS logo/crest (`add_crest` in style.py), positioned per the template.
4. Apply formatting per template standards: consistent header band, gold rule, footer with
   firm name + "Page X of Y" on every page.
5. Convert the finished PPTX to PDF and save it — PDF only, do not deliver or keep the
   PPTX as the final output.
6. Name the file `[ClientName]_[Date]_PortfolioReview.pdf` (see `archive/pdfs/` for two
   real examples of the naming and finished output: a 9-section review with RMD/529/annuity
   sections, and a leaner 6-section review without them).

## Reference assets

- `reference/template_pdf.pdf` — the gold-standard template render (all sections,
  including the optional RMD / 529 / annuity ones), used as the visual reference for
  build_template.py / style.py. Do not treat this as a real client's data — it demonstrates
  every possible section, most clients won't need all of them.
- `../../archive/pdfs/Bloom_Aug_6_pdf.pdf` and `../../archive/pdfs/Bafumi_Aug_6_pdf.pdf` —
  two real finished client decks, showing how the section list varies by client.
- `build_template.py` — generates the blank 11-slide structure-only template (title,
  contents, then one placeholder-zone slide per section). Its hardcoded `OUT` path points
  at the original dev machine location and will need updating (or an override) before
  re-running.
- `style.py` — every shared visual helper (colors, fonts, tables, charts, crest, footer,
  TOC slide). Both `build_template.py` and any per-client build script should import from
  here rather than re-implementing styling, so every deck stays visually identical.

## Data storage convention (from the original workflow)

Source materials were organized per-client under `prep_data_storage/[ClientName]/`
(riskalyze snapshot, attribution report, JPM report, screenshots), with finished PDFs going
to `final_client_deck/`. Recreate an equivalent structure under `outputs/` in this
reports-agent workspace if you want to keep that convention going forward.

## Runtime note

Unlike the PowerPoint-deck-builder skill (`../powerpoint-deck-builder/`), this pipeline
uses `python-pptx` and a PDF-conversion step rather than PowerPoint COM automation — it
does not require Windows or a local PowerPoint install, so it can run in a cloud sandbox as
well as on the user's machine. The PDF-conversion step (PPTX → PDF) still needs a
LibreOffice or PowerPoint install available wherever it runs; check for one before
assuming it's there.
