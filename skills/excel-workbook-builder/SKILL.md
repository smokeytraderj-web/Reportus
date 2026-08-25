---
name: excel-workbook-builder
description: >
  Build a client-ready Excel workbook in the GSWM clean/minimal "Stock List" style (serif
  firm name, thin navy rule, tracked-caps column labels, ticker+company inline in one
  column, subtle zebra rows, no heavy gridlines) from a plain data source, plus optionally
  a matching standalone PDF snapshot of the same table. Use this skill whenever someone
  asks for a stock/holdings tracker, a current-vs-target price sheet, a holdings summary,
  or any small recurring data table as a formatted .xlsx (not a PowerPoint or a full
  multi-page PDF report — those are ../powerpoint-deck-builder/ and
  ../template-pdf-report/ or ../client-deck-builder/).
---

# GSWM Excel workbook — clean/minimal "Stock List" style

Three scripts, all openpyxl/reportlab-driven — nothing here needs Windows or COM
automation, so all can run in a cloud sandbox as well as locally.

**Current style vs. legacy style:** `build_holdings_summary.py` is the current reference
implementation, matching `../../reports/excel-workbooks/examples/Steppins_Stock_List.pdf`
— white background, serif firm name, one thin navy rule, tracked-caps gray labels, ticker
(bold) + company name (gray) inline in a single wide "Company" column, subtle zebra rows,
no gridlines. `build_steppins.py` / `build_steppins_table.py` (below) predate that example
and use an older, boxier navy-banner-with-gridlines look — kept for reference/history, but
check with the user before using them for new client-facing work.

## `build_holdings_summary.py` — current style, both `.xlsx` and matching PDF

One script, two output functions (`build_xlsx()`, `build_pdf()`), so the workbook and its
PDF snapshot can never drift out of visual sync with each other. Reads a custodial
holdings export (`.xls`/`.xlsx`: Description, Symbol, Quantity, Price, Value, % of
Assets), skips the source file's account-header row entirely (see the compliance note
below), and writes:

- A two-part header: large serif `FIRM_LABEL` (a generic label, not the account's full
  legal name — see Compliance) top-left, small tracked-caps `REPORT_LABEL` top-right, one
  thin navy rule underneath, then a small gray source-note line.
- Small tracked-caps gray column headers (`spaced()` helper — see the whitespace gotcha
  below).
- One wide "Company" column holding ticker (bold, navy) + company name (gray) **together**
  rather than a separate fixed-width ticker column. This is deliberate: a narrow dedicated
  ticker column wraps awkwardly on longer tickers (`GOOGL` breaking mid-word) — putting
  both in one wide cell lets any ticker length sit naturally on one line.
- Subtle zebra striping (very light gray, not a strong tint), no cell gridlines/borders
  except the header underline and a rule above the Total row.
- Long company/fund names that would still overflow the column get `wrap_text=True` and a
  computed taller row height (see the `wraps` calculation) rather than being silently
  clipped by the next cell.
- Currency (`$#,##0.00`) and percent (`0.00%`) number formats; a bold Total row with a
  navy rule above it.
- `freeze_panes` on the header row, print setup (portrait, fit-to-width, explicit print
  area, ~0.5" margins, header row repeats via `print_title_rows`) so it also
  prints/exports cleanly.

**Whitespace gotcha in the PDF path:** ReportLab `Paragraph` markup is XML and collapses
runs of plain ASCII spaces, so building tracked-caps text with `" ".join(...)` renders as
one run-together word (e.g. "HOLDINGSSUMMARY"). `spaced(text, pdf=True)` uses a real
U+00A0 non-breaking space instead, which isn't collapsed. Excel cells don't collapse
whitespace, so `spaced(text)` (default, `pdf=False`) uses a plain space there.

**Ampersand gotcha in the PDF path:** company/fund names containing `&` (AT&T, Johnson &
Johnson, iShares Core S&P 500 ETF, JPMorgan Chase & Co) will break ReportLab's Paragraph
XML parsing and render garbled (e.g. "AT&T; Inc") unless escaped. Every text field going
into a `Paragraph()` goes through `xml.sax.saxutils.escape` first — don't remove that when
adapting this script.

To adapt for a new client/tracker: change the constants at the top of the script
(`SRC`, `SHEET`, `FIRM_LABEL`, `REPORT_LABEL`, `SOURCE_NOTE`, `XLSX_OUT`, `PDF_OUT`) and,
if the source columns differ, the `load_holdings()` parsing loop and the column list in
`build_xlsx()`/`build_pdf()`.

## `build_steppins.py` — legacy box-style workbook

Reads a source workbook (`steppins_formatted.xlsx`, expected in the same folder as the
script — not included here, since it's per-client input data) and writes a formatted
`.xlsx` with:

- A merged navy title banner row (Garamond, white, bold) plus a small gray subtitle row
  (source attribution).
- A banded header row (`Company`, `Ticker`, `Current Price`, `Target Price`,
  `Upside/Downside`), all cells boxed with a thin light-navy border.
- Zebra-striped data rows (alternating white / light navy-tint fill).
- Currency formatting (`$#,##0.00`) on price columns, percent formatting
  (`+0.0%;-0.0%`) on the upside/downside column, colored green for upside and red for
  downside.
- An em dash in gray for any row missing a current price or target price, rather than a
  blank cell.
- Rows with no target price (typically index/mutual funds without a discrete target) are
  computed but then hidden, not deleted — so the underlying data survives even though it
  doesn't render.
- `freeze_panes` on the header row, and print setup (portrait, fit-to-width/height,
  explicit print area, 0.4"/0.5" margins) so the sheet also prints or exports to PDF
  cleanly as a single page.

To adapt this for a new client or tracker: change `SRC`/`OUT`, the `cols`/`widths` list,
and the column-reading logic in the loop over `rows` to match the new source sheet's
layout. The box-drawing helper (`put(...)`) is generic and worth keeping as-is.

## `build_steppins_table.py` — a matching PDF snapshot (legacy style)

A separate, ReportLab-based script that renders the *same visual style* (navy header
banner, gridded table, zebra rows, green/red upside coloring) as a standalone one-page
PDF, independent of the `.xlsx` — useful when a client wants a quick PDF glance at the
tracker rather than opening Excel. **As delivered, its `HOLDINGS` list is hardcoded** at
the top of the script rather than reading from the workbook — before relying on this for
a real refresh, wire it up to read the same source data `build_steppins.py` reads (or the
finished `.xlsx`), so the two outputs can't drift apart.

This is a different rendering approach from both PDF skills already in this workspace
(`../template-pdf-report/` is HTML-to-Chrome, `../client-deck-builder/` is python-pptx) —
it uses `reportlab.platypus` tables directly. Keep it scoped to small single-table
snapshots like this one rather than extending it into a multi-page report generator; for
that, reach for one of the other two PDF skills instead.

## Reference examples

- `../../reports/excel-workbooks/examples/Steppins_Stock_List.pdf` — the **current** style
  reference (clean/minimal, see `build_holdings_summary.py` above).
- `../../archive/excel-workbooks/Steppins_Stock_Tracker.xlsx` — a real finished output of
  the **legacy** `build_steppins.py` (box style) — 15 holdings, one hidden (no target
  price). Kept for history, not a style reference for new work.
- `../../archive/excel-workbooks/` — finished holdings-summary outputs in the current
  style land here once delivered to a client (see `workflows/delivery.md`).

## Compliance note

Never carry a source file's account number or a client's full legal account/trust title
into the output. `build_holdings_summary.py`'s `load_holdings()` already skips the
source's account-header row entirely; `FIRM_LABEL` must be set to a short generic label
(confirm the exact wording with the user) rather than copied from that row. See
`../../goals/compliance.md`.
