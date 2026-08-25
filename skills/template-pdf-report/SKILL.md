---
name: template-pdf-report
description: >
  Rebuild a polished, client-ready PDF report from spreadsheet data so that it matches an
  existing PDF exactly — same fonts, colours, table grid, page furniture and pagination —
  then re-run it whenever the data refreshes. Use this skill whenever someone hands over a
  sample/template/"gold standard" PDF plus a data file and asks for "the same thing with new
  numbers", "make it look exactly like this", "turn this spreadsheet into a report", or wants
  a recurring client deck regenerated for a new period. Also use it when editing such a report
  later (changing columns, wording, sections, page order), because the report is generated
  from a build script rather than edited by hand — reach for this even if the request sounds
  like a small one-off tweak to a PDF. This is the HTML-to-headless-Chrome pipeline; for the
  separate python-pptx-based portfolio review deck, see ../client-deck-builder/SKILL.md.
---

# Matching an existing PDF report, driven by fresh data

The goal is a report indistinguishable from the template except for the numbers, and a build
script the user can re-run every period. Hand-editing a PDF is a dead end — the next data
refresh throws the edits away — so everything here produces a *generator*.

## Why HTML → headless Chrome

Most well-designed business PDFs were printed from HTML by Chrome, and rebuilding them the
same way is far more tractable than driving ReportLab or python-pptx. Two signs confirm it,
both worth checking early (step 1): fonts are ordinary system faces (Georgia, Calibri,
Arial), and point sizes convert to clean pixel values at 96 dpi — 9.75pt = 13px,
8.25pt = 11px, 18.75pt = 25px. If sizes land on clean px, the original was CSS, and you can
recover the CSS almost exactly.

Useful constants: **1pt = 1.3333px**. US-Letter landscape = 792×612pt = **1056×816px**.
Portrait = 612×792pt = 816×1056px.

## Step 1 — Measure the template, don't eyeball it

Run the bundled inspector to dump every text span (position, size, font, colour), every
vector drawing (rects, rules, fills), and page renders:

```bash
python scripts/inspect_pdf.py template.pdf --out ./_inspect --pages 1-4
```

Then **look at the rendered PNGs**. Numbers give you geometry; only your eyes catch that a
band is a gradient or that a chip has rounded corners.

Two things to extract carefully, because they define the whole grid:

- **Column boundaries** come from the table's cell-background or rule rectangles, not from
  the text. Read the `x0` values of the row rules — those are the true column edges. Convert
  to percentages of the table width so the CSS is resolution-independent.
- **Gradients and images don't appear as drawings.** A navy band that has no fill rectangle
  is a gradient (or an image). Sample the rendered pixels at the corners to recover it —
  `inspect_pdf.py --sample x,y` does this.

### The one insight that makes calibration exact

PyMuPDF's span `bbox` is **not** the ink extent. Its top is `baseline − ascent` and its
bottom is `baseline + descent` for the font at that size. That is exactly a CSS line box at
`line-height: normal`. So a span bbox tells you where the line box sits, and you can solve
for padding and line-height directly instead of guessing:

```
bbox_height ≈ (ascent + descent) × font_size     # Georgia ≈ 1.136em, Calibri ≈ 1.22em
line_box_top = bbox_top − (line_height − bbox_height) / 2
baseline     = bbox_top + ascent × font_size
```

Worked example: a Georgia-Bold 25px heading whose bbox top is 41.1px sits in a line box
whose top is 42.8px when `line-height: 1`, giving a baseline at 64.0px. If the design centres
that block in an 83px band, you can now tell whether it is centred, bottom-aligned, or
padded — and reproduce it in one try rather than ten.

## Step 2 — Build the page skeleton

Each page is a fixed-size box; nothing relies on flow between pages. This keeps pagination
under your control, which matters because you will want page breaks in specific places.

```css
@page { size: 11in 8.5in; margin: 0; }
.page { position: relative; width: 1056px; height: 816px; overflow: hidden;
        page-break-after: always; background: #fff; }
.page:last-child { page-break-after: auto; }
```

Set `-webkit-print-color-adjust: exact` on `body`, or every background disappears in print.
Use `table-layout: fixed` with percentage `<col>` widths so columns land exactly where the
template's rules are.

## Step 3 — Calibrate, then verify by diffing

Build, render, and compare span positions against the template rather than trusting your
arithmetic:

```bash
python scripts/compare_layout.py template.pdf output.pdf --page 4
```

It prints matched spans side by side with deltas. Chase anything over ~1px on structural
elements (band heights, table origins, column edges); sub-pixel differences are invisible.

Two calibration details that repeatedly cost time:

- **Letter-spacing.** Width scales linearly, so one measurement solves it: with `n` gaps
  between characters, `letter_spacing = (target_width − natural_width) / n`, where
  `natural_width` comes from rendering the same string at a known spacing. Different elements
  often use different values — don't assume one global figure.
- **Inline-block chips inside table cells** sit on the cell's text baseline, which pushes
  them down by the cell's leading. Set `line-height: 0` on that cell so the chip's own box
  is all that matters, then position it with padding.

## Step 4 — Data conventions worth getting right

These come up in nearly every financial or operational report:

- **Split composite fields carefully.** A note like `"Buy - solid AI monetization"` carries a
  label and prose. Split on the first separator, but only when the leading fragment is short
  and is genuinely a label — validate it against a whitelist. Sector names and commentary
  ("Large-cap energy", "Agree") sneak into the label position and must not be presented as
  ratings. When the lead isn't a real label, keep it at the front of the prose so nothing is
  lost.
- **Never invent a value silently.** If a field is genuinely absent, either render it blank
  or resolve it explicitly in a small, commented lookup table keyed to the row — and quote
  the source text in the comment so the user can audit each call. Then make the build *fail*
  when a new row appears with no entry, so a blank cell can't slip into a client document.
- **Compute page numbers in two passes** — first lay out sections into pages, then number
  them — so "Page 4 of 17" and any contents-page ranges stay correct when content shifts.
- **Pin page breaks explicitly** (e.g. `SPLITS = {1: [13], 5: [13]}` meaning "break section 1
  after 13 rows") rather than relying on CSS pagination. It is predictable, and it lets you
  reproduce the template's exact splits.
- **Derived stats belong in one function** so a summary page and a total row can't disagree.

## Step 5 — Render and save

Use the bundled renderer. It exists because **Chrome exits 0 even when it fails to write the
PDF** — if the target is open in a viewer, you get a success message and a stale file, which
is genuinely hard to notice and wastes a lot of time:

```bash
python scripts/render_pdf.py build.html out/report.pdf
```

It renders to a unique temp path, checks the file exists and is a plausible size, then copies
it into place and raises a clear error if the destination is locked.

If a destination stays locked, don't silently write next to it under a different name and
move on — say so. A stale file the user keeps reviewing is worse than a failed build.

## Step 6 — QA before you hand it over

```bash
python scripts/qa_pdf.py out/report.pdf --limit 561
```

This flags content running past the safe bottom margin, reports page count, and writes page
renders. Then **open the renders and look at them** — at minimum the cover, one page of every
distinct layout, and the last page. Check specifically for:

- text overflowing a cell or running under the footer
- headers whose alignment doesn't match their column's values (right-aligned numbers under a
  left-aligned heading is the classic tell)
- a wrapped cell leaving what looks like an empty row
- stale wording from the template that no longer applies to the new data

Count things programmatically too — rows rendered vs rows in the source, chips vs rows — so a
dropped record can't hide.

## Working on Windows

- Write Python to a file and run it; PowerShell 5.1 has no heredocs and `&&`/`||` are parse
  errors. Chain with `;` or `if ($?) { ... }`.
- Don't pipe a native exe's stderr (`2>&1`, `2>$null`) — PowerShell turns it into an error
  record. Prefer invoking Chrome from Python.
- Files under OneDrive lock hard while open in a viewer. `os.replace` and even `Rename-Item`
  will fail; the fix is to close the app, not to work around it.

## Reference implementation

`references/build_stock_picks.py` is a complete, working build of a 17-page landscape client
report from an Excel workbook — data loading, label parsing, pagination, page numbering, all
five page layouts, an embedded logo, and the safe-render step. It builds the Mudry & Rinaldi
stock-picks recommendations review. Read it when you want to see how the pieces fit together;
adapt its structure rather than starting from scratch. Note it has firm- and client-specific
constants near the top (`XLSX`, `CHROME`, `OUT_DIR`, `INFERRED_RATINGS`) that need updating
for a new client or a moved data file.

`references/calibration.md` (not yet added here) would hold the extraction checklist and the
font-metric table (ascent/descent for the common system faces) used in step 1.

## Keeping it re-runnable

Leave the user with one command that reads the current data file and rewrites the PDF, plus
a short note on where to change wording, column widths and page breaks. The generated HTML is
worth keeping on disk too — it opens in a browser, so the user can experiment with DevTools
and hand back concrete values.
