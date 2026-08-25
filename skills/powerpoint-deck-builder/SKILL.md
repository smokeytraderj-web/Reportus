---
name: powerpoint-deck-builder
description: >
  Generate a Gottfried & Somberg (GSWM) branded PowerPoint deck from a JSON content file,
  using the firm's GSWM_template.pptx and PowerShell + PowerPoint COM automation. Use this
  skill whenever the user asks for a GSWM PowerPoint report as a .pptx — a weekly market
  outlook ("Animal Spirits" style), a trust/fund review, a stock-picks recap, or any deck
  that should carry the firm's title slide, navy header band, GS logo, and disclosure
  footer. Two build scripts cover the two layouts used so far: builddeck.ps1 (bullets,
  data tables, the stock-picks recap table, and a bullets-left/chart-right layout) and
  builddeckcharts.ps1 (a simpler bullets-left/chart-image-right layout for chart-driven
  decks). See ../../reports/powerpoints/spec.md for the JSON content contract each script
  expects, and archive/powerpoints/ for real examples of finished decks.
---

# GSWM PowerPoint deck builder

Every deck is built on `GSWM_template.pptx`, which carries the firm's branding:

- Slide 1 = title slide (Wall St bull image, GS circular logo, navy band)
- Slide 2 = content slide (navy gradient header band, white serif title, GS logo
  watermark bottom-left, standing disclosure footer)

The scripts fill the title slide, then duplicate the branded content slide once per entry
in a JSON content file and populate it. **Never draw the branding by hand** — it lives in
the template so every deck stays pixel-identical to the firm's real decks. If the brand
template ever changes, update `GSWM_template.pptx` here, not the scripts.

## This must run on the user's Windows machine, not in the cloud sandbox

`builddeck.ps1` and `builddeckcharts.ps1` drive PowerPoint through COM automation
(`New-Object -ComObject PowerPoint.Application`). That only works on Windows with
PowerPoint installed. From a Claude session, that means routing execution through
`device_bash` (or handing the command to the user to run locally) — it will not work in
this cloud Linux container. Stage the JSON content file and any chart images onto the
user's device first, run the script there, then commit the resulting `.pptx` back.

## Which script to use

- **`builddeck.ps1`** — the general-purpose builder. Handles plain bullet slides, a
  right-aligned header metadata line, a footer line pair, a full-bleed chart + footnote
  slide, a plain data table slide, the firm's stock-picks recap table (banner rows +
  colored rating pills), and a bullets-left/chart-drop-zone-right layout.
- **`builddeckcharts.ps1`** — a narrower, chart-first variant. Every slide is a fixed
  bullets-left / chart-image-right layout; simpler JSON, no table or picks-table support.
  Use this for decks that are purely "here are the charts and the takeaway."

Both default to `GSWM_template.pptx` sitting next to the script (`$PSScriptRoot`) unless
`-TemplatePath` is passed explicitly — keep the template file in this folder.

## Invocation

```powershell
.\builddeck.ps1 -ContentPath .\content.json -OutputPath .\out\deck.pptx
# or, with an explicit template:
.\builddeck.ps1 -ContentPath .\content.json -OutputPath .\out\deck.pptx -TemplatePath .\GSWM_template.pptx
```

`builddeckcharts.ps1` takes the same three parameters. Image paths inside the JSON
(`chart_image`, `image`) may be relative — `builddeckcharts.ps1` resolves them against the
content file's own directory; `builddeck.ps1` expects a path it can `Test-Path` directly
(absolute, or relative to the script's working directory).

## Gotchas worth knowing before you touch these scripts

- **OneDrive locking.** Both scripts build on a local temp copy (`%TEMP%`) and only copy
  the finished `.pptx` to the real output path at the end. PowerPoint COM automation
  against a file inside a syncing OneDrive folder intermittently fails mid-edit
  (`HRESULT 0x80CA1007`) — don't "simplify" this by building directly in the output folder.
- **Chrome/PowerPoint exits 0 even on a locked destination in some paths** — if the target
  `.pptx` is open in PowerPoint when the script tries to copy over it, the copy step will
  throw; close the file before regenerating.
- **Table styling.** `AddTable`'s default style bands rows and draws gridlines the script
  doesn't want; both table-building blocks in `builddeck.ps1` immediately apply the
  `{2D5ABB26-0587-4C30-8999-92F81FD0307C}` ("No Style, No Grid") style and turn off
  banding, then strip every cell border explicitly.
- **Rating badges are drawn as separate AutoShapes**, not cell fills — they're positioned
  in a second pass after row heights are finalized, sized to hug their text
  (`AutoSize = msoAutoSizeShapeToFitText`) so a short tag like "Buy" and a long one like
  "Mkt Neutral" both render correctly.
- **Font sizes scale down as content grows** (bullet count, table row count) so a busy
  slide still clears the header band and stays above the logo watermark — don't hardcode a
  single font size if you extend these scripts.

## Reference examples

`archive/powerpoints/` holds finished decks built from this template: two "Animal
Spirits" market-outlook decks, a trust review, and a fund review. Open one alongside the
scripts to see how a given JSON shape (bullets, table, picks_table, chart) renders in
practice before building a new content file from scratch.
