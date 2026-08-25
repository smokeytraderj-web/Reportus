---
name: ycharts-performance-charts
description: Pull YTD performance charts for a stock or fund, three chart types, standalone YTD performance, YTD performance vs the S&P 500, and YTD performance vs the security's GICS sector benchmark. Trigger this whenever John asks to "chart performance," "pull a chart," "chart YTD," "chart it against the S&P," "chart it against sector/sector index," or references getting a performance chart for a ticker, even if he doesn't say "YCharts" by name. Primary path uses ycharts.com via the Claude in Chrome extension with an active YCharts login. If Chrome isn't connected, don't stop, use the no-browser fallback in this skill instead.
---

# YCharts YTD Performance Charts

Builds the three chart types John regularly needs for client decks and internal review, pulled live from ycharts.com using browser automation rather than the YCharts Excel add-in (see the separate YCharts spreadsheet workflow for `YCP()` formula-based work, that's a different tool for a different job).

1. **Standalone YTD performance** — one security, year to date, total return.
2. **YTD performance vs S&P 500** — the security plotted against `^SPX` (or `SPY` if the index ticker throws an error), year to date.
3. **YTD performance vs sector** — the security plotted against its GICS sector's SPDR Select Sector ETF, year to date.

## Prerequisites

- Claude in Chrome must be connected. If it isn't, tell John and stop, don't fall back to guessing chart values.
- John's YCharts session needs to already be logged in in that Chrome profile. **Never enter a YCharts username or password.** If a login page appears, stop, tell John the session isn't logged in, and ask him to log in in that browser tab himself. Resume once he confirms.

## Step 1: Resolve the tickers

- **Primary security**: whatever ticker John gives. For stocks and most ETFs, use it as-is. For mutual funds, use the bare 5-letter ticker on the website (the `M:` prefix is specific to the Excel add-in's `YCP()` formula, the YCharts website doesn't use it).
- **S&P 500 benchmark**: use the index ticker `^SPX`. If YCharts can't resolve it on the chart tool, fall back to the ETF `SPY`.
- **Sector benchmark**: ask John which GICS sector the security belongs to if it isn't obvious, then map it to the corresponding SPDR Select Sector ETF:

| GICS Sector | Sector ETF |
|---|---|
| Communication Services | XLC |
| Consumer Discretionary | XLY |
| Consumer Staples | XLP |
| Energy | XLE |
| Financials | XLF |
| Health Care | XLV |
| Industrials | XLI |
| Information Technology | XLK |
| Materials | XLB |
| Real Estate | XLRE |
| Utilities | XLU |

If John gives multiple securities at once (e.g. a short watchlist), repeat the whole workflow per security rather than cramming them onto one chart, unless he explicitly asks to compare them against each other instead of against benchmarks.

## Step 2: Open the chart tool

1. `tabs_context_mcp` (createIfEmpty: true) to get a tab, then `tabs_create_mcp` for a fresh tab for this task.
2. Navigate to the security's fundamental chart page:
   - Stocks/ETFs/mutual funds: `https://ycharts.com/companies/{TICKER}/chart`
   - Indices: `https://ycharts.com/indices/%5E{TICKER}/chart` (note the URL-encoded `^`)
3. If the page shows a paywall or login prompt instead of the chart tool, stop and follow the login rule above.

## Step 3: Configure the chart

Use `find` and `read_page` to locate controls rather than guessing coordinates blind, then confirm with a screenshot before reading values off the chart.

1. **Metric**: set the plotted metric to **Total Return** (matches the convention already used in John's YCharts spreadsheet work, `total_return_forward_adjusted_price`). It's usually a dropdown or metric-search box near the top of the chart tool labeled something like "Metric" or "Add a metric."
2. **Date range**: set the range to **Year to Date**. YCharts exposes this as a preset button (YTD) alongside 1M/3M/6M/1Y/etc, use the preset rather than typing custom dates, it's less error-prone.
3. **Comparison securities** (chart types 2 and 3 only): use the chart tool's "Add security" / "Compare" field to add `^SPX` (or `SPY`) for the S&P chart, or the sector ETF for the sector chart. Add only one comparison security per chart, don't combine S&P and sector on the same chart unless John asks for that as a fourth variant.
4. Screenshot the result and visually confirm: right date range, right securities in the legend, percentage total return on the y-axis (not raw price).

## Step 4: Capture the chart

Prefer YCharts' own export over a browser screenshot, it renders cleaner (no browser chrome, correct DPI) and is what ends up in client decks:

1. Look for a **Download** / **Export** / camera-icon control on the chart (usually top-right of the chart panel). If present, use it to save a PNG.
2. If no native export is available or it's gated behind a plan John doesn't have, fall back to `computer` screenshot with `zoom` cropped tightly to the chart panel, `save_to_disk: true`.

Save each image to the workspace, then copy the three files to `/mnt/user-data/outputs` and use `present_files`. Name them clearly:

```
{TICKER}_YTD.png
{TICKER}_vs_SP500_YTD.png
{TICKER}_vs_{SECTOR_ETF}_YTD.png
```

## Fallback: Claude in Chrome isn't connected

Don't just stop if the browser extension isn't available. Switch to a lower-fidelity path that needs no browser at all, using `web_search` (and `web_fetch` only on URLs that search actually surfaced, since direct API/CSV endpoints for this kind of data are consistently robots-blocked, don't burn calls trying `stooq.com` or `query1/2.finance.yahoo.com` directly, they'll fail).

This path trades a polished exported line chart for a quick, honest comparison chart. Tell John up front, in one line, that Chrome isn't connected so you're pulling snapshot figures instead of an exported line chart, then proceed rather than stopping.

1. For each leg (the security itself, `^SPX`/S&P 500, and the sector ETF), run a targeted search: `"{TICKER} year to date return"`, `"S&P 500 YTD return 2026"`, `"{SECTOR_ETF} YTD return 2026"`. Aggregator sites (stockanalysis.com, finance news writeups) frequently state the YTD % directly in the snippet.
2. `stockanalysis.com/stocks/{ticker}/` and `/history/` are reliably fetchable (no robots block) if you land there through search results first, but their historical table only returns the most recent ~50 trading days via a plain fetch (pagination past that is JS-driven and won't show up), so treat it as a fresh-price sanity check, not a source for a full Jan-to-date series.
3. Build the three charts with the native chart tool (`chart_display_v0`, bar style) instead of image exports:
   - Chart 1: one bar, the security's YTD total return.
   - Chart 2: two bars, security vs S&P 500 YTD total return.
   - Chart 3: two bars, security vs sector ETF YTD total return.
4. Title each chart with the as-of date from the source (e.g. "AAPL YTD Total Return, as of Aug 25 2026") so it's clear it's a snapshot pulled from public data, not a YCharts export.
5. In your reply, note plainly that this was the fallback path and a full exported line chart is available once Chrome reconnects, don't present it as equivalent without saying so.

## Notes

- These are raw chart pulls, not a finished slide. If John wants them dropped into a GSWM deck, that's the `work deck` / `financial-media-decks` or `powerpoint-deck-builder` workflow, hand the images off there rather than trying to build the deck inside this skill.
- No em dashes in any labels or filenames you generate, commas instead, per John's standing formatting rule.
- Close any tabs you opened for this task when you're done, per the browser tool's cleanup convention.
