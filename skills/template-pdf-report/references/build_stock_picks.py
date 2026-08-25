# -*- coding: utf-8 -*-
"""
Build the Mudry & Rinaldi recommendations review PDF from the live YCharts workbook,
reproducing the layout of layout/stock_picks_template.pdf exactly.

Month names are replaced by appointment numbers; the recommendation/measurement
dates carry the calendar information in the subtitles.

Pipeline:  xlsx -> HTML (Georgia/Calibri, 11x8.5in landscape) -> headless Chrome -> PDF
"""

import os
import re
import html
import base64
import shutil
import subprocess
import tempfile
import datetime
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX = r"C:\Users\Jack.Muirhead\OneDrive - Commonwealth Advisors\JJM Misc\Sheets\mudry_rinaldi_reports.xlsx"
OUT_DIR = os.path.join(BASE, "final_client_deck")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

CLIENT = "Eric Mudry &amp; Steve Rinaldi"
FIRM = "Gottfried &amp; Somberg Wealth Management, LLC"

# Sheet order -> appointment number. Themes mirror the template's contents page.
APPOINTMENTS = [
    ("January",  "Stock List tracker, plus top picks from Dave, Matt &amp; Nick"),
    ("February", "Asset-light / AI-disintermediation names, plus earnings-season ideas"),
    ("March",    "Core mega-cap ideas, Energy, US Shale, and Software"),
    ("April",    "Positions to build, positions to trim, and asset-allocation ideas"),
    ("May",      "Reiterations &amp; updates, Aerospace, and new/other names"),
    ("June",     "2H opportunities, Capital Markets, Aerospace &amp; Defense, European military spending"),
]

# Data rows emitted before forcing a page break (matches the template's pagination).
SPLITS = {1: [13], 5: [13], 6: [10]}

# Source Firm values that are colleagues rather than outside research shops.
INTERNAL = {"dave", "matt", "nick"}

# Words that make up an actual rating. Anything else in the leading position of a
# note is a sector or a comment ("Large-cap energy", "Agree", "Software switch"),
# so it moves into the Notes column and the Rating shows an em dash.
RATING_WORDS = {"buy", "sell", "hold", "pass", "add", "trim", "build", "strong",
                "reiterate", "overweight", "underweight", "neutral"}

# Rows whose workbook note carries no explicit rating. Each call below is read off
# that row's own note text (quoted), so every ticker shows a rating. Keyed by
# (appointment number, ticker) because a name can recur with a different view.
INFERRED_RATINGS = {
    # "Agree - appears oversold; GLP-1 volume growth offsets pricing pressure."
    (1, "NVO"): "Buy",
    # "Like the name, but large existing position ... hesitant to add."
    (1, "AMZN"): "Hold",
    # "Agree - Buy since late 2025. Top-5 global pharma, fwd P/E 18.5x."
    (1, "AZN"): "Buy",
    # "Agree - housing market set to improve in 2026 and beyond."
    (1, "HD"): "Buy",
    # "price target $140 versus $174 price implies -18% downside."
    (3, "XOM"): "Pass",
    # "price target $181 versus $213 price implies -14% downside."
    (3, "CVX"): "Pass",
    # "realize losses in individual software positions and reinvest into IGV."
    (4, "IGV"): "Buy",
    # "rate stabilization has reduced financing risk and M&A premiums are rising."
    (4, "IJR"): "Buy",
    # "thesis intact, record backlog; 'priced to perfection' ... cut upside."
    (5, "RNMBY"): "Hold",
    # "Not a reiterate - B-737 Max is competitive but flawed ... tariff noise."
    (5, "BA"): "Pass",
    # "Strong 1Q, beat on revenue and earnings ... PT $183 (+19.6%)."
    (5, "ICE"): "Buy",
    # "Take capital loss ... re-establish after 31 days. Still best-in-class."
    (6, "BX"): "Hold",
    # "Multiple compression, not fundamental weakness ... PT EUR 1,350 (+19%)."
    (6, "RNMBY"): "Buy",
    # "Trades 23x, fair versus US defense ... Revenue growth 6-7%."
    (6, "THLLY"): "Hold",
    # "High-single-digit growth expected. Fwd P/E 18-19x."
    (6, "BAESY"): "Hold",
}


# ----------------------------------------------------------------- data layer

def fmt_date(d):
    if isinstance(d, datetime.datetime):
        d = d.date()
    return "{} {}, {}".format(d.strftime("%B"), d.day, d.year)


def is_rating(text):
    """True only for an actual call (Buy, Hold/Pass, Reiterate Buy, ...).
    Sectors and commentary are not ratings."""
    tokens = [t for t in re.split(r"[/\s]+", text.lower()) if t]
    return bool(tokens) and all(t in RATING_WORDS for t in tokens)


def split_rating(note):
    """Workbook convention: 'Rating - note text'. Parentheticals are dropped from
    the rating. If the leading phrase is not an actual rating (a sector such as
    'Large-cap energy', or a comment such as 'Agree'), the Rating column shows an
    em dash and the phrase is kept at the front of the note."""
    note = note.split(">>")[0].strip()
    idx = note.find(" - ")
    if idx != -1 and idx <= 40:
        lead = re.sub(r"\s*\([^)]*\)", "", note[:idx]).strip()
        rest = note[idx + 3:].strip()
        if is_rating(lead):
            return lead, rest
        return "\u2014", "{} \u2014 {}".format(lead, rest)
    return "\u2014", note


def load_appointments():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    out = []
    for num, (sheet, theme) in enumerate(APPOINTMENTS, start=1):
        ws = wb[sheet]
        rows = [r for r in ws.iter_rows(values_only=True)]

        as_of = rec_date = bench = None
        sections = []          # [(name, [row, ...]), ...]

        for r in rows:
            vals = list(r) + [None] * (10 - len(r))
            b, c, d, e, f, g, h, i, j = vals[1:10]
            if b is None:
                continue
            label = str(b).strip()

            if label.startswith("As-Of Date"):
                as_of = c
                continue
            if label.startswith("Default Rec. Date"):
                rec_date = c
                continue
            if label.startswith("S&P 500 Total Return (benchmark)"):
                bench = f
                continue
            if label == "Company Name" or label.endswith("(live YCharts)"):
                continue

            if c is None and d is None:          # section divider row
                sections.append((label, []))
                continue

            ticker = str(c).strip()
            rating, notes = split_rating(str(h or ""))
            if rating == "—":
                rating = INFERRED_RATINGS.get((num, ticker), "—")
            firm = (str(j).strip() if j else "")
            if firm and firm.lower() not in INTERNAL:
                notes = "{} (Source: {})".format(notes.rstrip(), firm)

            sections[-1][1].append({
                "name": label,
                "ticker": ticker,
                "rating": rating,
                "price": float(i),
                "ret": float(f),
                "vs": float(g),
                "notes": notes,
            })

        out.append({
            "num": num,
            "theme": theme,
            "as_of": as_of,
            "rec_date": rec_date,
            "bench": float(bench),
            "sections": sections,
        })
    return out


# -------------------------------------------------------------- presentation

def pct(v):
    return "{}{:.1f}%".format("+" if v >= 0 else "", v * 100)


def pct_class(v):
    return "pos" if v >= 0 else "neg"


def money(v):
    return "${:,.2f}".format(v)


def paginate(appt):
    """Split an appointment's sections across pages, repeating section headings
    with (CONT'D) where a section straddles a break."""
    breaks = list(SPLITS.get(appt["num"], []))
    pages, cur, count = [], [], 0
    for name, rows in appt["sections"]:
        head, bucket = name, []
        for row in rows:
            bucket.append(row)
            count += 1
            if breaks and count == breaks[0]:
                breaks.pop(0)
                cur.append((head, bucket))
                pages.append(cur)
                cur, bucket = [], []
                head = name + " (CONT'D)"
        if bucket or not rows:
            cur.append((head, bucket))
    if cur:
        pages.append(cur)
    return pages


LOGO_PATH = os.path.join(BASE, "layout", "gs_logo.png")


def logo_tag():
    with open(LOGO_PATH, "rb") as fh:
        b64 = base64.b64encode(fh.read()).decode("ascii")
    return '<img class="seal" src="data:image/png;base64,{}" alt="">'.format(b64)

CSS = """
@page { size: 11in 8.5in; margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }

.page {
  position: relative; width: 1056px; height: 816px; overflow: hidden;
  background: #fff; page-break-after: always; break-after: page;
}
.page:last-child { page-break-after: auto; break-after: auto; }

/* ---- dark pages (cover + dividers) ---- */
.dark { background: linear-gradient(135deg, #223a63 0%, #0a1425 100%); }
.dark::before { content: ""; position: absolute; left: 0; top: 0; width: 13px; height: 100%; background: #b08c57; }

.eyebrow { font-family: Calibri, sans-serif; font-weight: bold; color: #c9a66b;
           letter-spacing: .275em; }

/* ---- cover ---- */
.seal-wrap { position: absolute; left: 0; right: 0; top: 62px; text-align: center; }
.seal { width: 160px; height: 160px; display: inline-block; }

.cover-eyebrow { position: absolute; left: 100.8px; top: 312.2px; font-size: 12px; line-height: 17px; }
.cover-h1 { position: absolute; left: 100.8px; top: 340px; margin: 0;
            font-family: Georgia, serif; font-weight: bold; font-size: 46px; line-height: 46px;
            color: #fff; }
.cover-sub { position: absolute; left: 100.8px; top: 395.6px;
             font-family: Georgia, serif; font-size: 20px; line-height: 20px; color: #fff; }
.cover-rule { position: absolute; left: 100.8px; top: 446.9px; width: 86px; height: 3px; background: #b08c57; }
.cover-desc { position: absolute; left: 100.8px; top: 463.6px; width: 716px;
              font-family: Calibri, sans-serif; font-style: italic; font-size: 13px; line-height: 20px;
              color: #b9c2d4; }
.cover-foot { position: absolute; left: 100.8px; top: 750.3px;
              font-family: Calibri, sans-serif; font-size: 11px; line-height: 13.4px; color: #8592ac; }

/* ---- divider ---- */
.div-eyebrow { position: absolute; left: 100.8px; top: 208.3px; font-size: 12px; line-height: 14.6px; }
.div-h1 { position: absolute; left: 100.8px; top: 233.8px; margin: 0;
          font-family: Georgia, serif; font-weight: bold; font-size: 52px; line-height: 52px; color: #fff; }
.div-sub { position: absolute; left: 100.8px; top: 303.2px;
           font-family: Calibri, sans-serif; font-size: 14px; line-height: 17.1px; color: #c7cee0; }
.spbox { position: absolute; left: 101.5px; top: 346.5px; width: 182px;
         border: 1px solid #4a5d85; padding: 13px 21.2px 15.8px; }
.spbox .lbl { font-family: Calibri, sans-serif; font-size: 10px; line-height: 14px;
              letter-spacing: 2px; white-space: nowrap; color: #9fb0d1; }
.spbox .val { font-family: Georgia, serif; font-weight: bold; font-size: 30px; line-height: 30px;
              margin-top: 6px; }
.spbox .val.pos { color: #7fd39a; }
.spbox .val.neg { color: #e8867d; }
.div-listhead { position: absolute; left: 100.8px; top: 456.2px; font-size: 11px; line-height: 13.4px;
                letter-spacing: .183em; }
.div-list { position: absolute; left: 100.8px; top: 483.8px; margin: 0; padding: 0; list-style: none;
            font-family: Calibri, sans-serif; font-size: 13px; line-height: 15.9px; color: #e4e9f3; }
.div-list li { height: 26px; }
.div-foot { position: absolute; left: 100.8px; right: 100.8px; top: 755.7px; display: flex;
            justify-content: space-between; font-family: Calibri, sans-serif; font-size: 10px;
            line-height: 12.2px; color: #8592ac; }

/* ---- banded pages ---- */
.band { position: absolute; left: 0; top: 0; width: 100%; height: 83px;
        background: linear-gradient(135deg, #223a63 0%, #101d35 100%);
        border-bottom: 3px solid #b08c57;
        display: flex; align-items: flex-end; justify-content: space-between;
        padding: 0 52.8px 12.1px; }
.band h1 { margin: 0; font-family: Georgia, serif; font-weight: bold; font-size: 25px; line-height: 25px;
           color: #fff; }
.band-meta { text-align: right; font-family: Calibri, sans-serif; font-size: 11px; line-height: 17px;
             color: #c7cee0; }
.band-meta b { color: #fff; }

.band-flat { position: absolute; left: 0; top: 0; width: 100%; height: 114px; background: #17294a;
             padding: 0 57.6px 30.8px; display: flex; flex-direction: column; justify-content: flex-end; }
.band-flat .eyebrow { font-size: 11px; line-height: 16px; }
.band-flat h1 { margin: 0; font-family: Georgia, serif; font-weight: bold; font-size: 28px; line-height: 28px;
                color: #fff; }

/* ---- data table ---- */
.tbl-wrap { position: absolute; left: 48px; top: 102.9px; width: 960px; }
table.tbl { width: 960px; border-collapse: collapse; table-layout: fixed; }
table.tbl th { background: #1a2f52; color: #fff; font-family: Calibri, sans-serif; font-weight: bold;
               font-size: 9.5px; line-height: 14px; letter-spacing: 1px; padding: 6px 10px;
               text-align: left; }
table.tbl th.r { text-align: right; }
col.c1 { width: 16.97%; } col.c2 { width: 9.07%; } col.c3 { width: 8.96%; }
col.c4 { width: 8.03%; } col.c5 { width: 8.01%; } col.c6 { width: 48.96%; }

tbody tr:nth-child(odd) td { background: #f7f8fc; }
tr.sect td { background: #dbe2ee !important; color: #243356; font-family: Calibri, sans-serif;
             font-weight: bold; font-size: 10px; line-height: 14px; letter-spacing: 1.5px;
             text-transform: uppercase; padding: 6px 10px; }
tbody td { padding: 6.6px 10px; border-bottom: 1px solid #e6e9f0; vertical-align: top; }
td.tick { font-family: Georgia, serif; font-weight: bold; font-size: 12.5px; line-height: 15px; color: #182a4a; }
td.tick .co { font-family: Calibri, sans-serif; font-weight: normal; font-size: 9.5px; line-height: 11px;
              color: #6b768e; }
td.rate { line-height: 0; padding-top: 7.5px; }
.pill { display: inline-block; padding: 1px 6.7px; border: 1px solid #b9c2d4; border-radius: 3px;
        background: #eef1f7; font-family: Calibri, sans-serif; font-weight: bold; font-size: 9.5px;
        line-height: 13px; color: #243356; white-space: nowrap; }
.pill.muted { background: #f7f8fc; border-color: #dde1ea; }
td.num { text-align: right; font-family: Georgia, serif; font-size: 11.5px; line-height: 13px; color: #1a2333; }
td.num.b { font-weight: bold; }
td.note { font-family: Calibri, sans-serif; font-size: 11.5px; line-height: 14px; color: #333d54; }
td.pos, .pos { color: #1a7f37; }
td.neg, .neg { color: #c0392b; }

.foot { position: absolute; left: 0; bottom: 0; width: 100%; height: 44px; padding: 2px 52.8px 0;
        border-top: 1px solid #e6e9f0; display: flex; align-items: center;
        justify-content: space-between; font-family: Calibri, sans-serif; font-size: 9.5px;
        line-height: 11.6px; color: #8592ac; }
.foot .firm { font-weight: bold; color: #4a5570; }

/* ---- contents ---- */
table.toc { position: absolute; left: 58px; top: 172px; width: 940px;
            border-collapse: collapse; table-layout: fixed; }
table.toc th { background: #1a2f52; color: #fff; font-family: Calibri, sans-serif; font-weight: bold;
               font-size: 10.5px; line-height: 12.8px; letter-spacing: 1px; padding: 10.15px 11.6px 8.15px;
               text-align: left; }
table.toc th.r, table.toc td.r { text-align: right; }
col.t1 { width: 16%; } col.t2 { width: 60%; } col.t4 { width: 24%; }
table.toc tbody tr:nth-child(odd) td { background: #f6f8fb; }
table.toc td { padding: 13.1px 11.6px 11.1px; border-bottom: 1px solid #e2e6ee; vertical-align: top; }
td.appt { font-family: Georgia, serif; font-weight: bold; font-size: 13px; line-height: 14.7px; color: #182a4a; }
td.theme { font-family: Calibri, sans-serif; font-size: 13px; line-height: 15.9px; color: #4a5570; }
td.spx { font-family: Georgia, serif; font-weight: bold; font-size: 13px; line-height: 14.7px; }
td.pages { font-family: Calibri, sans-serif; font-size: 13px; line-height: 15.9px; color: #8592ac; }

/* ---- recap ---- */
table.score { position: absolute; left: 58px; top: 190px; width: 940px;
              border-collapse: collapse; table-layout: fixed; }
table.score th { background: #1a2f52; color: #fff; font-family: Calibri, sans-serif; font-weight: bold;
                 font-size: 11px; line-height: 13.4px; letter-spacing: 1px;
                 padding: 12px 16px 10px; text-align: left; }
table.score th.r, table.score td.r { text-align: right; }
col.s1 { width: 34%; } col.s2 { width: 33%; } col.s3 { width: 33%; }
table.score tbody tr:nth-child(odd) td { background: #f6f8fb; }
table.score td { padding: 20px 16px 18px; border-bottom: 1px solid #e2e6ee; vertical-align: middle; }
table.score td.appt { font-family: Georgia, serif; font-weight: bold; font-size: 16px; line-height: 18px;
                      color: #182a4a; }
table.score td.n { font-family: Georgia, serif; font-size: 18px; line-height: 20px; color: #1a2333; }
table.score td.g { font-family: Georgia, serif; font-weight: bold; font-size: 18px; line-height: 20px; }
table.score tr.tot td { background: #eef1f7 !important; border-top: 2px solid #1a2f52;
                        border-bottom: none; }
table.score tr.tot td.appt, table.score tr.tot td.n { color: #182a4a; font-weight: bold; }

.score-note { position: absolute; left: 57.6px; top: 690px; width: 900px;
              font-family: Calibri, sans-serif; font-size: 10.5px; line-height: 15.5px; color: #6b768e; }
"""


def page_dark_cover(as_of):
    return """
<div class="page dark">
  <div class="seal-wrap">{logo}</div>

  <h1 class="cover-h1">{client}</h1>
  <div class="cover-sub">2026 Recommendations Review &nbsp;·&nbsp; Portfolio Reviews 1 &ndash; 6</div>
  <div class="cover-rule"></div>
  <div class="cover-desc">A record of reviews of recommendations, investment theses, and
    performance versus the S&amp;P 500 TR Index.</div>
  <div class="cover-foot">Prepared for Client Review &nbsp;|&nbsp; Data as of {as_of} &nbsp;|&nbsp; Source: YCharts</div>
</div>""".format(logo=logo_tag(), client=CLIENT, as_of=as_of)


def page_contents(appts, page_map, as_of):
    rows = []
    for a in appts:
        first, last = page_map[a["num"]]
        span = "Page {}".format(first) if first == last else "Pages {}&ndash;{}".format(first, last)
        rows.append("""
      <tr>
        <td class="appt">Portfolio Review {n}</td>
        <td class="theme">{theme}</td>
        <td class="pages r">{span}</td>
      </tr>""".format(n=a["num"], theme=a["theme"], span=span))

    return """
<div class="page">
  <div class="band-flat">
    <h1>Contents</h1>
  </div>
  <table class="toc">
    <colgroup><col class="t1"><col class="t2"><col class="t4"></colgroup>
    <thead><tr><th>Portfolio Review</th><th>Key Themes</th><th class="r">Pages</th></tr></thead>
    <tbody>{rows}
    </tbody>
  </table>
</div>""".format(rows="".join(rows))


def page_divider(appt, pageno, total):
    items = "".join("<li>&mdash;&nbsp; {}</li>".format(html.escape(n)) for n, _ in appt["sections"])
    return """
<div class="page dark">
  <h1 class="div-h1">Portfolio Review {n}</h1>
  <div class="div-sub">Recommendation date {rec} &nbsp;·&nbsp; measured through {asof}</div>
  <div class="spbox">
    <div class="lbl">S&amp;P 500 TOTAL RETURN</div>
    <div class="val {cls}">{spx}</div>
  </div>
  <div class="eyebrow div-listhead">IN THIS SECTION</div>
  <ul class="div-list">{items}</ul>
  <div class="div-foot"><span>{firm}</span><span>Page {p} of {t}</span></div>
</div>""".format(n=appt["num"], rec=fmt_date(appt["rec_date"]), asof=fmt_date(appt["as_of"]),
                 cls=pct_class(appt["bench"]), spx=pct(appt["bench"]), items=items,
                 firm=FIRM, p=pageno, t=total)


def page_table(appt, sections, idx, nsub, pageno, total):
    body = []
    for name, rows in sections:
        body.append('<tr class="sect"><td colspan="6">{}</td></tr>'.format(html.escape(name)))
        for r in rows:
            pill = '<span class="pill">{}</span>'.format(html.escape(r["rating"]))
            body.append("""
      <tr>
        <td class="tick">{tk}<div class="co">{co}</div></td>
        <td class="rate">{pill}</td>
        <td class="num">{px}</td>
        <td class="num b {rc}">{rv}</td>
        <td class="num b {vc}">{vs}</td>
        <td class="note">{nt}</td>
      </tr>""".format(tk=html.escape(r["ticker"]), co=html.escape(r["name"]), pill=pill,
                      px=money(r["price"]),
                      rc=pct_class(r["ret"]), rv=pct(r["ret"]),
                      vc=pct_class(r["vs"]), vs=pct(r["vs"]),
                      nt=html.escape(r["notes"])))

    return """
<div class="page">
  <div class="band">
    <h1>Portfolio Review {n} ({i} of {ns})</h1>
    <div class="band-meta">
      <div>Total return from <b>{rec}</b> through <b>{asof}</b></div>
      <div>S&amp;P 500 Total Return <b>{spx}</b></div>
    </div>
  </div>
  <div class="tbl-wrap">
    <table class="tbl">
      <colgroup><col class="c1"><col class="c2"><col class="c3"><col class="c4"><col class="c5"><col class="c6"></colgroup>
      <thead><tr><th>Name</th><th>Rating</th><th class="r">Price at Rec.</th><th class="r">Return</th>
        <th class="r">vs. S&amp;P</th><th>Notes</th></tr></thead>
      <tbody>{body}
      </tbody>
    </table>
  </div>
  <div class="foot"><span class="firm">{firm}</span>
    <span>Total return basis&nbsp;|&nbsp;Source: YCharts</span>
    <span>Page {p} of {t}</span></div>
</div>""".format(n=appt["num"], i=idx, ns=nsub, rec=fmt_date(appt["rec_date"]),
                 asof=fmt_date(appt["as_of"]), spx=pct(appt["bench"]), body="".join(body),
                 firm=FIRM, p=pageno, t=total)


def stats(appt):
    """Simple (unweighted) averages across every name discussed at an appointment."""
    rows = [r for _, rs in appt["sections"] for r in rs]
    n = len(rows)
    return {
        "n": n,
        "ret": sum(r["ret"] for r in rows) / n,
        "vs": sum(r["vs"] for r in rows) / n,
        "hit": sum(1 for r in rows if r["vs"] > 0) / float(n),
    }


def page_recap(appts, as_of):
    body, tot_n, tot_r = [], 0, 0.0
    for a in appts:
        s = stats(a)
        tot_n += s["n"]
        tot_r += s["ret"] * s["n"]
        body.append("""
      <tr>
        <td class="appt">Appointment {n}</td>
        <td class="n r">{spx}</td>
        <td class="g r {rc}">{ret}</td>
      </tr>""".format(n=a["num"], spx=pct(a["bench"]),
                      rc=pct_class(s["ret"]), ret=pct(s["ret"])))

    avg_r = tot_r / tot_n
    body.append("""
      <tr class="tot">
        <td class="appt">All appointments</td>
        <td class="n r">&mdash;</td>
        <td class="g r {rc}">{ret}</td>
      </tr>""".format(rc=pct_class(avg_r), ret=pct(avg_r)))

    return """
<div class="page">
  <div class="band-flat">
    <div class="eyebrow">2026 YEAR-TO-DATE</div>
    <h1>Performance Recap</h1>
  </div>
  <table class="score">
    <colgroup><col class="s1"><col class="s2"><col class="s3"></colgroup>
    <thead><tr><th>Appointment</th><th class="r">S&amp;P 500 Total Return</th>
      <th class="r">Our Average Return</th></tr></thead>
    <tbody>{body}
    </tbody>
  </table>
  <div class="score-note">Our Average Return is the simple, unweighted mean total return of the names discussed at
    each appointment, measured from that appointment&rsquo;s recommendation date through {asof}. The S&amp;P 500
    Total Return Index (^SPXTR) is measured over the identical window. Source: YCharts.</div>
</div>""".format(body="".join(body), asof=as_of)


def build_html(appts):
    as_of = fmt_date(appts[0]["as_of"])

    # First pass: work out page numbering.
    paged = [(a, paginate(a)) for a in appts]
    total = 2 + sum(1 + len(p) for _, p in paged)
    page_map, n = {}, 2
    for a, pages in paged:
        first = n + 1
        n += 1 + len(pages)
        page_map[a["num"]] = (first, n)

    parts = [page_dark_cover(as_of), page_contents(appts, page_map, as_of)]
    n = 2
    for a, pages in paged:
        n += 1
        parts.append(page_divider(a, n, total))
        for i, sections in enumerate(pages, start=1):
            n += 1
            parts.append(page_table(a, sections, i, len(pages), n, total))

    # The closing performance-recap page was dropped at the client's request.
    # To bring it back, append page_recap(appts, as_of) here and add 1 to `total`.

    return """<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>Mudry &amp; Rinaldi - 2026 Portfolio Reviews</title>
<style>{css}</style></head><body>{body}</body></html>""".format(css=CSS, body="".join(parts))


def main():
    appts = load_appointments()

    # Every ticker must carry a rating; a new name in the workbook with no explicit
    # rating needs an entry in INFERRED_RATINGS rather than shipping a blank cell.
    missing = ["Portfolio Review {} / {}".format(a["num"], r["ticker"])
               for a in appts for _, rs in a["sections"] for r in rs if r["rating"] == "—"]
    if missing:
        raise RuntimeError("No rating for: " + ", ".join(missing) +
                           " - add them to INFERRED_RATINGS.")

    doc = build_html(appts)

    html_path = os.path.join(BASE, "layout", "_stock_picks_build.html")
    with open(html_path, "w", encoding="utf-8") as fh:
        fh.write(doc)

    as_of = appts[0]["as_of"]
    if isinstance(as_of, datetime.datetime):
        as_of = as_of.date()
    if not os.path.isdir(OUT_DIR):
        os.makedirs(OUT_DIR)
    # Matches the delivered filename so a re-run refreshes that file rather than
    # dropping a second copy alongside it.
    pdf_path = os.path.join(OUT_DIR, "Mudry-Rinaldi Summer Review.pdf")

    # Chrome exits 0 even when it cannot write the target (e.g. the PDF is open in a
    # viewer), so render to a temp file and verify before replacing the real one.
    tmp_pdf = os.path.join(tempfile.mkdtemp(prefix="deckbuild_"), "render.pdf")
    subprocess.check_call([
        CHROME, "--headless", "--disable-gpu", "--no-pdf-header-footer",
        "--print-to-pdf=" + tmp_pdf, "--no-margins",
        "file:///" + html_path.replace("\\", "/").replace(" ", "%20"),
    ])
    if not os.path.exists(tmp_pdf) or os.path.getsize(tmp_pdf) < 50000:
        raise RuntimeError("Chrome did not produce a usable PDF at " + tmp_pdf)

    try:
        shutil.copyfile(tmp_pdf, pdf_path)
    except OSError:
        raise RuntimeError(
            "Could not overwrite {} - it is open in another program. "
            "Close it and re-run.".format(pdf_path))
    finally:
        shutil.rmtree(os.path.dirname(tmp_pdf), ignore_errors=True)

    print("HTML:", html_path)
    print("PDF :", pdf_path)


if __name__ == "__main__":
    main()
