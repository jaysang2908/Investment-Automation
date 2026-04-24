"""
restyle_reports.py — swap CSS + inject topbar into pre-rendered reports
that were generated with the old IBM Plex / dark-navy theme.

Run once after updating Report_Template.html. Does NOT need FMP API.
"""
import os, re, datetime

REPORTS_DIR = "static/reports"
TEMPLATE_PATH = "Report_Template.html"

# Tickers whose reports were generated from Excel (already use new template — skip)
EXCEL_TICKERS = {"WFC", "INTC", "TSLA", "SOFI", "JPM", "C", "BAC", "UAL"}

TODAY = datetime.date.today().strftime("%B %Y")

# ── Load new head + nav elements from template ──────────────────────────────
with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
    template = f.read()

head_match = re.search(r"<head>(.*?)</head>", template, re.DOTALL)
new_head_inner = head_match.group(1)

# Only extract <nav class="topbar">...</nav> (NOT the hero section)
nav_match = re.search(r'(<nav class="topbar">.*?</nav>)', template, re.DOTALL)
new_topbar_nav = nav_match.group(1).strip() if nav_match else ""

# Extract tab-nav (has no data placeholders — pure anchor links)
tabnav_match = re.search(r'(<nav class="tab-nav[^"]*">.*?</nav>)', template, re.DOTALL)
new_tabnav = tabnav_match.group(1).strip() if tabnav_match else ""

# ── Process each report ──────────────────────────────────────────────────────
updated = 0
skipped = 0

for fname in sorted(os.listdir(REPORTS_DIR)):
    if not fname.endswith("_report.html"):
        continue
    ticker = fname.replace("_report.html", "")
    if ticker in EXCEL_TICKERS:
        skipped += 1
        continue

    path = os.path.join(REPORTS_DIR, fname)
    with open(path, "r", encoding="utf-8") as f:
        html = f.read()

    # ── 1. Extract title BEFORE replacing head ─────────────────────────────
    title_match = re.search(r"<title>(.*?)</title>", html)
    existing_title = title_match.group(1) if title_match else ticker
    company_part = (existing_title
                    .replace(" — Independent Research Report", "")
                    .replace(" — Equity Research", "")
                    .strip())
    # If it's already a placeholder (failed earlier restyle), fall back to ticker
    if "{{" in company_part:
        company_part = ticker

    # ── 2. Replace <head> with new CSS ────────────────────────────────────
    html = re.sub(r"<head>.*?</head>", f"<head>{new_head_inner}</head>", html, flags=re.DOTALL)
    html = re.sub(r"<title>.*?</title>",
                  f"<title>{company_part} — Equity Research</title>", html)

    # ── 3. Extract values from old rendered content for nav placeholders ──
    # CURRENT_PRICE: <span class="current">273.43</span>
    price_match = re.search(r'<span class="current">([\d.]+)</span>', html)
    current_price = price_match.group(1) if price_match else "—"

    # ── 4. Build filled topbar nav (3 placeholders only) ──────────────────
    filled_nav = (new_topbar_nav
                  .replace("{{TICKER_SHORT}}", ticker)
                  .replace("{{CURRENT_PRICE}}", current_price)
                  .replace("{{REPORT_DATE}}", TODAY))

    # ── 5. Inject nav + tab-nav after <body> (only if not already done) ──
    if "<!-- ═══ TOPBAR ═══ -->" not in html:
        inject = f"\n<!-- ═══ TOPBAR ═══ -->\n{filled_nav}\n{new_tabnav}\n"
        html = html.replace("<body>", f"<body>{inject}", 1)
    else:
        # Re-inject: strip old nav block and replace with fresh one
        html = re.sub(
            r'<!-- ═══ TOPBAR ═══ -->.*?(?=<!-- ═══ HEADER|<header|<div class="main)',
            f"<!-- ═══ TOPBAR ═══ -->\n{filled_nav}\n{new_tabnav}\n",
            html, flags=re.DOTALL
        )

    # ── 6. Wrap body content in <div class="main"> if not already ─────────
    if 'class="main"' not in html and '<main' not in html:
        body_content_start = html.find("<!-- ═══ HEADER ═══ -->")
        if body_content_start == -1:
            body_content_start = html.find('<header class="report-header">')
        if body_content_start == -1:
            # Find end of tabnav
            end_of_tabnav = html.rfind("</nav>", 0, html.find('<header') if html.find('<header') > 0 else len(html))
            if end_of_tabnav != -1:
                body_content_start = end_of_tabnav + len("</nav>")

        body_end = html.rfind("</body>")
        if body_content_start != -1 and body_end != -1:
            old_content = html[body_content_start:body_end]
            html = (html[:body_content_start]
                    + '\n<div class="main">\n'
                    + old_content
                    + '\n</div><!-- end main -->\n'
                    + html[body_end:])

    # ── 7. Fix ROIC chart: replace hardcoded max:45 with dynamic suggestedMax ──
    html = html.replace(
        "min: 0, max: 45,",
        "min: 0, suggestedMax: Math.ceil(Math.max(...roicData, waccValue) * 1.25 / 10) * 10,"
    )

    # ── 8. Fix EBITINT_NOTE placeholder if still unresolved ──────────────────
    if "{{EBITINT_NOTE}}" in html:
        # Try to extract per-year EBIT/Int values already rendered in the old content
        # Look for DEBITDA rows to infer the structure (EBIT/Int was rendered per-year in old template)
        # Fallback: show "See annual report for interest coverage by year"
        html = html.replace("{{EBITINT_NOTE}}",
                            "See annual report — interest coverage varies by year")

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"  Updated {ticker}  (price={current_price})")
    updated += 1

print(f"\nDone: {updated} updated, {skipped} skipped (Excel-based)")
