"""
restyle_reports.py — swap CSS + inject topbar into pre-rendered reports
that were generated with the old IBM Plex / dark-navy theme.

Run once after updating Report_Template.html. Does NOT need FMP API.
"""
import os, re

REPORTS_DIR = "static/reports"
TEMPLATE_PATH = "Report_Template.html"

# Tickers whose reports were generated from Excel (already use new template — skip)
EXCEL_TICKERS = {"WFC", "INTC", "TSLA", "SOFI", "JPM", "C", "BAC", "UAL"}

# ── Load new head + topbar from template ────────────────────────────────────
with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
    template = f.read()

head_match = re.search(r"<head>(.*?)</head>", template, re.DOTALL)
new_head_inner = head_match.group(1)

# Topbar + tab-nav block (between <body> tag and <!-- ═══ MAIN CONTENT ═══ -->)
topbar_match = re.search(r"<body>(.*?)<!-- ═══ MAIN CONTENT", template, re.DOTALL)
new_topbar = topbar_match.group(1).strip()

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

    # Replace <head>...</head>
    html = re.sub(r"<head>.*?</head>", f"<head>{new_head_inner}</head>", html, flags=re.DOTALL)

    # Replace the placeholder title (the template has {{COMPANY_NAME}} but the rendered file has real name)
    # Extract existing title to preserve it
    title_match = re.search(r"<title>(.*?)</title>", html)
    if title_match:
        existing_title = title_match.group(1)
        # Update title to new style (strip old suffix)
        company_part = existing_title.replace(" — Independent Research Report", "").replace(" — Equity Research", "").strip()
        html = re.sub(r"<title>.*?</title>", f"<title>{company_part} — Equity Research</title>", html)

    # Inject topbar after <body> tag (before existing content)
    if "<!-- ═══ TOPBAR ═══ -->" not in html:
        html = html.replace("<body>", f"<body>\n{new_topbar}", 1)

    # Wrap existing body content in <main class="main"> if not already wrapped
    # (old reports have bare body content; add a thin wrapper for padding)
    if 'class="main"' not in html and '<main' not in html:
        # Find the end of topbar injection and wrap rest
        body_content_start = html.find("<!-- ═══ HEADER ═══ -->")
        if body_content_start == -1:
            body_content_start = html.find('<header class="report-header">')
        if body_content_start == -1:
            body_content_start = html.find("<!-- ═══ TOPBAR ═══ -->")
            if body_content_start != -1:
                # skip past topbar block
                end_of_topbar = html.find("</nav>", body_content_start)
                if end_of_topbar != -1:
                    body_content_start = end_of_topbar + len("</nav>")

        # Find </body>
        body_end = html.rfind("</body>")

        if body_content_start != -1 and body_end != -1:
            old_content = html[body_content_start:body_end]
            html = (html[:body_content_start]
                    + '\n<div class="main">\n'
                    + old_content
                    + '\n</div><!-- end main -->\n'
                    + html[body_end:])

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"  Updated {ticker}")
    updated += 1

print(f"\nDone: {updated} updated, {skipped} skipped (Excel-based, already new template)")
