"""Audit: compare dashboard Auto_Score (outputs.csv) vs hero score in each HTML report."""
import csv, io, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import csv_schema as _schema

ROOT     = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(ROOT, "outputs.csv")
RPT_DIR  = os.path.join(ROOT, "static", "reports")

with open(CSV_PATH, "r", encoding="utf-8") as f:
    content = _schema.migrate(f.read())
csv_rows = {r["Ticker"]: r for r in csv.DictReader(io.StringIO(content)) if r.get("Ticker")}

print(f"{'Ticker':<6}  {'Dashboard':>10}  {'Report':>8}  {'':>6}  Notes")
print("-" * 70)
mismatches = 0
for ticker in sorted(csv_rows):
    dash_score = csv_rows[ticker].get("Auto_Score", "").strip()
    rpt_path   = os.path.join(RPT_DIR, f"{ticker}_report.html")
    if not os.path.exists(rpt_path):
        print(f"{ticker:<6}  {dash_score:>10}  {'NO REPORT':>8}  {'---':>6}")
        continue
    with open(rpt_path, "r", encoding="utf-8") as f:
        html = f.read()
    m = re.search(r'verdict-score-num[^>]*>([0-9.]+)<', html)
    rpt_score = m.group(1) if m else "N/F"
    match = "OK" if dash_score == rpt_score else "MISMATCH"
    if match == "MISMATCH":
        mismatches += 1
    bc  = csv_rows[ticker].get("Manual_Clarity", "") or ""
    ltp = csv_rows[ticker].get("Manual_LTP", "") or ""
    notes = f"quals BC={bc} LTP={ltp}" if (bc or ltp) else ""
    flag = "!!!" if match == "MISMATCH" else "   "
    print(f"{ticker:<6}  {dash_score:>10}  {rpt_score:>8}  {flag}  {notes}")

print(f"\nMismatches: {mismatches} / {len(csv_rows)}")
