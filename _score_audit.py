"""Audit: compare dashboard Auto_Score (outputs.csv) vs hero score in each HTML report.

F-R extension: also diff GG_Price / EM_Price / Price_Target across:
  outputs.csv  ↔  static/data/{TICKER}_data.json  ↔  HTML hero

Run: python _score_audit.py [--full]
  --full  also shows GG/EM/PT column audit (always runs for mismatches).
"""
import csv, io, json, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import csv_schema as _schema

ROOT     = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(ROOT, "outputs.csv")
RPT_DIR  = os.path.join(ROOT, "static", "reports")
DAT_DIR  = os.path.join(ROOT, "static", "data")

FULL_MODE = "--full" in sys.argv

with open(CSV_PATH, "r", encoding="utf-8") as f:
    content = _schema.migrate(f.read())
csv_rows = {r["Ticker"]: r for r in csv.DictReader(io.StringIO(content)) if r.get("Ticker")}

print(f"{'Ticker':<6}  {'Dashboard':>10}  {'Report':>8}  {'':>6}  Notes")
print("-" * 70)
mismatches = 0
valuation_diffs = []

for ticker in sorted(csv_rows):
    dash_score = csv_rows[ticker].get("Auto_Score", "").strip()
    rpt_path   = os.path.join(RPT_DIR, f"{ticker}_report.html")
    dat_path   = os.path.join(DAT_DIR, f"{ticker}_data.json")

    if not os.path.exists(rpt_path):
        print(f"{ticker:<6}  {dash_score:>10}  {'NO REPORT':>8}  {'---':>6}")
        continue

    with open(rpt_path, "r", encoding="utf-8") as f:
        html = f.read()

    # ── Score audit (Rule 10) ──────────────────────────────────────────────
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

    # ── F-R: Valuation field diff (CSV ↔ JSON ↔ HTML) ────────────────────
    if (FULL_MODE or match == "MISMATCH") and os.path.exists(dat_path):
        try:
            with open(dat_path, "r", encoding="utf-8") as f:
                jdat = json.load(f)
            dp = jdat.get("dcf_prices") or {}

            # CSV values
            csv_gg = csv_rows[ticker].get("GG_Price", "").strip() or None
            csv_em = csv_rows[ticker].get("EM_Price", "").strip() or None

            # JSON values (engine output)
            json_gg = dp.get("gg_price")
            json_em = dp.get("em_price")

            # HTML price-value (the displayed price target)
            pt_m = re.search(r'price-value price-target[^>]*>([0-9.N/A]+)<', html)
            html_pt = pt_m.group(1) if pt_m else "N/F"

            # Summarise discrepancies
            gg_note = ""
            if csv_gg is not None and json_gg is not None:
                if abs(float(csv_gg) - float(json_gg)) > 1.0:
                    gg_note = f" GG: csv={csv_gg} json={json_gg:.2f} DIFF"
            em_note = ""
            if csv_em is not None and json_em is not None:
                if abs(float(csv_em) - float(json_em)) > 1.0:
                    em_note = f" EM: csv={csv_em} json={json_em:.2f} DIFF"

            if gg_note or em_note or FULL_MODE:
                line = f"       {'GG':>4}={csv_gg or 'N/A':>8}/{json_gg if json_gg else 'None':>8}  "
                line += f"{'EM':>4}={csv_em or 'N/A':>8}/{json_em if json_em else 'None':>8}  "
                line += f"HTML_PT={html_pt}"
                if gg_note or em_note:
                    line += f"  !!{gg_note}{em_note}"
                    valuation_diffs.append(ticker)
                print(line)
        except Exception as e:
            print(f"       [F-R audit error: {e}]")

print(f"\nMismatches: {mismatches} / {len(csv_rows)}")
if valuation_diffs:
    print(f"Valuation field diffs: {', '.join(valuation_diffs)}")
elif FULL_MODE:
    print("No GG/EM/PT field diffs detected.")
