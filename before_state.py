"""Capture pre-rerun state for before/after comparison table."""
import json, os, re, csv, io
import csv_schema as _schema

ROOT    = os.path.dirname(os.path.abspath(__file__))
DAT     = os.path.join(ROOT, "static", "data")
RPT     = os.path.join(ROOT, "static", "reports")
CSV_PATH = os.path.join(ROOT, "outputs.csv")

TICKERS = ["AAPL","ABBV","ADBE","AMD","COST","DIS","FDX","HCA","KO","NKE",
           "NVDA","PEP","TGT","TSLA","WMT"]

with open(CSV_PATH, encoding="utf-8") as f:
    content = _schema.migrate(f.read())
csv_rows = {r["Ticker"]: r for r in csv.DictReader(io.StringIO(content)) if r.get("Ticker")}

print(f"{'Ticker':<6}  {'Score':>6}  {'WACC%':>7}  {'GG $':>8}  {'GG%':>7}  {'EM $':>8}  {'EM%':>7}  {'PrimaryPT':>10}  Notes")
print("-" * 110)

for t in TICKERS:
    path = os.path.join(DAT, t + "_data.json")
    if not os.path.exists(path):
        print(f"{t:<6}  -- NO JSON --")
        continue

    with open(path) as f: d = json.load(f)
    dp  = d.get("dcf_prices", {})
    sm  = d.get("scorecard_metrics", {}) or {}
    csv_r = csv_rows.get(t, {})

    score   = csv_r.get("Auto_Score", "?")
    wacc    = dp.get("wacc_val") or (d.get("wacc_refs") or {}).get("wacc_val")
    gg_px   = dp.get("gg_price")
    gg_up   = dp.get("gg_upside")
    em_px   = dp.get("em_price")
    em_up   = dp.get("em_upside")
    wacc_fl = dp.get("wacc_floored", False)

    # Derive primary PT from HTML
    rpt_path = os.path.join(RPT, t + "_report.html")
    pt_str = "?"
    if os.path.exists(rpt_path):
        html = open(rpt_path, encoding="utf-8").read()
        m = re.search(r'price-value price-target[^>]*>([^<]+)', html)
        if m: pt_str = m.group(1).strip()

    wacc_s  = f"{wacc*100:.1f}%{'F' if wacc_fl else ''}" if wacc else "?"
    gg_s    = f"${gg_px:.0f}" if gg_px else "N/A"
    gg_u    = f"{gg_up*100:+.0f}%" if gg_up is not None else "N/A"
    em_s    = f"${em_px:.0f}" if em_px else "N/A"
    em_u    = f"{em_up*100:+.0f}%" if em_up is not None else "N/A"

    notes = []
    if dp.get("wacc_floored"): notes.append("WACC-floored")
    if dp.get("em_anchored"):  notes.append("anchored")
    if dp.get("em_hist_anchor_capped"): notes.append("capped@28x")
    if dp.get("quality_em_premium"):    notes.append("qual-prem")
    if not dp.get("growth_tier"):       notes.append("STALE-JSON")

    print(f"{t:<6}  {score:>6}  {wacc_s:>7}  {gg_s:>8}  {gg_u:>7}  {em_s:>8}  {em_u:>7}  {pt_str:>10}  {', '.join(notes)}")
