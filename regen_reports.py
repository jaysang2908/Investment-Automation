"""
regen_reports.py
Re-renders all HTML reports from the cached data store (no FMP API calls).
Use this after updating report_bridge.py commentary or Report_Template.html.

Run with:  python regen_reports.py
"""
import os, sys, json, traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data_store import load_ticker_data
from report_bridge import build_report_data, render_html_report

REPORTS_DIR = os.path.join("static", "reports")
DATA_DIR    = os.path.join("static", "data")

# Load any saved qualitative overrides so adj_score is preserved
QUAL_PATH = os.path.join("static", "data", "qualitative_overrides.json")
def _load_qual():
    if os.path.exists(QUAL_PATH):
        try:
            with open(QUAL_PATH) as f:
                return json.load(f)
        except Exception:
            pass
    return {}

qual_overrides = _load_qual()

# Discover all cached tickers
data_files = sorted(f for f in os.listdir(DATA_DIR) if f.endswith("_data.json"))
tickers = [f.replace("_data.json", "") for f in data_files]

print(f"Found {len(tickers)} cached tickers: {', '.join(tickers)}\n")

ok, failed = [], []

for ticker in tickers:
    try:
        stored = load_ticker_data(ticker)
        if not stored:
            print(f"  {ticker}: no data file — skip")
            continue

        profile           = stored.get("profile") or {}
        is_data           = stored.get("is_data") or []
        bs_data           = stored.get("bs_data") or []
        cf_data           = stored.get("cf_data") or []
        years             = stored.get("years") or []
        wacc_val          = stored.get("wacc_val")
        dcf_prices        = stored.get("dcf_prices") or {}
        scorecard_metrics = stored.get("scorecard_metrics") or {}
        analyst_ests      = stored.get("analyst_ests") or []

        if not is_data or not years:
            print(f"  {ticker}: missing financial data — skip")
            continue

        # Pull qual overrides if saved
        ov = qual_overrides.get(ticker, {})
        biz_clarity = ov.get("biz_clarity") or None
        ltp         = ov.get("ltp") or None
        adj_score   = ov.get("adj_score") or None

        data = build_report_data(
            ticker        = ticker,
            profile       = profile,
            is_data       = is_data,
            bs_data       = bs_data,
            cf_data       = cf_data,
            years         = years,
            wacc_val      = wacc_val,
            dcf_prices    = dcf_prices,
            scorecard_metrics = scorecard_metrics,
            biz_clarity   = biz_clarity,
            ltp           = ltp,
            adj_score     = adj_score,
            analyst_ests  = analyst_ests,
        )

        html = render_html_report(data)

        out_path = os.path.join(REPORTS_DIR, f"{ticker}_report.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)

        score = data.get("FINAL_SCORE", "?")
        print(f"  {ticker}: OK  (score={score})")
        ok.append(ticker)

    except Exception as e:
        print(f"  {ticker}: FAILED — {e}")
        traceback.print_exc()
        failed.append(ticker)

print(f"\n{'='*50}")
print(f"Done: {len(ok)} regenerated, {len(failed)} failed")
if failed:
    print(f"Failed: {', '.join(failed)}")
