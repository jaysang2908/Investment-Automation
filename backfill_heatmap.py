"""
backfill_heatmap.py
-------------------
One-shot script to populate outputs.csv from all previously run tickers.
Run locally: python backfill_heatmap.py

Requires:
  - FMP_API_KEY  set as environment variable  (or hardcode below)
  - GITHUB_TOKEN set as environment variable  (personal access token, repo scope)
  - pip install requests openpyxl yfinance
"""

import base64
import datetime
import os
import sys
import requests

# ── Config — edit if needed ───────────────────────────────────────────────────
FMP_API_KEY   = os.environ.get("FMP_API_KEY",   "")   # or paste key directly
GITHUB_TOKEN  = os.environ.get("GITHUB_TOKEN",  "")   # or paste token directly
GITHUB_REPO   = "jaysang2908/Investment-Automation"
GITHUB_BRANCH = "main"

TICKERS = [
    "AAPL", "ABBV", "ADBE", "AMD", "BAC",
    "COST", "CSCO", "C",    "F",   "INTC",
    "JNJ",  "JPM",  "KO",   "META","MSFT",
    "NVDA", "SOFI", "TSLA", "TSM", "V",
    "WMT",
]

CSV_HEADER = (
    "Ticker,Price,MktCap_B,ROIC,Rev_CAGR,FCF_NI,D_EBITDA,"
    "PE_Current,PE_5yr,PFCF_Current,PFCF_5yr,"
    "Auto_Score,Floor_Cap,Date\n"
)

# ── Load model module ─────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
import fmp_3statementv6 as mdl
mdl.API_KEY = FMP_API_KEY

# ── GitHub helpers ────────────────────────────────────────────────────────────
GH_HEADERS = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept":        "application/vnd.github.v3+json",
}
GH_API = f"https://api.github.com/repos/{GITHUB_REPO}/contents/outputs.csv"


def _read_csv_from_github():
    r = requests.get(GH_API, headers=GH_HEADERS, params={"ref": GITHUB_BRANCH}, timeout=8)
    if r.status_code == 200:
        info = r.json()
        return info["sha"], base64.b64decode(info["content"]).decode()
    return None, CSV_HEADER


def _write_csv_to_github(sha, content):
    payload = {
        "message": "Backfill heatmap from existing models",
        "branch":  GITHUB_BRANCH,
        "content": base64.b64encode(content.encode()).decode(),
    }
    if sha:
        payload["sha"] = sha
    r = requests.put(GH_API, headers=GH_HEADERS, json=payload, timeout=15)
    if r.status_code not in (200, 201):
        print(f"  ERROR writing CSV: {r.status_code} — {r.json().get('message','')}")
    else:
        print("  CSV written to GitHub.")


def _f(v, dp=4):
    return "" if v is None else f"{v:.{dp}f}"


# ── Main loop ─────────────────────────────────────────────────────────────────
def run():
    if not FMP_API_KEY:
        sys.exit("ERROR: FMP_API_KEY not set.")
    if not GITHUB_TOKEN:
        sys.exit("ERROR: GITHUB_TOKEN not set.")

    print(f"Reading existing outputs.csv from GitHub...")
    sha, content = _read_csv_from_github()

    # Find tickers already in the CSV so we don't duplicate
    existing = set()
    for line in content.splitlines()[1:]:
        if line.strip():
            existing.add(line.split(",")[0].strip())
    print(f"  Already present: {sorted(existing) or 'none'}")

    today = datetime.date.today().isoformat()
    rows_added = 0

    for ticker in TICKERS:
        if ticker in existing:
            print(f"  Skipping {ticker} — already in CSV")
            continue

        print(f"\nProcessing {ticker}...")
        try:
            is_data = mdl.fetch("income-statement",        ticker)[:mdl.YEARS][::-1]
            bs_data = mdl.fetch("balance-sheet-statement", ticker)[:mdl.YEARS][::-1]
            cf_data = mdl.fetch("cash-flow-statement",     ticker)[:mdl.YEARS][::-1]

            if not is_data:
                print(f"  No data — skipping.")
                continue

            years = [
                d.get("fiscalYear") or d.get("calendarYear") or d["date"][:4]
                for d in is_data
            ]

            # Fetch price + market cap
            price = None
            mkt_cap = None
            try:
                prof = requests.get(
                    f"https://financialmodelingprep.com/stable/profile"
                    f"?symbol={ticker}&apikey={FMP_API_KEY}", timeout=8
                ).json()
                rec     = prof[0] if isinstance(prof, list) else prof
                price   = float(rec.get("price")     or 0) or None
                mkt_cap = float(rec.get("mktCap") or rec.get("marketCap") or 0) or None
            except Exception as e:
                print(f"  Warning: could not fetch profile — {e}")

            # Build scorecard (we don't need the full workbook, but build_scorecard
            # requires wb to create a sheet — use a throwaway workbook)
            from openpyxl import Workbook
            wb = Workbook()
            mdl.build_cover(wb, ticker, years, is_data)
            pl_refs   = mdl.build_pl(wb, is_data, years, ticker)
            bs_refs   = mdl.build_bs(wb, bs_data, years, ticker)
            cf_refs   = mdl.build_cf(wb, cf_data, years, ticker)
            mdl.build_ratios(wb, is_data, bs_data, cf_data, years, ticker, pl_refs, bs_refs, cf_refs)
            mdl.build_segments(wb, ticker, years)
            wacc_refs = mdl.build_wacc(wb, ticker, is_data, bs_data, None)
            mdl.build_dcf(wb, ticker, is_data, bs_data, cf_data, years, pl_refs, bs_refs, wacc_refs,
                          current_price=price)
            _, m = mdl.build_scorecard(wb, ticker, is_data, bs_data, cf_data, years)

            mkt_cap_b = (mkt_cap / 1e9) if mkt_cap else None

            row = ",".join([
                ticker,
                _f(price,   2),
                _f(mkt_cap_b, 2),
                _f(m.get("roic")),
                _f(m.get("rev_cagr")),
                _f(m.get("fcf_ni")),
                _f(m.get("d_ebitda"), 2),
                _f(m.get("pe_current"),   1),
                _f(m.get("pe_5yr_avg"),   1),
                _f(m.get("pfcf_current"), 1),
                _f(m.get("pfcf_5yr_avg"), 1),
                "" if m.get("auto_score") is None else str(m["auto_score"]),
                "" if m.get("floor_cap")  is None else str(m["floor_cap"]),
                today,
            ]) + "\n"

            content += row
            rows_added += 1
            print(f"  Done — auto_score={m.get('auto_score')}  price={price}")

        except Exception as e:
            import traceback
            print(f"  ERROR for {ticker}: {e}")
            traceback.print_exc()

    if rows_added == 0:
        print("\nNo new rows to write.")
        return

    print(f"\nWriting {rows_added} new row(s) to GitHub...")
    _write_csv_to_github(sha, content)
    print("Done.")


if __name__ == "__main__":
    run()
