"""
seed_data_store.py
One-time script to populate static/data/TICKER_data.json for all 24 core tickers
(and any extras in the reports directory) so the DCF calculator never hits FMP
at request time.

Usage:  python seed_data_store.py
        python seed_data_store.py AAPL MSFT   # seed specific tickers only

Costs:  5 FMP calls per ticker (IS + BS + CF + Profile + Analyst estimates)
        24 tickers = 120 calls total (~3 min with polite sleep)
"""

import os, sys, time, datetime
import requests as _req

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fmp_3statementv6 as mdl
from data_store import save_ticker_data, load_ticker_data

CORE_TICKERS = [
    "NVDA","MSFT","AAPL","ADBE","COST","AMD","JNJ","META","TSM","V",
    "KO","NFLX","ABBV","CSCO","WMT","F","WFC","INTC","TSLA","SOFI",
    "JPM","C","BAC","UAL",
]

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "static", "reports")


def _discovered_tickers():
    """Return tickers that have generated reports but aren't in the core list."""
    if not os.path.isdir(REPORTS_DIR):
        return []
    tickers = []
    for fname in os.listdir(REPORTS_DIR):
        if fname.endswith("_report.html"):
            t = fname.replace("_report.html", "")
            if t not in CORE_TICKERS:
                tickers.append(t)
    return sorted(tickers)


def seed_ticker(ticker, force=False):
    """
    Fetch and store data for one ticker.
    Skips if already stored and not force=True.
    Returns True on success, False on failure/skip.
    """
    if not force and load_ticker_data(ticker):
        print(f"  {ticker:6s}  already cached — skip")
        return True

    try:
        is_data = mdl.fetch("income-statement",        ticker)[:mdl.YEARS][::-1]
        bs_data = mdl.fetch("balance-sheet-statement", ticker)[:mdl.YEARS][::-1]
        cf_data = mdl.fetch("cash-flow-statement",     ticker)[:mdl.YEARS][::-1]

        if not is_data:
            print(f"  {ticker:6s}  SKIP — no income statement data")
            return False

        profile = {}
        try:
            _p = _req.get(
                f"https://financialmodelingprep.com/stable/profile"
                f"?symbol={ticker}&apikey={mdl.API_KEY}", timeout=8
            ).json()
            profile = (_p[0] if isinstance(_p, list) and _p else _p or {})
        except Exception:
            pass

        years = [
            d.get("fiscalYear") or d.get("calendarYear") or d["date"][:4]
            for d in is_data
        ]

        analyst_ests = []
        try:
            _ae = _req.get(
                f"https://financialmodelingprep.com/stable/analyst-estimates"
                f"?symbol={ticker}&period=annual&limit=4&apikey={mdl.API_KEY}", timeout=8
            ).json()
            if isinstance(_ae, list):
                analyst_ests = sorted(
                    [e for e in _ae if str(e.get("date",""))[:4] > str(years[-1])],
                    key=lambda x: x.get("date","")
                )[:2]
        except Exception:
            pass

        save_ticker_data(ticker, is_data, bs_data, cf_data, profile, years,
                         None, {}, {}, analyst_ests)
        price = float(profile.get("price") or 0) or None
        print(f"  {ticker:6s}  OK  price=${price:.2f}" if price else f"  {ticker:6s}  OK")
        return True

    except Exception as e:
        msg = str(e)
        if "429" in msg or "402" in msg or "limit" in msg.lower():
            print(f"  {ticker:6s}  QUOTA — {e}")
            return "quota"
        print(f"  {ticker:6s}  FAIL — {e}")
        return False


def main():
    # Determine which tickers to seed
    if len(sys.argv) > 1:
        targets = [t.upper() for t in sys.argv[1:]]
        force = True  # explicit list = always refresh
    else:
        targets = CORE_TICKERS + _discovered_tickers()
        force = False

    print(f"\nSeeding data store for {len(targets)} ticker(s)")
    print(f"Stored in: static/data/")
    print("=" * 50)

    ok = 0; skipped = 0; failed = 0
    for ticker in targets:
        result = seed_ticker(ticker, force=force)
        if result is True:
            # distinguish cached vs freshly fetched
            ok += 1
        elif result == "quota":
            print(f"\n  FMP quota exhausted — stopping. Run again when quota resets.")
            break
        else:
            failed += 1
        time.sleep(3)  # polite gap between tickers

    print("=" * 50)
    print(f"Done: {ok} seeded, {failed} failed")
    print(f"\nNext step: git add static/data/ && git commit -m 'seed data store' && git push")


if __name__ == "__main__":
    main()
