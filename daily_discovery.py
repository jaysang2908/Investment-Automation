"""
daily_discovery.py
Generates one new research report per day from a discovery pool.
Handles FMP 402 quota errors gracefully — tries the next random ticker
until one succeeds, then commits + pushes to GitHub/Render.

Usage:  python daily_discovery.py
Cron:   scheduled daily at 7:17 AM via Claude Code.
"""

import os, sys, random, time, subprocess, datetime, builtins
import requests as _req
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fmp_3statementv6 as mdl
from report_bridge import build_report_data, render_html_report

OUT_DIR  = os.path.join(os.path.dirname(__file__), "static", "reports")
LOG_FILE = os.path.join(os.path.dirname(__file__), "discovery_log.txt")
os.makedirs(OUT_DIR, exist_ok=True)

# ── Tickers already covered — never regenerate these ─────────────────────────
EXISTING = {
    "NVDA","MSFT","AAPL","ADBE","COST","AMD","JNJ","META","TSM","V",
    "KO","NFLX","ABBV","CSCO","WMT","F","WFC","INTC","TSLA","SOFI",
    "JPM","C","BAC","UAL",
}

# ── Discovery pool — high-quality liquid US equities (expand as needed) ───────
POOL = [
    # Mega-cap tech / software
    "GOOGL","AMZN","ORCL","CRM","NOW","SNOW","PANW","CRWD","INTU",
    "AMAT","KLAC","LRCX","MU","AVGO","QCOM","TXN","HPQ","DELL","SMCI",
    "ZM","DOCU","OKTA","DDOG","TEAM","ATLASSIAN",
    # Financials
    "GS","MS","BLK","SCHW","AXP","COF","USB","PNC","TFC","MET","PRU",
    "ICE","CME","SPGI","MCO","MSCI",
    # Healthcare / Biotech
    "UNH","LLY","BMY","MRK","AMGN","GILD","REGN","VRTX","TMO","DHR",
    "MDT","SYK","BSX","ABT","ISRG","HCA","CI","CVS","ZTS",
    # Consumer / Retail / Travel
    "PG","PEP","MCD","SBUX","NKE","TGT","HD","LOW","TJX","BKNG",
    "MAR","HLT","CMG","YUM","DASH","ABNB","RCL","CCL",
    # Industrials / Aerospace / Energy
    "CAT","DE","HON","UPS","FDX","RTX","LMT","BA","GE","EMR","ETN",
    "XOM","CVX","COP","SLB","OXY","NEE","DUK","SO",
    # Communication / Media
    "DIS","CMCSA","NFLX","T","VZ","CHTR","WBD","PARA","SPOT","TTD",
    # REITs
    "PLD","AMT","EQIX","CCI","SPG","O","WELL","DLR",
    # International ADRs / Global
    "BABA","JD","PDD","SE","GRAB","NU","MELI","SHOP","ASML","SAP","TM","SONY",
]

# Remove existing + deduplicate
POOL = list(dict.fromkeys([t for t in POOL if t not in EXISTING]))


# ── Quota check ───────────────────────────────────────────────────────────────
def _quota_ok():
    """Quick check: returns False if FMP quota is clearly exhausted."""
    try:
        r = _req.get(
            f"https://financialmodelingprep.com/stable/profile"
            f"?symbol=AAPL&apikey={mdl.API_KEY}", timeout=6
        )
        if r.status_code == 402:
            return False
        body = r.json()
        if isinstance(body, dict) and ("Error" in body or "message" in str(body).lower()):
            return False
        return True
    except Exception:
        return True   # unclear — let it try


# ── Report generator ──────────────────────────────────────────────────────────
def try_generate(ticker):
    """
    Attempt full report generation for ticker.
    Returns (True, score)  on success.
    Returns (False, "quota") on 402 / rate-limit.
    Returns (False, reason) on any other failure.
    """
    logs = []
    _orig = builtins.print
    builtins.print = lambda *a, **k: logs.append(" ".join(str(x) for x in a))
    try:
        is_data = mdl.fetch("income-statement",        ticker)[:mdl.YEARS][::-1]
        bs_data = mdl.fetch("balance-sheet-statement", ticker)[:mdl.YEARS][::-1]
        cf_data = mdl.fetch("cash-flow-statement",     ticker)[:mdl.YEARS][::-1]

        if not is_data:
            return False, "no data"

        profile = {}; current_price = None; market_cap = None
        try:
            _p = _req.get(
                f"https://financialmodelingprep.com/stable/profile"
                f"?symbol={ticker}&apikey={mdl.API_KEY}", timeout=8
            )
            if _p.status_code == 402:
                return False, "quota"
            pb = _p.json()
            profile       = (pb[0] if isinstance(pb, list) and pb else pb or {})
            current_price = float(profile.get("price") or 0) or None
            market_cap    = float(profile.get("mktCap") or 0) or None
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
                    [e for e in _ae if e.get("date", "")[:4] > str(years[-1])],
                    key=lambda x: x.get("date", "")
                )[:2]
        except Exception:
            pass

        wb = Workbook()
        pl_refs  = mdl.build_pl(wb, is_data, years, ticker)
        mdl.build_cover(wb, ticker, years, is_data)
        bs_refs  = mdl.build_bs(wb, bs_data, years, ticker)
        cf_refs  = mdl.build_cf(wb, cf_data, years, ticker)
        mdl.build_ratios(wb, is_data, bs_data, cf_data, years, ticker, pl_refs, bs_refs, cf_refs)
        mdl.build_segments(wb, ticker, years)
        wacc_refs = mdl.build_wacc(wb, ticker, is_data, bs_data, None)
        dcf_refs  = mdl.build_dcf(
            wb, ticker, is_data, bs_data, cf_data, years,
            pl_refs, bs_refs, wacc_refs, current_price=current_price
        )
        _, scorecard_metrics = mdl.build_scorecard(
            wb, ticker, is_data, bs_data, cf_data, years
        )

        auto  = scorecard_metrics.get("auto_score") or 0
        floor = scorecard_metrics.get("floor_cap")
        score = round(min(auto, floor) if floor else auto, 1)

        report_data = build_report_data(
            ticker=ticker, profile=profile,
            is_data=is_data, bs_data=bs_data, cf_data=cf_data, years=years,
            wacc_val=wacc_refs.get("wacc_val"),
            dcf_prices=(dcf_refs or {}).get("dcf_prices") or {},
            scorecard_metrics=scorecard_metrics,
            current_price=current_price, market_cap=market_cap,
            adj_score=score, analyst_ests=analyst_ests,
        )
        html_content = render_html_report(report_data)

        out_path = os.path.join(OUT_DIR, f"{ticker}_report.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        return True, score

    except Exception as e:
        msg = str(e).lower()
        is_quota = any(x in msg for x in ["402", "limit reach", "subscription", "upgrade your plan"])
        return False, "quota" if is_quota else f"error: {e}"
    finally:
        builtins.print = _orig


# ── Git push ──────────────────────────────────────────────────────────────────
def _git_push(ticker, today):
    base = os.path.dirname(os.path.abspath(__file__))
    try:
        subprocess.run(
            ["git", "add", f"static/reports/{ticker}_report.html"],
            cwd=base, check=True, capture_output=True
        )
        subprocess.run(
            ["git", "commit", "-m", f"Daily discovery: {ticker} — {today}"],
            cwd=base, check=True, capture_output=True
        )
        subprocess.run(["git", "push"], cwd=base, check=True, capture_output=True)
        return True
    except subprocess.CalledProcessError as e:
        stderr = e.stderr.decode() if e.stderr else str(e)
        print(f"  Git push failed: {stderr}")
        return False


# ── Log ───────────────────────────────────────────────────────────────────────
def _log(msg):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    today = datetime.date.today().strftime("%Y-%m-%d")
    print(f"\n{'='*60}")
    print(f"  Daily Discovery — {today}")
    print(f"{'='*60}\n")

    # Skip tickers already reported on
    already_done = {
        f.replace("_report.html", "")
        for f in os.listdir(OUT_DIR) if f.endswith("_report.html")
    }
    candidates = [t for t in POOL if t not in already_done]

    if not candidates:
        msg = f"{today} | Pool exhausted — all discovery tickers already have reports."
        print(f"  {msg}")
        _log(msg)
        return

    if not _quota_ok():
        msg = f"{today} | FMP quota exhausted — skipped."
        print(f"  {msg}")
        _log(msg)
        return

    # Reproducible daily shuffle (same order for a given date → no duplicates on retry)
    rng = random.Random(today)
    rng.shuffle(candidates)

    quota_strikes = 0
    tried = []

    for ticker in candidates:
        if quota_strikes >= 3:
            msg = f"{today} | 3 consecutive quota errors — FMP limit hit. Tried: {tried}"
            print(f"\n  {msg}")
            _log(msg)
            break

        print(f"  Trying {ticker}...", end=" ", flush=True)
        tried.append(ticker)
        success, result = try_generate(ticker)

        if success:
            print(f"✓  score={result}")
            pushed = _git_push(ticker, today)
            status = "pushed" if pushed else "saved locally"
            msg = f"{today} | SUCCESS {ticker}  score={result}  {status}  (tried: {tried})"
            print(f"\n  ✓ {ticker}_report.html generated and {status}.")
            _log(msg)
            return

        if result == "quota":
            print("⚠  402 rate limit")
            quota_strikes += 1
            time.sleep(4)
        else:
            print(f"✗  {result}")
            quota_strikes = 0
            time.sleep(2)

    msg = f"{today} | No report generated. Tried: {tried}"
    print(f"\n  {msg}")
    _log(msg)


if __name__ == "__main__":
    main()
