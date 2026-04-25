"""
daily_discovery.py
Generates up to 5 new research reports per day from a discovery pool.
Handles FMP 402 quota errors — retries with the next random ticker.
On success: writes to outputs.csv + commits + pushes to GitHub/Render.

Usage:  python daily_discovery.py
Cron:   scheduled daily at 7:17 AM via Claude Code.
"""

import os, sys, random, time, subprocess, datetime, builtins
import requests as _req
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fmp_3statementv6 as mdl
import csv_schema as _schema
from report_bridge import build_report_data, render_html_report
from data_store import save_ticker_data

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR  = os.path.join(BASE_DIR, "static", "reports")
CSV_PATH = os.path.join(BASE_DIR, "outputs.csv")
LOG_FILE = os.path.join(BASE_DIR, "discovery_log.txt")
os.makedirs(OUT_DIR, exist_ok=True)

MAX_DAILY     = 5    # reports to generate per run
QUOTA_STRIKES = 3    # consecutive 402s before giving up

# ── Tickers already in core coverage ─────────────────────────────────────────
EXISTING = {
    "NVDA","MSFT","AAPL","ADBE","COST","AMD","JNJ","META","TSM","V",
    "KO","NFLX","ABBV","CSCO","WMT","F","WFC","INTC","TSLA","SOFI",
    "JPM","C","BAC","UAL",
}

# ── Discovery pool — high-quality liquid US equities ─────────────────────────
POOL = [
    # Mega-cap tech / software
    "GOOGL","AMZN","ORCL","CRM","NOW","SNOW","PANW","CRWD","INTU",
    "AMAT","KLAC","LRCX","MU","AVGO","QCOM","TXN","HPQ","DELL",
    "DDOG","TEAM","ZM","OKTA","NET","MDB","GTLB",
    # Financials
    "GS","MS","BLK","SCHW","AXP","COF","USB","PNC","TFC","MET","PRU",
    "ICE","CME","SPGI","MCO","MSCI","FDS",
    # Healthcare / Biotech
    "UNH","LLY","BMY","MRK","AMGN","GILD","REGN","VRTX","TMO","DHR",
    "MDT","SYK","BSX","ABT","ISRG","HCA","CI","CVS","ZTS","DXCM",
    # Consumer / Retail / Travel
    "PG","PEP","MCD","SBUX","NKE","TGT","HD","LOW","TJX","BKNG",
    "MAR","HLT","CMG","YUM","DASH","ABNB","RCL","CCL","LVS",
    # Industrials / Aerospace / Energy
    "CAT","DE","HON","UPS","FDX","RTX","LMT","BA","GE","EMR","ETN",
    "XOM","CVX","COP","SLB","OXY","NEE","DUK","SO","AEP",
    # Communication / Media
    "DIS","CMCSA","T","VZ","CHTR","SPOT","TTD","PINS","SNAP",
    # REITs
    "PLD","AMT","EQIX","CCI","SPG","O","WELL","DLR","EXR",
    # International ADRs
    "BABA","JD","SE","MELI","SHOP","ASML","SAP","TM","SONY","NVO","NOVO-B",
]

POOL = list(dict.fromkeys([t for t in POOL if t not in EXISTING]))


# ── Quota check ───────────────────────────────────────────────────────────────
def _quota_ok():
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
        return True


# ── Report generator ──────────────────────────────────────────────────────────
def try_generate(ticker):
    """
    Attempt full report generation.
    Returns (True, data_dict) on success.
    Returns (False, "quota")  on 402 / rate-limit.
    Returns (False, reason)   on other failures.
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

        # Save raw data so DCF calculator can load without re-hitting FMP
        dcf_p = (dcf_refs or {}).get("dcf_prices") or {}
        save_ticker_data(
            ticker, is_data, bs_data, cf_data, profile, years,
            wacc_refs.get("wacc_val"), dcf_p, scorecard_metrics, analyst_ests
        )

        return True, {
            "score":      score,
            "metrics":    scorecard_metrics,
            "dcf_prices": dcf_p,
            "price":      current_price,
            "mkt_cap":    market_cap,
            "revenue_b":  (is_data[-1].get("revenue") or 0) / 1e9,
            "ocf_b":      (cf_data[-1].get("operatingCashFlow") or 0) / 1e9,
            "fcf_b":      (cf_data[-1].get("freeCashFlow") or
                           (cf_data[-1].get("operatingCashFlow") or 0) -
                           abs(cf_data[-1].get("capitalExpenditure") or 0)) / 1e9,
        }

    except Exception as e:
        msg = str(e).lower()
        is_quota = any(x in msg for x in ["402", "limit reach", "subscription", "upgrade your plan"])
        return False, "quota" if is_quota else f"error: {e}"
    finally:
        builtins.print = _orig


# ── CSV write ─────────────────────────────────────────────────────────────────
def _write_csv_rows(rows):
    """Append/update rows in outputs.csv. rows = list of (ticker, data_dict)."""
    if os.path.exists(CSV_PATH):
        with open(CSV_PATH, "r", encoding="utf-8") as f:
            content = f.read()
    else:
        content = _schema.HEADER

    content = _schema.migrate(content)

    def _f(v, dp=4):
        return "" if v is None else f"{v:.{dp}f}"

    for ticker, data in rows:
        metrics = data.get("metrics", {})
        dp_     = data.get("dcf_prices", {})
        price   = data.get("price")
        mkt_b   = (data.get("mkt_cap") or 0) / 1e9 or None
        gg_px   = dp_.get("gg_price")
        em_px   = dp_.get("em_price")
        gg_up   = round((gg_px - price) / price, 4) if gg_px and price else None
        em_up   = round((em_px - price) / price, 4) if em_px and price else None

        row = {
            "Ticker":         ticker,
            "Price":          _f(price, 2),
            "MktCap_B":       _f(mkt_b, 2),
            "GG_Price":       _f(gg_px, 2),
            "GG_Upside":      _f(gg_up, 4),
            "EM_Price":       _f(em_px, 2),
            "EM_Upside":      _f(em_up, 4),
            "PE_Current":     _f(metrics.get("pe_current"), 1),
            "PE_5yr":         _f(metrics.get("pe_5yr_avg"), 1),
            "PFCF_Current":   _f(metrics.get("pfcf_current"), 1),
            "PFCF_5yr":       _f(metrics.get("pfcf_5yr_avg"), 1),
            "ROIC":           _f(metrics.get("roic")),
            "Rev_CAGR":       _f(metrics.get("rev_cagr")),
            "FCF_NI":         _f(metrics.get("fcf_ni")),
            "D_EBITDA":       _f(metrics.get("d_ebitda"), 2),
            "Revenue_B":      _f(data.get("revenue_b"), 2),
            "OCF_B":          _f(data.get("ocf_b"), 2),
            "FCF_B":          _f(data.get("fcf_b"), 2),
            "Auto_Score":     str(metrics.get("auto_score") or ""),
            "Floor_Cap":      str(metrics.get("floor_cap") or ""),
            "Manual_Clarity": "",
            "Manual_LTP":     "",
            "Date":           datetime.date.today().isoformat(),
        }
        content += ",".join(row.get(c, "") for c in _schema.COLUMNS) + "\n"

    with open(CSV_PATH, "w", encoding="utf-8") as f:
        f.write(content)


# ── Git commit + push (all files at once) ─────────────────────────────────────
def _git_push_all(html_tickers, today):
    try:
        files = [f"static/reports/{t}_report.html" for t in html_tickers] + ["outputs.csv"]
        subprocess.run(["git", "add"] + files, cwd=BASE_DIR, check=True, capture_output=True)
        names = ", ".join(html_tickers)
        subprocess.run(
            ["git", "commit", "-m",
             f"Daily discovery ({today}): {names}\n\n"
             f"Auto-generated {len(html_tickers)} report(s) via daily_discovery.py"],
            cwd=BASE_DIR, check=True, capture_output=True
        )
        subprocess.run(["git", "push"], cwd=BASE_DIR, check=True, capture_output=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"  Git push failed: {e.stderr.decode() if e.stderr else e}")
        return False


# ── Log ───────────────────────────────────────────────────────────────────────
def _log(msg):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    today = datetime.date.today().strftime("%Y-%m-%d")
    print(f"\n{'='*65}")
    print(f"  Daily Discovery — {today}  (target: {MAX_DAILY} reports)")
    print(f"{'='*65}\n")

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

    # Reproducible daily shuffle (same seed = same order = idempotent retries)
    rng = random.Random(today)
    rng.shuffle(candidates)

    successes    = []   # list of (ticker, data_dict)
    quota_count  = 0
    tried        = []

    for ticker in candidates:
        if len(successes) >= MAX_DAILY:
            break
        if quota_count >= QUOTA_STRIKES:
            print(f"\n  {QUOTA_STRIKES} consecutive quota errors — FMP limit hit for today.")
            break

        n = len(successes) + 1
        print(f"  [{n}/{MAX_DAILY}] {ticker}...", end=" ", flush=True)
        tried.append(ticker)
        success, result = try_generate(ticker)

        if success:
            print(f"OK  score={result['score']}")
            successes.append((ticker, result))
            quota_count = 0
            if len(successes) < MAX_DAILY:
                time.sleep(3)   # polite gap between API calls

        elif result == "quota":
            print("WARN  402 rate limit")
            quota_count += 1
            time.sleep(5)
        else:
            print(f"FAIL  {result}")
            quota_count = 0
            time.sleep(2)

    # ── Persist all results ───────────────────────────────────────────────────
    if successes:
        tickers_done = [t for t, _ in successes]

        _write_csv_rows(successes)

        pushed = _git_push_all(tickers_done, today)
        status = "pushed to GitHub -> live on Render" if pushed else "saved locally (push manually)"

        scores_str = "  ".join(f"{t}={d['score']}" for t, d in successes)
        msg = (f"{today} | {len(successes)}/{MAX_DAILY} reports: {', '.join(tickers_done)} | "
               f"scores: {scores_str} | tried: {tried} | {status}")
        print(f"\n  OK {len(successes)} report(s): {', '.join(tickers_done)}")
        print(f"  -> {status}")
        _log(msg)
    else:
        msg = f"{today} | 0 reports generated. Tried: {tried}"
        print(f"\n  {msg}")
        _log(msg)


if __name__ == "__main__":
    main()
