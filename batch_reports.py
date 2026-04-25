"""
batch_reports.py
Generates HTML reports for all 24 dashboard tickers using the updated engine.
Saves to static/reports/TICKER_report.html and prints score summary.
"""

import io, os, sys, time, traceback
import builtins

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fmp_3statementv6 as mdl
from report_bridge import build_report_data, render_html_report
import requests as _req
from openpyxl import Workbook

OUT_DIR = os.path.join(os.path.dirname(__file__), "static", "reports")
os.makedirs(OUT_DIR, exist_ok=True)

TIER_PTS = {"HIGH": 10, "MOD-HIGH": 7, "MOD-LOW": 3, "LOW": 0}

# Tickers + optional qualitative inputs (biz_clarity, ltp)
# HIGH / MOD-HIGH / MOD-LOW / LOW  — or None to leave blank
TICKERS = [
    ("NVDA", None, None),
    ("MSFT", None, None),
    ("AAPL", None, None),
    ("ADBE", None, None),
    ("COST", None, None),
    ("AMD",  None, None),
    ("JNJ",  None, None),
    ("META", None, None),
    ("TSM",  None, None),
    ("V",    None, None),
    ("KO",   None, None),
    ("NFLX", None, None),
    ("ABBV", None, None),
    ("CSCO", None, None),
    ("WMT",  None, None),
    ("F",    None, None),
    ("WFC",  None, None),
    ("INTC", None, None),
    ("TSLA", None, None),
    ("SOFI", None, None),
    ("JPM",  None, None),
    ("C",    None, None),
    ("BAC",  None, None),
    ("UAL",  None, None),
]

results = []

for ticker, biz_clarity, ltp in TICKERS:
    print(f"\n{'='*60}")
    print(f"  {ticker}")
    print(f"{'='*60}")

    # Suppress sub-prints from model
    logs = []
    _orig = builtins.print
    builtins.print = lambda *a, **k: logs.append(" ".join(str(x) for x in a))

    try:
        is_data = mdl.fetch("income-statement",        ticker)[:mdl.YEARS][::-1]
        bs_data = mdl.fetch("balance-sheet-statement", ticker)[:mdl.YEARS][::-1]
        cf_data = mdl.fetch("cash-flow-statement",     ticker)[:mdl.YEARS][::-1]

        if not is_data:
            raise ValueError("No income statement data returned")

        profile = {}
        current_price = None
        market_cap = None
        try:
            _p = _req.get(
                f"https://financialmodelingprep.com/stable/profile"
                f"?symbol={ticker}&apikey={mdl.API_KEY}", timeout=8
            ).json()
            profile       = (_p[0] if isinstance(_p, list) and _p else _p or {})
            current_price = float(profile.get("price") or 0) or None
            market_cap    = float(profile.get("mktCap") or profile.get("marketCap") or 0) or None
        except Exception:
            pass

        years = [
            d.get("fiscalYear") or d.get("calendarYear") or d["date"][:4]
            for d in is_data
        ]

        # Analyst estimates for forward multiples (FY+1, FY+2)
        analyst_ests = []
        try:
            _ae = _req.get(
                f"https://financialmodelingprep.com/stable/analyst-estimates"
                f"?symbol={ticker}&period=annual&limit=4&apikey={mdl.API_KEY}", timeout=8
            ).json()
            if isinstance(_ae, list):
                # Keep only estimates dated after last historical year
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
        _, scorecard_metrics = mdl.build_scorecard(wb, ticker, is_data, bs_data, cf_data, years)

        # Compute adj_score
        auto_score = scorecard_metrics.get("auto_score") or 0
        bc_pts  = TIER_PTS.get(biz_clarity, 0) * 2.5 / 10
        ltp_pts = TIER_PTS.get(ltp, 0) * 10.0 / 10
        adj_score = round(auto_score + bc_pts + ltp_pts, 1)
        floor_cap = scorecard_metrics.get("floor_cap")
        if floor_cap is not None:
            adj_score = min(adj_score, floor_cap)

        report_data = build_report_data(
            ticker            = ticker,
            profile           = profile,
            is_data           = is_data,
            bs_data           = bs_data,
            cf_data           = cf_data,
            years             = years,
            wacc_val          = wacc_refs.get("wacc_val"),
            dcf_prices        = (dcf_refs or {}).get("dcf_prices") or {},
            scorecard_metrics = scorecard_metrics,
            manual_rating     = None,
            current_price     = current_price,
            market_cap        = market_cap,
            biz_clarity       = biz_clarity or None,
            ltp               = ltp or None,
            adj_score         = adj_score,
            analyst_ests      = analyst_ests,
        )
        html_content = render_html_report(report_data)

        out_path = os.path.join(OUT_DIR, f"{ticker}_report.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        dcf_prices = (dcf_refs or {}).get("dcf_prices") or {}
        gg_px  = dcf_prices.get("gg_price")
        em_px  = dcf_prices.get("em_price")
        gg_up  = round((gg_px - current_price) / current_price, 3) if gg_px and current_price else None
        em_up  = round((em_px - current_price) / current_price, 3) if em_px and current_price else None

        results.append({
            "ticker":      ticker,
            "adj_score":   adj_score,
            "auto_score":  auto_score,
            "floor_cap":   floor_cap,
            "is_bank":     scorecard_metrics.get("is_bank", False),
            "roic":        scorecard_metrics.get("roic"),
            "rev_cagr":    scorecard_metrics.get("rev_cagr"),
            "fcf_ni":      scorecard_metrics.get("fcf_ni"),
            "d_ebitda":    scorecard_metrics.get("d_ebitda"),
            "equity_assets": scorecard_metrics.get("equity_assets"),
            "pe_current":  scorecard_metrics.get("pe_current"),
            "pe_5yr":      scorecard_metrics.get("pe_5yr_avg"),
            "pfcf_current":scorecard_metrics.get("pfcf_current"),
            "pfcf_5yr":    scorecard_metrics.get("pfcf_5yr_avg"),
            "gg_upside":   gg_up,
            "em_upside":   em_up,
            "current_price": current_price,
            "saved":       out_path,
        })

        builtins.print = _orig
        print(f"  OK {ticker}  adj_score={adj_score}  auto={auto_score}"
              f"  floor={floor_cap}  is_bank={scorecard_metrics.get('is_bank',False)}")

    except Exception as e:
        builtins.print = _orig
        print(f"  FAIL {ticker}: {e}")
        traceback.print_exc()
        results.append({"ticker": ticker, "adj_score": None, "error": str(e)})

    time.sleep(8)   # avoid rate-limiting (5 API calls per ticker)


# ── Summary ───────────────────────────────────────────────────────────────────
print("\n\n" + "="*70)
print("SCORE SUMMARY")
print("="*70)
for r in sorted(results, key=lambda x: x.get("adj_score") or -1, reverse=True):
    if r.get("error"):
        print(f"  {r['ticker']:6s}  ERROR: {r['error']}")
        continue
    bank = " [BANK]" if r.get("is_bank") else ""
    floor = f"  cap={r['floor_cap']}" if r.get("floor_cap") else ""
    print(f"  {r['ticker']:6s}  score={r['adj_score']:<6}{floor}{bank}"
          f"  ROIC={r['roic']:.1%}" if r.get('roic') else
          f"  {r['ticker']:6s}  score={r.get('adj_score')}")

print("\nAll reports saved to static/reports/")
print("\nCopy the results above to update dashboard.html data array.")
