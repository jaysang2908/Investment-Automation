"""
Runs NVDA, NFLX, COST through the updated engine (SCORE_CAP=10, NA-rescale note).
Regenerates HTML reports and prints detailed scorecard data for each ticker.
"""

import io, os, sys, time, traceback, builtins
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fmp_3statementv6 as mdl
from report_bridge import build_report_data, render_html_report
from data_store import save_ticker_data
import requests as _req
from openpyxl import Workbook

TICKERS = [
    ("COST", None, None),
    ("NVDA", None, None),
    ("NFLX", None, None),
]

TIER_PTS = {"HIGH": 10, "MOD-HIGH": 7, "MOD": 7, "MOD-LOW": 3, "LOW": 0}
OUT_DIR  = os.path.join(os.path.dirname(__file__), "static", "reports")

CRITERIA = [
    ("Moat Profile",   "tier_moat",     "score_moat",     10.0),
    ("Management",     "tier_mgmt",     "score_mgmt",      7.5),
    ("Rev 3yr CAGR",   "tier_rev_cagr", "score_rev_cagr", 10.0),
    ("FCF Quality",    "tier_fcf_ni",   "score_fcf_ni",   10.0),
    ("Cap Returns",    "tier_cap_ret",  "score_cap_ret",   5.0),
    ("ROIC",           "tier_roic",     "score_roic",      7.5),
    ("Leverage",       "tier_leverage", "score_leverage",  5.0),
    ("Interest Cover", "tier_ebit_int", "score_ebit_int",  7.5),
    ("Exec Risk",      "tier_exec",     "score_exec",      5.0),
    ("P/E",            "tier_pe",       "score_pe",       10.0),
    ("P/FCF",          "tier_pfcf",     "score_pfcf",     10.0),
]

all_results = {}

for ticker, biz_clarity, ltp in TICKERS:
    print(f"\n{'='*60}\n  {ticker}\n{'='*60}")

    logs = []
    _orig = builtins.print
    builtins.print = lambda *a, **k: logs.append(" ".join(str(x) for x in a))

    try:
        is_data = mdl.fetch("income-statement",        ticker)[:mdl.YEARS][::-1]
        bs_data = mdl.fetch("balance-sheet-statement", ticker)[:mdl.YEARS][::-1]
        cf_data = mdl.fetch("cash-flow-statement",     ticker)[:mdl.YEARS][::-1]

        if not is_data:
            raise ValueError("No income statement data")

        profile = {}
        current_price = None
        market_cap    = None
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

        wb       = Workbook()
        pl_refs  = mdl.build_pl(wb, is_data, years, ticker)
        mdl.build_cover(wb, ticker, years, is_data)
        bs_refs  = mdl.build_bs(wb, bs_data, years, ticker)
        cf_refs  = mdl.build_cf(wb, cf_data, years, ticker)
        mdl.build_ratios(wb, is_data, bs_data, cf_data, years, ticker, pl_refs, bs_refs, cf_refs)
        mdl.build_segments(wb, ticker, years)
        wacc_refs = mdl.build_wacc(wb, ticker, is_data, bs_data, None)
        dcf_refs  = mdl.build_dcf(
            wb, ticker, is_data, bs_data, cf_data, years,
            pl_refs, bs_refs, wacc_refs, current_price=current_price, cf_refs=cf_refs
        )
        _, sm = mdl.build_scorecard(
            wb, ticker, is_data, bs_data, cf_data, years,
            dcf_gg_price=(dcf_refs.get("dcf_prices") or {}).get("gg_price"),
            evs_regime=bool((dcf_refs.get("dcf_prices") or {}).get("evs_regime")),
        )

        W       = sm.get("weights") or {}
        _w_bc   = float(W.get("BC",  2.5))
        _w_ltp  = float(W.get("LTP", 10.0))
        bc_pts  = TIER_PTS.get(biz_clarity, 0) * _w_bc  / 10
        ltp_pts = TIER_PTS.get(ltp,         0) * _w_ltp / 10
        auto_score = sm.get("auto_score") or 0
        adj_score  = round(auto_score + bc_pts + ltp_pts, 1)
        floor_cap  = sm.get("floor_cap")
        if floor_cap is not None:
            adj_score = min(adj_score, floor_cap)

        report_data  = build_report_data(
            ticker=ticker, profile=profile,
            is_data=is_data, bs_data=bs_data, cf_data=cf_data, years=years,
            wacc_val=wacc_refs.get("wacc_val"),
            dcf_prices=(dcf_refs or {}).get("dcf_prices") or {},
            scorecard_metrics=sm, manual_rating=None,
            current_price=current_price, market_cap=market_cap,
            biz_clarity=biz_clarity or None, ltp=ltp or None,
            adj_score=adj_score, analyst_ests=analyst_ests,
        )
        html_content = render_html_report(report_data)

        out_path = os.path.join(OUT_DIR, f"{ticker}_report.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        save_ticker_data(
            ticker, is_data, bs_data, cf_data, profile, years,
            wacc_refs.get("wacc_val"),
            (dcf_refs or {}).get("dcf_prices") or {},
            sm, analyst_ests
        )

        builtins.print = _orig

        # ── Scorecard detail ──────────────────────────────────────────────
        scored_w = sm.get("scored_weight", 0.0)
        active_w = sm.get("active_weight", 87.5)
        low_conf = sm.get("low_data_confidence", False)
        rescale  = (scored_w > 0 and scored_w < active_w and not low_conf)
        rf       = round(active_w / scored_w, 3) if rescale and scored_w > 0 else 1.0

        rows = []
        raw_sum = 0.0
        for name, tier_key, score_key, weight in CRITERIA:
            tier  = sm.get(tier_key)
            score = sm.get(score_key)
            s_cap = min(float(score), 10.0) if score is not None else 0.0
            wtd   = round(s_cap / 10.0 * weight, 2) if tier is not None else 0.0
            rows.append((name, tier or "N/A", round(s_cap, 2), weight, wtd))
            if tier is not None and score is not None:
                raw_sum += s_cap / 10.0 * weight

        if rescale:
            raw_sum_rescaled = round(raw_sum * rf, 2)
        else:
            raw_sum_rescaled = round(raw_sum, 2)

        all_results[ticker] = {
            "rows":          rows,
            "auto_score":    auto_score,
            "adj_score":     adj_score,
            "floor_cap":     floor_cap,
            "sector":        sm.get("sector_bucket"),
            "scored_w":      scored_w,
            "active_w":      active_w,
            "rescale":       rescale,
            "rf":            rf,
            "raw_sum":       round(raw_sum, 2),
            "raw_sum_rsc":   raw_sum_rescaled,
            "roic":          sm.get("roic"),
            "rev_cagr":      sm.get("rev_cagr"),
            "fcf_ni":        sm.get("fcf_ni"),
            "d_ebitda":      sm.get("d_ebitda"),
            "price":         current_price,
        }

        print(f"  OK {ticker}  auto={auto_score}  adj={adj_score}  sector={sm.get('sector_bucket')}")
        print(f"     scored_w={scored_w}  active_w={active_w}  rescale={rescale}  rf={rf}")
        for r in rows:
            print(f"     {r[0]:20s}  {str(r[1]):9s}  score={r[2]:5.2f}  wt={r[3]:4.1f}  wtd={r[4]:.2f}")

    except Exception as e:
        builtins.print = _orig
        print(f"  FAIL {ticker}: {e}")
        traceback.print_exc()
        all_results[ticker] = {"error": str(e)}

    time.sleep(8)


# ── Final summary table ───────────────────────────────────────────────────────
print("\n\n" + "="*80)
print("SCORECARD TABLE")
print("="*80)
for t in ["COST", "NVDA", "NFLX"]:
    r = all_results.get(t, {})
    if r.get("error"):
        print(f"\n{t}: ERROR — {r['error']}")
        continue
    print(f"\n{t}  sector={r['sector']}  price=${r.get('price') or 'N/A'}")
    print(f"  ROIC={r['roic']:.1%}  RevCAGR={r['rev_cagr']:.1%}  FCF/NI={r['fcf_ni']:.1%}  D/EBITDA={r.get('d_ebitda') or 'N/A'}")
    print(f"  Rescale: {r['rescale']}  factor={r['rf']}  raw_sum={r['raw_sum']}  raw_rescaled={r['raw_sum_rsc']}")
    print(f"  Auto Score: {r['auto_score']} / 10   Adj Score: {r['adj_score']} / 10")
    print(f"  {'Criterion':<22} {'Tier':<10} {'Score':>6} {'Wt':>5} {'WTD':>6}")
    print(f"  {'-'*54}")
    for name, tier, score, weight, wtd in r["rows"]:
        print(f"  {name:<22} {tier:<10} {score:>6.2f} {weight:>5.1f} {wtd:>6.2f}")
