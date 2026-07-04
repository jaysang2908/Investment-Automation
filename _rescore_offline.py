"""
_rescore_offline.py
Re-runs the full engine (WACC + DCF + Scorecard + Report) for a list of tickers
using ONLY data from static/data/{ticker}_data.json — zero FMP API calls.

How it avoids FMP calls:
  1. Pre-populates _RATIOS_CACHE (in fmp_3statementv6) from stored scorecard_metrics
     so _fetch_ratios() returns cached data instead of hitting the API.
  2. Passes stored `profile` dict to build_wacc() — skips the /stable/profile call.
  3. Passes `manual_rating` from credit_ratings.json — skips the /stable/ratings call.
  4. Passes stored `analyst_ests` to build_scorecard() — no analyst estimates call.
  5. Peer beta + peer P/E calls will 429 → silent fallback (Damodaran industry beta;
     sector_pe_med = None → benchmark falls back to own hist_avg, which is correct).
  6. _fetch_ttm_ni() will 429 → trailing_pe falls back to FMP snapshot from cache.
     When analyst_ests are present, pe_current = forward_pe_val (no FMP needed).

Limitations:
  - Current price is stale (from last fetch, stored in profile["price"]).
  - Peer sector P/E median is None (correct for new benchmark logic, but shown as N/A
    in the note).
  - WACC may differ slightly from a fresh run (peer beta component missing → pure
    Damodaran + raw_beta blend). Credit-rating Kd IS applied via manual_rating.

Usage:
    python _rescore_offline.py                  # all pending tickers
    python _rescore_offline.py CSCO META NFLX   # specific tickers
"""
import builtins, csv, datetime, io, os, sys, traceback, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook
import fmp_3statementv6 as mdl
from report_bridge import build_report_data, render_html_report
from data_store import save_ticker_data, load_ticker_data
import csv_schema as _schema

_ROOT    = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(_ROOT, "outputs.csv")
RPT_DIR  = os.path.join(_ROOT, "static", "reports")
XLS_DIR  = os.path.join(_ROOT, "static", "excel")
os.makedirs(RPT_DIR, exist_ok=True)
os.makedirs(XLS_DIR, exist_ok=True)

TIER_PTS = {"HIGH": 10, "MOD-HIGH": 7, "MOD": 7, "MOD-LOW": 3, "LOW": 0}

# Load credit ratings for WACC auto-feed
_RATINGS_PATH = os.path.join(_ROOT, "static", "data", "credit_ratings.json")
_CREDIT_RATINGS: dict = {}
try:
    with open(_RATINGS_PATH, "r", encoding="utf-8") as _f:
        _CREDIT_RATINGS = json.load(_f)
    print(f"[ratings] Loaded {len(_CREDIT_RATINGS) - 1} ticker ratings")
except Exception as _e:
    print(f"[ratings] Warning: {_e}")

def _auto_rating(ticker: str):
    return _CREDIT_RATINGS.get(ticker.upper(), {}).get("wacc_rating") or None


def _inject_ratios_cache(ticker: str, sm: dict) -> None:
    """Pre-populate fmp_3statementv6._RATIOS_CACHE from stored scorecard_metrics.

    Constructs synthetic ratio entries so build_scorecard() and build_dcf() see
    historically-accurate pe_5yr_avg / pfcf_5yr_avg / ev_ebitda_5yr_avg without
    hitting FMP.  Five identical entries are injected — their average equals the
    stored historical average exactly.
    """
    pe_avg    = sm.get("pe_5yr_avg")
    pfcf_avg  = sm.get("pfcf_5yr_avg")
    ev_avg    = sm.get("ev_ebitda_5yr_avg")
    pe_cur    = sm.get("pe_current")
    pfcf_cur  = sm.get("pfcf_current")

    # Build 5 entries: index-0 uses the "current" values so trailing_pe/trailing_pfcf
    # are as accurate as possible; entries 1-4 pad to reproduce the 5yr averages.
    def _make_entry(pe_v, pfcf_v, ev_v):
        e = {}
        if pe_v and pe_v > 0:    e["priceToEarningsRatio"]    = pe_v
        if pfcf_v and pfcf_v > 0: e["priceToFreeCashFlowRatio"] = pfcf_v
        if ev_v and ev_v > 0:    e["enterpriseValueMultiple"]  = ev_v
        return e

    entries = [_make_entry(pe_cur or pe_avg, pfcf_cur or pfcf_avg, ev_avg)]
    for _ in range(4):
        entries.append(_make_entry(pe_avg, pfcf_avg, ev_avg))

    # Only inject if at least one meaningful field exists
    if any(entries[0]):
        key = f"{ticker.upper()}:5"
        mdl._RATIOS_CACHE[key] = entries
        print(f"  [cache] Injected ratios: pe_avg={pe_avg}  pfcf_avg={pfcf_avg}  ev_avg={ev_avg}")
    else:
        print(f"  [cache] No ratio data to inject for {ticker}")


def _read_csv():
    if not os.path.exists(CSV_PATH):
        return {}
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        content = _schema.migrate(f.read())
    rows = {}
    for row in csv.DictReader(io.StringIO(content)):
        t = row.get("Ticker", "").strip().upper()
        if t:
            rows[t] = row
    return rows


def _write_row(ticker, scorecard_metrics, dcf_prices,
               current_price, market_cap, is_data, cf_data,
               biz_clarity=None, ltp=None, display_score=None):
    sm = scorecard_metrics or {}
    dp = dcf_prices or {}

    def _f(v, dp_=4):
        return "" if v is None else f"{v:.{dp_}f}"

    _fx = dp.get("fx_to_usd") or 1.0   # reportedCurrency → USD for display cols
    rev_b   = (is_data[-1].get("revenue") or 0) / 1e9 * _fx if is_data else None
    ocf_b   = (cf_data[-1].get("operatingCashFlow") or 0) / 1e9 * _fx if cf_data else None
    fcf_raw = cf_data[-1].get("freeCashFlow") if cf_data else None
    if fcf_raw is None and cf_data:
        fcf_raw = ((cf_data[-1].get("operatingCashFlow") or 0) +
                   (cf_data[-1].get("capitalExpenditure") or 0))
    fcf_b     = (fcf_raw / 1e9 * _fx) if fcf_raw is not None else None
    mkt_cap_b = (market_cap / 1e9) if market_cap else None

    _auto  = sm.get("auto_score")
    _score = display_score if display_score is not None else _auto

    new_row = {
        "Ticker":         ticker,
        "Price":          _f(current_price, 2),
        "MktCap_B":       _f(mkt_cap_b, 2),
        "GG_Price":       _f(dp.get("gg_price"),  2),
        "GG_Upside":      _f(dp.get("gg_upside"), 4),
        "EM_Price":       _f(dp.get("em_price"),  2),
        "EM_Upside":      _f(dp.get("em_upside"), 4),
        "PE_Current":     _f(sm.get("pe_current"),   1),
        "PE_5yr":         _f(sm.get("pe_5yr_avg"),   1),
        "PFCF_Current":   _f(sm.get("pfcf_current"), 1),
        "PFCF_5yr":       _f(sm.get("pfcf_5yr_avg"), 1),
        "ROIC":           _f(sm.get("roic")),
        "Rev_CAGR":       _f(sm.get("rev_cagr")),
        "FCF_NI":         _f(sm.get("fcf_ni")),
        "D_EBITDA":       _f(sm.get("d_ebitda"), 2),
        "Revenue_B":      _f(rev_b,  2),
        "OCF_B":          _f(ocf_b,  2),
        "FCF_B":          _f(fcf_b,  2),
        "Auto_Score":     "" if _score is None else str(_score),
        "Floor_Cap":      "" if sm.get("floor_cap") is None else str(sm["floor_cap"]),
        "Manual_Clarity": biz_clarity or "",
        "Manual_LTP":     ltp or "",
        "Date":           datetime.date.today().isoformat(),
    }
    new_line = ",".join(new_row.get(c, "") for c in _schema.COLUMNS)

    if os.path.exists(CSV_PATH):
        with open(CSV_PATH, "r", encoding="utf-8") as f:
            existing = _schema.migrate(f.read())
    else:
        existing = _schema.HEADER

    lines      = existing.splitlines()
    header_ln  = lines[0] if lines else _schema.HEADER.rstrip()
    data_lines = [l for l in lines[1:] if l.strip()]
    data_lines = [l for l in data_lines
                  if l.split(",")[0].strip().upper() != ticker]
    data_lines.append(new_line)
    data_lines.sort(key=lambda l: l.split(",")[0].strip())

    updated = header_ln + "\n" + "\n".join(data_lines) + "\n"
    with open(CSV_PATH, "w", encoding="utf-8") as f:
        f.write(updated)


def run_ticker(ticker: str, old_row: dict) -> dict:
    """Re-score one ticker from cached data. Returns result dict."""
    ticker = ticker.upper()
    old_score  = old_row.get("Auto_Score", "")
    bc_manual  = old_row.get("Manual_Clarity", "") or None
    ltp_manual = old_row.get("Manual_LTP", "") or None

    # Load cached data from static/data/{ticker}_data.json
    cached = load_ticker_data(ticker)
    if not cached:
        return {"ticker": ticker, "error": "no cached data", "old_score": old_score}

    is_data      = cached.get("is_data", [])
    bs_data      = cached.get("bs_data", [])
    cf_data      = cached.get("cf_data", [])
    profile      = cached.get("profile", {})
    years        = cached.get("years", [])
    analyst_ests = cached.get("analyst_ests", [])
    prev_sm      = cached.get("scorecard_metrics", {})

    if not (is_data and bs_data and cf_data):
        return {"ticker": ticker, "error": "incomplete cached financials", "old_score": old_score}

    current_price = float(profile.get("price") or 0) or None
    market_cap    = float(profile.get("mktCap") or profile.get("marketCap") or 0) or None
    wacc_rating   = _auto_rating(ticker)

    print(f"  Loaded cache (fetched: {cached.get('fetched','?')})  "
          f"price={current_price}  rating={wacc_rating}  analyst_ests={len(analyst_ests)}")

    # Inject stored ratios into in-process cache to avoid FMP calls
    _inject_ratios_cache(ticker, prev_sm)

    logs = []
    _orig_print = builtins.print
    builtins.print = lambda *a, **k: logs.append(" ".join(str(x) for x in a))

    try:
        _bank_credit = mdl.fetch_bank_credit_data(ticker)   # EDGAR, not FMP

        wb       = Workbook()
        pl_refs  = mdl.build_pl(wb, is_data, years, ticker)
        mdl.build_cover(wb, ticker, years, is_data)
        bs_refs  = mdl.build_bs(wb, bs_data, years, ticker)
        cf_refs  = mdl.build_cf(wb, cf_data, years, ticker)
        mdl.build_ratios(wb, is_data, bs_data, cf_data, years, ticker,
                         pl_refs, bs_refs, cf_refs, bank_credit=_bank_credit)
        mdl.build_segments(wb, ticker, years)

        # WACC — pass profile (has beta) and credit rating (skips FMP /ratings call)
        wacc_refs = mdl.build_wacc(wb, ticker, is_data, bs_data,
                                   manual_rating=wacc_rating, profile=profile)

        dcf_refs  = mdl.build_dcf(
            wb, ticker, is_data, bs_data, cf_data, years,
            pl_refs, bs_refs, wacc_refs,
            current_price=current_price, cf_refs=cf_refs
        )
        dcf_prices = (dcf_refs or {}).get("dcf_prices") or {}

        _, scorecard_metrics = mdl.build_scorecard(
            wb, ticker, is_data, bs_data, cf_data, years,
            biz_clarity=bc_manual,
            ltp=ltp_manual,
            dcf_gg_price=dcf_prices.get("gg_price"),
            evs_regime=bool(dcf_prices.get("evs_regime")),
            bank_credit=_bank_credit,
            analyst_ests=analyst_ests,
            profile=profile,
        )

        auto_score     = scorecard_metrics.get("auto_score") or 0
        auto_score_raw = scorecard_metrics.get("auto_score_raw") or 0
        _W      = scorecard_metrics.get("weights") or {}
        _w_bc   = float(_W.get("BC",  2.5))
        _w_ltp  = float(_W.get("LTP", 10.0))
        bc_pts  = TIER_PTS.get(bc_manual,  0) * _w_bc  / 10
        ltp_pts = TIER_PTS.get(ltp_manual, 0) * _w_ltp / 10
        adj_score = round((auto_score_raw + bc_pts + ltp_pts) / 10, 1)
        floor_cap = scorecard_metrics.get("floor_cap")
        if floor_cap is not None:
            adj_score = min(adj_score, float(floor_cap))

        _display_score = adj_score if (bc_manual or ltp_manual) else auto_score

        report_data = build_report_data(
            ticker=ticker, profile=profile,
            is_data=is_data, bs_data=bs_data, cf_data=cf_data, years=years,
            wacc_val=wacc_refs.get("wacc_val"),
            dcf_prices=dcf_prices,
            scorecard_metrics=scorecard_metrics,
            manual_rating=wacc_rating,
            current_price=current_price, market_cap=market_cap,
            biz_clarity=bc_manual, ltp=ltp_manual,
            adj_score=adj_score, analyst_ests=analyst_ests,
        )
        html_content = render_html_report(report_data)
        with open(os.path.join(RPT_DIR, f"{ticker}_report.html"), "w", encoding="utf-8") as f:
            f.write(html_content)

        buf = io.BytesIO(); wb.save(buf)
        with open(os.path.join(XLS_DIR, f"{ticker}_model.xlsx"), "wb") as f:
            f.write(buf.getvalue())

        save_ticker_data(
            ticker, is_data, bs_data, cf_data, profile, years,
            wacc_refs.get("wacc_val"), dcf_prices, scorecard_metrics, analyst_ests
        )

        _write_row(
            ticker=ticker,
            scorecard_metrics=scorecard_metrics,
            dcf_prices=dcf_prices,
            current_price=current_price,
            market_cap=market_cap,
            is_data=is_data,
            cf_data=cf_data,
            biz_clarity=bc_manual,
            ltp=ltp_manual,
            display_score=_display_score,
        )

        builtins.print = _orig_print

        return {
            "ticker":     ticker,
            "old_score":  old_score,
            "auto":       auto_score,
            "adj":        adj_score,
            "display":    _display_score,
            "floor_cap":  floor_cap,
            "bucket":     scorecard_metrics.get("sector_bucket", ""),
            "tier_pe":    scorecard_metrics.get("tier_pe"),
            "tier_pfcf":  scorecard_metrics.get("tier_pfcf"),
            "pe_current": scorecard_metrics.get("pe_current"),
            "pe_5yr":     scorecard_metrics.get("pe_5yr_avg"),
            "wacc":       wacc_refs.get("wacc_val"),
            "gg_up":      dcf_prices.get("gg_upside"),
            "em_up":      dcf_prices.get("em_upside"),
            "em_base":    dcf_prices.get("em_base_mult"),
            "em_anchored": dcf_prices.get("em_anchored"),
        }

    except Exception as e:
        builtins.print = _orig_print
        print(f"  FAIL: {e}")
        traceback.print_exc()
        return {"ticker": ticker, "error": str(e), "old_score": old_score}


def main():
    # Default: all tickers with cached data that haven't been run with Rule 21
    # or that were fetched before latest engine changes.
    if len(sys.argv) > 1:
        tickers = [t.upper() for t in sys.argv[1:]]
    else:
        tickers = [
            # May 25 batch (Rule 21 applied, but missing P/E fix + WACC auto-feed)
            "CSCO", "HCA", "META", "NFLX", "TGT", "TSLA", "TSM",
            "ABBV", "KO", "PEP", "V",
            # May 24 batch (Rule 21 NOT applied + missing P/E fix + WACC auto-feed)
            "BAC", "C", "CVX", "F", "FDX", "INTC", "JPM", "SNAP", "SOFI",
        ]

    old_rows = _read_csv()
    results  = []

    print(f"\nOffline re-score: {len(tickers)} tickers (no FMP API calls)")
    print("=" * 70)

    for i, ticker in enumerate(tickers, 1):
        print(f"\n[{i}/{len(tickers)}]  {ticker}")
        old_row = old_rows.get(ticker, {})
        result  = run_ticker(ticker, old_row)
        results.append(result)

        if result.get("error"):
            print(f"  FAILED: {result['error']}")
        else:
            try:
                d = result["auto"] - float(result["old_score"])
                d_str = f"d{d:+.1f}"
            except Exception:
                d_str = "dN/A"
            cap_str = f"  cap={result['floor_cap']}" if result.get("floor_cap") else ""
            print(f"  OK  old={result['old_score']}  new={result['auto']}  {d_str}"
                  f"  adj={result['adj']}{cap_str}  [{result['bucket']}]")
            print(f"  tier_pe={result['tier_pe']}  pe_cur={result['pe_current']}  pe5={result['pe_5yr']}"
                  f"  wacc={result['wacc'] and round(result['wacc']*100,2)}%")
            print(f"  em_base={result['em_base']}  em_anchored={result['em_anchored']}"
                  f"  gg_up={result['gg_up'] and round(result['gg_up']*100,1)}%"
                  f"  em_up={result['em_up'] and round(result['em_up']*100,1)}%")

    print("\n" + "=" * 70)
    print("SUMMARY  (sorted by auto_score desc)")
    print("=" * 70)
    ok = [r for r in results if not r.get("error")]
    ok.sort(key=lambda r: r.get("auto") or 0, reverse=True)
    for r in ok:
        try:
            d = r["auto"] - float(r["old_score"])
            d_str = f"{d:+.1f}"
        except Exception:
            d_str = " N/A"
        print(f"  {r['ticker']:<5}  old={r.get('old_score') or 'N/A':>5}  "
              f"new={r.get('auto') or 'N/A':>5}  d={d_str:>5}  adj={r.get('adj') or 'N/A'}"
              f"  tier_pe={str(r.get('tier_pe') or 'N/A'):<9}  em_anch={r.get('em_anchored')}")
    failed = [r for r in results if r.get("error")]
    if failed:
        print("\nFAILED:")
        for r in failed:
            print(f"  {r['ticker']}: {r['error']}")

    print(f"\n{len(ok)} OK  |  {len(failed)} failed")
    print("outputs.csv updated. Git push to deploy.")


if __name__ == "__main__":
    main()
