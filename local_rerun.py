"""
local_rerun.py
Re-runs every ticker in outputs.csv through the LOCAL updated engine
(fmp_3statementv6.py), refreshes outputs.csv, and prints a before/after
score comparison so scoring shifts are immediately visible.

Run: python local_rerun.py
"""

import builtins, csv, datetime, io, os, sys, time, traceback
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests as _req
from openpyxl import Workbook

import fmp_3statementv6 as mdl
from report_bridge import build_report_data, render_html_report
from data_store import save_ticker_data
import csv_schema as _schema

# ── Paths ─────────────────────────────────────────────────────────────────────
_ROOT     = os.path.dirname(os.path.abspath(__file__))
CSV_PATH  = os.path.join(_ROOT, "outputs.csv")
RPT_DIR   = os.path.join(_ROOT, "static", "reports")
XLS_DIR   = os.path.join(_ROOT, "static", "excel")
os.makedirs(RPT_DIR, exist_ok=True)
os.makedirs(XLS_DIR, exist_ok=True)

# Discrete tier pts used only for BC/LTP manual add-on
TIER_PTS = {"HIGH": 10, "MOD-HIGH": 7, "MOD": 7, "MOD-LOW": 3, "LOW": 0}

DELAY = 12  # seconds between tickers — 5 FMP calls each


# ── Read current outputs.csv ──────────────────────────────────────────────────
def _read_csv() -> dict:
    """Return {TICKER: row_dict} from current outputs.csv."""
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


# ── Write one row back to outputs.csv ────────────────────────────────────────
def _write_row(ticker, scorecard_metrics, dcf_prices,
               current_price, market_cap, is_data, cf_data,
               biz_clarity=None, ltp=None, display_score=None):
    """Write one row to outputs.csv.

    display_score — the score that appears in the report hero (adj_score when
    quals are present, auto_score otherwise).  This is the authoritative value
    written to Auto_Score so the dashboard always matches the report.
    """
    sm  = scorecard_metrics or {}
    dp  = dcf_prices or {}

    def _f(v, dp_=4):
        return "" if v is None else f"{v:.{dp_}f}"

    rev_b  = (is_data[-1].get("revenue") or 0) / 1e9 if is_data else None
    ocf_b  = (cf_data[-1].get("operatingCashFlow") or 0) / 1e9 if cf_data else None
    fcf_raw = cf_data[-1].get("freeCashFlow") if cf_data else None
    if fcf_raw is None and cf_data:
        fcf_raw = ((cf_data[-1].get("operatingCashFlow") or 0) +
                   (cf_data[-1].get("capitalExpenditure") or 0))
    fcf_b     = (fcf_raw / 1e9) if fcf_raw is not None else None
    mkt_cap_b = (market_cap / 1e9) if market_cap else None

    # Auto_Score must equal the hero score shown in the HTML report.
    # Use display_score when provided; fall back to quant auto_score only.
    _stored_score = display_score if display_score is not None else sm.get("auto_score")

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
        "SBC_B":          _f(sm.get("sbc_trailing_b"), 2),
        "SBC_pct":        _f(sm.get("sbc_pct_fcf"), 4),
        "Auto_Score":     "" if _stored_score is None else str(_stored_score),
        "Floor_Cap":      "" if sm.get("floor_cap")  is None else str(sm["floor_cap"]),
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


# ── Main run ──────────────────────────────────────────────────────────────────
def main():
    old_rows  = _read_csv()
    all_tickers = sorted(old_rows.keys()) if old_rows else []

    REMAINING = {"WMT", "TSM", "SOFI", "TGT", "SNAP"}
    tickers = [t for t in all_tickers if t in REMAINING]

    if not tickers:
        print("outputs.csv is empty or missing — nothing to re-run.")
        return

    print(f"Re-running {len(tickers)} tickers: {', '.join(tickers)}")
    print(f"Delay between tickers: {DELAY}s\n")

    results = []

    for i, ticker in enumerate(tickers, 1):
        old_row  = old_rows.get(ticker, {})
        old_score = old_row.get("Auto_Score", "")
        bc_manual = old_row.get("Manual_Clarity", "") or None
        ltp_manual = old_row.get("Manual_LTP", "") or None

        print(f"[{i:2d}/{len(tickers)}]  {ticker:<6}  old_score={old_score or 'N/A':<6}", end="  ")

        logs = []
        _orig = builtins.print
        builtins.print = lambda *a, **k: logs.append(" ".join(str(x) for x in a))

        try:
            is_data = mdl.fetch("income-statement",        ticker)[:mdl.YEARS][::-1]
            bs_data = mdl.fetch("balance-sheet-statement", ticker)[:mdl.YEARS][::-1]
            cf_data = mdl.fetch("cash-flow-statement",     ticker)[:mdl.YEARS][::-1]

            if not is_data:
                raise ValueError("No income statement data")

            # Profile
            profile = {}; current_price = None; market_cap = None
            try:
                _p = _req.get(
                    f"https://financialmodelingprep.com/stable/profile"
                    f"?symbol={ticker}&apikey={mdl.API_KEY}", timeout=8
                ).json()
                profile       = _p[0] if isinstance(_p, list) and _p else _p or {}
                current_price = float(profile.get("price") or 0) or None
                market_cap    = float(profile.get("mktCap") or profile.get("marketCap") or 0) or None
            except Exception:
                pass

            years = [
                d.get("fiscalYear") or d.get("calendarYear") or d["date"][:4]
                for d in is_data
            ]

            # Analyst estimates
            analyst_ests = []
            try:
                _ae = _req.get(
                    f"https://financialmodelingprep.com/stable/analyst-estimates"
                    f"?symbol={ticker}&period=annual&limit=5&apikey={mdl.API_KEY}", timeout=8
                ).json()
                if isinstance(_ae, list):
                    analyst_ests = sorted(
                        [e for e in _ae if str(e.get("date",""))[:4] > str(years[-1])],
                        key=lambda x: x.get("date", "")
                    )[:5]
            except Exception:
                pass

            # Pre-fetch EDGAR bank credit data (free API, no quota impact)
            # Only fetches for known bank tickers; returns {} for all others
            _bank_credit = mdl.fetch_bank_credit_data(ticker)

            # Build workbook
            wb        = Workbook()
            pl_refs   = mdl.build_pl(wb, is_data, years, ticker)
            mdl.build_cover(wb, ticker, years, is_data)
            bs_refs   = mdl.build_bs(wb, bs_data, years, ticker)
            cf_refs   = mdl.build_cf(wb, cf_data, years, ticker)
            mdl.build_ratios(wb, is_data, bs_data, cf_data, years, ticker, pl_refs, bs_refs, cf_refs,
                             bank_credit=_bank_credit)
            mdl.build_segments(wb, ticker, years)
            wacc_refs = mdl.build_wacc(wb, ticker, is_data, bs_data, None,
                                       profile=profile)
            dcf_refs  = mdl.build_dcf(
                wb, ticker, is_data, bs_data, cf_data, years,
                pl_refs, bs_refs, wacc_refs, current_price=current_price, cf_refs=cf_refs,
                profile=profile,
            )
            _, scorecard_metrics = mdl.build_scorecard(
                wb, ticker, is_data, bs_data, cf_data, years,
                biz_clarity=bc_manual or None,
                ltp=ltp_manual or None,
                dcf_gg_price=(dcf_refs.get("dcf_prices") or {}).get("gg_price"),
                evs_regime=bool((dcf_refs.get("dcf_prices") or {}).get("evs_regime")),
                bank_credit=_bank_credit,
                analyst_ests=analyst_ests,
                profile=profile,
            )

            # Adj score
            auto_score     = scorecard_metrics.get("auto_score") or 0      # normalized 0-10, for display/printing
            auto_score_raw = scorecard_metrics.get("auto_score_raw") or 0  # raw 0-87.5, for adj base
            _W      = scorecard_metrics.get("weights") or {}
            _w_bc   = float(_W.get("BC",  2.5))
            _w_ltp  = float(_W.get("LTP", 10.0))
            bc_pts  = TIER_PTS.get(bc_manual,  0) * _w_bc  / 10
            ltp_pts = TIER_PTS.get(ltp_manual, 0) * _w_ltp / 10
            adj_score = round((auto_score_raw + bc_pts + ltp_pts) / 10, 1)  # normalize to 0-10
            floor_cap = scorecard_metrics.get("floor_cap")  # already normalized 0-10
            if floor_cap is not None:
                adj_score = min(adj_score, float(floor_cap))

            # Build + save HTML report
            dcf_prices = (dcf_refs or {}).get("dcf_prices") or {}
            report_data = build_report_data(
                ticker=ticker, profile=profile,
                is_data=is_data, bs_data=bs_data, cf_data=cf_data, years=years,
                wacc_val=wacc_refs.get("wacc_val"),
                dcf_prices=dcf_prices,
                scorecard_metrics=scorecard_metrics,
                manual_rating=None,
                current_price=current_price, market_cap=market_cap,
                biz_clarity=bc_manual or None, ltp=ltp_manual or None,
                adj_score=adj_score, analyst_ests=analyst_ests,
            )
            html_content = render_html_report(report_data)
            with open(os.path.join(RPT_DIR, f"{ticker}_report.html"), "w", encoding="utf-8") as f:
                f.write(html_content)

            # Save Excel
            buf = io.BytesIO(); wb.save(buf)
            with open(os.path.join(XLS_DIR, f"{ticker}_model.xlsx"), "wb") as f:
                f.write(buf.getvalue())

            # Save DCF data cache (including price_history + consensus_pt so
            # re-renders make ZERO additional FMP calls)
            save_ticker_data(
                ticker, is_data, bs_data, cf_data, profile, years,
                wacc_refs.get("wacc_val"), dcf_prices, scorecard_metrics, analyst_ests,
                price_history=report_data.get("_price_history_payload"),
                consensus_pt=report_data.get("_consensus_pt_payload"),
            )

            # Update outputs.csv — store the score that appears in the report hero
            # (adj_score when quals are entered, auto_score when quant-only).
            _display = adj_score if (bc_manual or ltp_manual) else auto_score
            _write_row(
                ticker=ticker,
                scorecard_metrics=scorecard_metrics,
                dcf_prices=dcf_prices,
                current_price=current_price,
                market_cap=market_cap,
                is_data=is_data,
                cf_data=cf_data,
                biz_clarity=bc_manual or None,
                ltp=ltp_manual or None,
                display_score=_display,
            )

            builtins.print = _orig
            d_str = ""
            try:
                d_str = f"  d{auto_score - float(old_score):+.1f}"
            except Exception:
                pass
            bucket = scorecard_metrics.get("sector_bucket", "")
            cap_str = f"  cap={floor_cap}" if floor_cap else ""
            print(f"OK  new={auto_score:<6}{d_str}  adj={adj_score}{cap_str}  [{bucket}]")

            results.append({
                "ticker":    ticker,
                "old_score": old_score,
                "auto":      auto_score,
                "adj":       adj_score,
                "floor_cap": floor_cap,
                "bucket":    bucket,
                "price":     current_price,
                "gg_up":     dcf_prices.get("gg_upside"),
                "em_up":     dcf_prices.get("em_upside"),
                "evs":       dcf_prices.get("evs_regime", False),
            })

        except Exception as e:
            builtins.print = _orig
            print(f"FAIL  {e}")
            traceback.print_exc()
            results.append({"ticker": ticker, "error": str(e), "old_score": old_score})

        if i < len(tickers):
            time.sleep(DELAY)

    # ── Summary ───────────────────────────────────────────────────────────────
    print("\n\n" + "=" * 76)
    print("SCORE SUMMARY  (sorted by new auto_score desc)")
    print("=" * 76)
    print(f"  {'Ticker':<6}  {'Old':>6}  {'New':>6}  {'d':>6}  {'Adj':>6}  {'Cap':>5}  {'Bucket'}")
    print(f"  {'-'*70}")

    ok = [r for r in results if not r.get("error")]
    ok.sort(key=lambda r: r.get("auto") or -1, reverse=True)

    for r in ok:
        old_f = float(r["old_score"]) if r.get("old_score") else None
        delta = f"{r['auto'] - old_f:+.1f}" if (old_f is not None and r.get("auto") is not None) else "N/A"
        cap_s = f"{r['floor_cap']:.0f}" if r.get("floor_cap") else ""
        evs_s = " [EVS]" if r.get("evs") else ""
        print(f"  {r['ticker']:<6}  {r.get('old_score') or 'N/A':>6}  "
              f"{r.get('auto') or 'N/A':>6}  {delta:>6}  "
              f"{r.get('adj') or 'N/A':>6}  {cap_s:>5}  {r.get('bucket','')}{evs_s}")

    if any(r.get("error") for r in results):
        print("\nFAILED:")
        for r in results:
            if r.get("error"):
                print(f"  {r['ticker']}: {r['error']}")

    total = len(ok)
    improved = sum(1 for r in ok if r.get("old_score") and r.get("auto") is not None
                   and r["auto"] > float(r["old_score"]))
    declined = sum(1 for r in ok if r.get("old_score") and r.get("auto") is not None
                   and r["auto"] < float(r["old_score"]))
    print(f"\n  {total} tickers OK  |  {improved} improved  |  {declined} declined")
    print(f"\noutputs.csv updated. Dashboard will reflect new scores.")


if __name__ == "__main__":
    main()
