"""
Re-scores COST, JNJ, JPM, BAC from locally-cached FMP data — no API calls.
Uses the current engine (SBC metrics, exec-risk fix, DCF-contamination removal, PEG modifier).
"""
import builtins, csv, datetime, io, os, sys, traceback
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
TICKERS  = ["COST", "JNJ", "JPM", "BAC"]


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
               biz_clarity=None, ltp=None):
    sm = scorecard_metrics or {}
    dp = dcf_prices or {}

    def _f(v, dp_=4):
        return "" if v is None else f"{v:.{dp_}f}"

    rev_b   = (is_data[-1].get("revenue") or 0) / 1e9 if is_data else None
    ocf_b   = (cf_data[-1].get("operatingCashFlow") or 0) / 1e9 if cf_data else None
    fcf_raw = cf_data[-1].get("freeCashFlow") if cf_data else None
    if fcf_raw is None and cf_data:
        fcf_raw = ((cf_data[-1].get("operatingCashFlow") or 0) +
                   (cf_data[-1].get("capitalExpenditure") or 0))
    fcf_b     = (fcf_raw / 1e9) if fcf_raw is not None else None
    mkt_cap_b = (market_cap / 1e9) if market_cap else None

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
        "Auto_Score":     "" if sm.get("auto_score") is None else str(sm["auto_score"]),
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
    data_lines = [l for l in data_lines if l.split(",")[0].strip().upper() != ticker]
    data_lines.append(new_line)
    data_lines.sort(key=lambda l: l.split(",")[0].strip())

    updated = header_ln + "\n" + "\n".join(data_lines) + "\n"
    with open(CSV_PATH, "w", encoding="utf-8") as f:
        f.write(updated)


def main():
    old_rows = _read_csv()
    results  = []

    for i, ticker in enumerate(TICKERS, 1):
        old_row    = old_rows.get(ticker, {})
        old_score  = old_row.get("Auto_Score", "")
        bc_manual  = old_row.get("Manual_Clarity", "") or None
        ltp_manual = old_row.get("Manual_LTP", "") or None

        print(f"\n[{i}/{len(TICKERS)}]  {ticker:<6}  old_score={old_score or 'N/A'}", flush=True)

        cached = load_ticker_data(ticker)
        if not cached:
            print(f"  No cached data — skipping.")
            results.append({"ticker": ticker, "error": "no cache", "old_score": old_score})
            continue

        is_data      = cached["is_data"]
        bs_data      = cached["bs_data"]
        cf_data      = cached["cf_data"]
        profile      = cached.get("profile", {})
        years        = cached["years"]
        analyst_ests = cached.get("analyst_ests", [])
        current_price = float(profile.get("price") or 0) or None
        market_cap    = float(profile.get("mktCap") or profile.get("marketCap") or 0) or None

        print(f"  cache fetched={cached.get('fetched','?')}  price={current_price}  "
              f"bc={bc_manual}  ltp={ltp_manual}", flush=True)

        logs = []
        _orig = builtins.print
        builtins.print = lambda *a, **k: logs.append(" ".join(str(x) for x in a))

        try:
            _bank_credit = mdl.fetch_bank_credit_data(ticker)

            wb       = Workbook()
            pl_refs  = mdl.build_pl(wb, is_data, years, ticker)
            mdl.build_cover(wb, ticker, years, is_data)
            bs_refs  = mdl.build_bs(wb, bs_data, years, ticker)
            cf_refs  = mdl.build_cf(wb, cf_data, years, ticker)
            mdl.build_ratios(wb, is_data, bs_data, cf_data, years, ticker,
                             pl_refs, bs_refs, cf_refs, bank_credit=_bank_credit)
            mdl.build_segments(wb, ticker, years)
            wacc_refs = mdl.build_wacc(wb, ticker, is_data, bs_data, None)
            dcf_refs  = mdl.build_dcf(
                wb, ticker, is_data, bs_data, cf_data, years,
                pl_refs, bs_refs, wacc_refs,
                current_price=current_price, cf_refs=cf_refs
            )
            _, scorecard_metrics = mdl.build_scorecard(
                wb, ticker, is_data, bs_data, cf_data, years,
                biz_clarity=bc_manual,
                ltp=ltp_manual,
                dcf_gg_price=(dcf_refs.get("dcf_prices") or {}).get("gg_price"),
                evs_regime=bool((dcf_refs.get("dcf_prices") or {}).get("evs_regime")),
                bank_credit=_bank_credit,
                analyst_ests=analyst_ests,
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

            dcf_prices  = (dcf_refs or {}).get("dcf_prices") or {}
            report_data = build_report_data(
                ticker=ticker, profile=profile,
                is_data=is_data, bs_data=bs_data, cf_data=cf_data, years=years,
                wacc_val=wacc_refs.get("wacc_val"),
                dcf_prices=dcf_prices,
                scorecard_metrics=scorecard_metrics,
                manual_rating=None,
                current_price=current_price, market_cap=market_cap,
                biz_clarity=bc_manual, ltp=ltp_manual,
                adj_score=adj_score, analyst_ests=analyst_ests,
            )
            html_content = render_html_report(report_data)
            with open(os.path.join(RPT_DIR, f"{ticker}_report.html"), "w", encoding="utf-8") as f:
                f.write(html_content)

            import io as _io
            buf = _io.BytesIO(); wb.save(buf)
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
            )

            builtins.print = _orig

            sbc_pct = scorecard_metrics.get("sbc_pct_fcf")
            sbc_str = f"SBC={sbc_pct:.1%}" if sbc_pct else "SBC=N/A"
            d_str   = f"  d{auto_score - float(old_score):+.1f}" if old_score else ""
            cap_str = f"  cap={floor_cap}" if floor_cap else ""
            print(f"  OK  new={auto_score}{d_str}  adj={adj_score}{cap_str}  "
                  f"[{scorecard_metrics.get('sector_bucket','')}]  {sbc_str}", flush=True)

            results.append({
                "ticker":    ticker,
                "old_score": old_score,
                "auto":      auto_score,
                "adj":       adj_score,
                "floor_cap": floor_cap,
                "sbc_pct":   sbc_pct,
                "pfcf_adj":  scorecard_metrics.get("pfcf_adj"),
                "pfcf_rep":  scorecard_metrics.get("pfcf_current"),
                "gg_up":     dcf_prices.get("gg_upside"),
                "em_up":     dcf_prices.get("em_upside"),
            })

        except Exception as e:
            builtins.print = _orig
            print(f"  FAIL: {e}", flush=True)
            traceback.print_exc()
            results.append({"ticker": ticker, "error": str(e), "old_score": old_score})

    print("\n" + "="*70)
    print(f"  {'Ticker':<6}  {'Before':>6}  {'After':>6}  {'Delta':>6}  {'Adj':>6}  SBC/FCF  P/FCF adj")
    print(f"  {'-'*65}")
    for r in results:
        if r.get("error"):
            print(f"  {r['ticker']:<6}  FAILED: {r['error']}")
        else:
            old_f = float(r["old_score"]) if r.get("old_score") else None
            delta = f"{r['auto'] - old_f:+.1f}" if old_f is not None else " N/A"
            sbc   = f"{r['sbc_pct']:.1%}" if r.get("sbc_pct") else "  N/A"
            adj_p = f"{r['pfcf_adj']:.1f}x" if r.get("pfcf_adj") else "  N/A"
            rep_p = f"{r['pfcf_rep']:.1f}x" if r.get("pfcf_rep") else "  N/A"
            print(f"  {r['ticker']:<6}  {r.get('old_score') or 'N/A':>6}  "
                  f"{r.get('auto') or 'N/A':>6}  {delta:>6}  {r.get('adj') or 'N/A':>6}  "
                  f"{sbc:>7}  {rep_p} -> {adj_p}")
    print("="*70)


if __name__ == "__main__":
    main()
