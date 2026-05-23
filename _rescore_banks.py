"""
_rescore_banks.py
Re-runs build_scorecard() for bank tickers where is_bank was wrongly False,
using cached financial data (no FMP API calls).
Then re-renders the HTML report and updates outputs.csv.

Tickers affected: JPM, BAC, SOFI (is_bank=False despite sector_bucket=bank)
"""
import csv, io, os, sys, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook
from data_store import load_ticker_data, save_ticker_data
from report_bridge import build_report_data, render_html_report
import fmp_3statementv6 as mdl
import csv_schema as _schema

_ROOT    = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(_ROOT, "outputs.csv")
RPT_DIR  = os.path.join(_ROOT, "static", "reports")

TIER_PTS = {"HIGH": 10, "MOD-HIGH": 7, "MOD": 7, "MOD-LOW": 3, "LOW": 0}

# Tickers where is_bank=False but sector_bucket=bank
TARGETS = ["JPM", "BAC", "SOFI"]

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

def _write_score(ticker, display_score, sm, dp, current_price, market_cap,
                 is_data, cf_data, bc_manual, ltp_manual):
    """Upsert one row in outputs.csv."""
    def _f(v, p=4):
        return "" if v is None else f"{v:.{p}f}"
    rev_b  = (is_data[-1].get("revenue") or 0) / 1e9 if is_data else None
    ocf_b  = (cf_data[-1].get("operatingCashFlow") or 0) / 1e9 if cf_data else None
    fcf_raw = cf_data[-1].get("freeCashFlow") if cf_data else None
    if fcf_raw is None and cf_data:
        fcf_raw = ((cf_data[-1].get("operatingCashFlow") or 0) +
                   (cf_data[-1].get("capitalExpenditure") or 0))
    fcf_b = (fcf_raw / 1e9) if fcf_raw is not None else None
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
        "Auto_Score":     "" if display_score is None else str(display_score),
        "Floor_Cap":      "" if sm.get("floor_cap") is None else str(sm["floor_cap"]),
        "Manual_Clarity": bc_manual or "",
        "Manual_LTP":     ltp_manual or "",
        "Date":           datetime.date.today().isoformat(),
    }
    new_line = ",".join(new_row.get(c, "") for c in _schema.COLUMNS)

    if os.path.exists(CSV_PATH):
        with open(CSV_PATH, "r", encoding="utf-8") as f:
            existing = _schema.migrate(f.read())
    else:
        existing = _schema.HEADER
    lines = existing.splitlines()
    header = lines[0] if lines else _schema.HEADER.rstrip()
    data   = [l for l in lines[1:] if l.strip()]
    data   = [l for l in data if l.split(",")[0].strip().upper() != ticker]
    data.append(new_line)
    data.sort(key=lambda l: l.split(",")[0].strip())
    with open(CSV_PATH, "w", encoding="utf-8") as f:
        f.write(header + "\n" + "\n".join(data) + "\n")


def main():
    csv_rows = _read_csv()

    for ticker in TARGETS:
        cached = load_ticker_data(ticker)
        if not cached:
            print(f"{ticker}: not cached — skipping")
            continue

        sm_old = cached.get("scorecard_metrics") or {}
        print(f"\n{'='*60}")
        print(f"{ticker}  is_bank_old={sm_old.get('is_bank')}  auto_score_old={sm_old.get('auto_score')}")

        is_data      = cached.get("is_data") or []
        bs_data      = cached.get("bs_data") or []
        cf_data      = cached.get("cf_data") or []
        years        = cached.get("years") or []
        profile      = cached.get("profile") or {}
        wacc_val     = cached.get("wacc_val")
        dcf_prices   = cached.get("dcf_prices") or {}
        analyst_ests = cached.get("analyst_ests") or []

        csv_row    = csv_rows.get(ticker, {})
        bc_manual  = csv_row.get("Manual_Clarity") or None
        ltp_manual = csv_row.get("Manual_LTP") or None

        current_price = float(profile.get("price") or 0) or None
        market_cap    = float(profile.get("mktCap") or profile.get("marketCap") or 0) or None

        # Re-run scorecard with cached financials + profile fallback
        wb = Workbook()
        try:
            _, sm_new = mdl.build_scorecard(
                wb, ticker, is_data, bs_data, cf_data, years,
                biz_clarity=bc_manual or None,
                ltp=ltp_manual or None,
                dcf_gg_price=dcf_prices.get("gg_price"),
                evs_regime=bool(dcf_prices.get("evs_regime")),
                bank_credit=None,   # not cached; fine — bank_credit affects ratios tab only
                analyst_ests=analyst_ests,
                profile=profile,    # ← the fix: prevents is_bank=False from silent API failure
            )
        except Exception as e:
            print(f"  FAIL build_scorecard: {e}")
            import traceback; traceback.print_exc()
            continue

        print(f"  is_bank_new={sm_new.get('is_bank')}  "
              f"auto_score={sm_new.get('auto_score')}  "
              f"floor_cap={sm_new.get('floor_cap')}  "
              f"EBITInt_w={sm_new.get('weights',{}).get('EBITInt')}  "
              f"tier_ebit_int={sm_new.get('tier_ebit_int')}  "
              f"tier_leverage={sm_new.get('tier_leverage')}")

        # Compute adj_score
        auto_score     = sm_new.get("auto_score") or 0
        auto_score_raw = sm_new.get("auto_score_raw") or 0
        _W   = sm_new.get("weights") or {}
        bc_pts  = TIER_PTS.get(bc_manual,  0) * float(_W.get("BC",  2.5))  / 10
        ltp_pts = TIER_PTS.get(ltp_manual, 0) * float(_W.get("LTP", 10.0)) / 10
        adj_score = round((auto_score_raw + bc_pts + ltp_pts) / 10, 1)
        floor_cap = sm_new.get("floor_cap")
        if floor_cap is not None:
            adj_score = min(adj_score, float(floor_cap))

        display_score = adj_score if (bc_manual or ltp_manual) else auto_score
        print(f"  adj_score={adj_score}  display={display_score}")

        # Persist updated scorecard_metrics to cache
        save_ticker_data(ticker, is_data, bs_data, cf_data, profile, years,
                         wacc_val, dcf_prices, sm_new, analyst_ests)

        # Re-render HTML report
        try:
            report_data = build_report_data(
                ticker=ticker, profile=profile,
                is_data=is_data, bs_data=bs_data, cf_data=cf_data, years=years,
                wacc_val=wacc_val, dcf_prices=dcf_prices,
                scorecard_metrics=sm_new, manual_rating=None,
                current_price=current_price, market_cap=market_cap,
                biz_clarity=bc_manual, ltp=ltp_manual,
                adj_score=adj_score if (bc_manual or ltp_manual) else None,
                analyst_ests=analyst_ests,
            )
            html = render_html_report(report_data)
            out_path = os.path.join(RPT_DIR, f"{ticker}_report.html")
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(html)
            print(f"  HTML re-rendered OK  ({len(html):,} bytes)")
        except Exception as e:
            print(f"  FAIL render: {e}")
            import traceback; traceback.print_exc()

        # Update outputs.csv
        _write_score(ticker, display_score, sm_new, dcf_prices,
                     current_price, market_cap, is_data, cf_data,
                     bc_manual, ltp_manual)
        print(f"  outputs.csv updated  Auto_Score={display_score}")

    print("\nDone.")

if __name__ == "__main__":
    main()
