"""
_batch1_rescore.py — scorecard-ONLY offline re-score (zero FMP calls).

Purpose: propagate scorecard-engine fixes (2026-07 batch 1: ROIC sign bug,
FCF/NI both-negative gap, TTM-first P/E basis, FX-corrected foreign multiples)
to every ticker with a cached JSON — WITHOUT re-running build_dcf. This is the
key difference vs _rescore_offline.py: the full re-score recomputes GG/EM fair
values with offline fallbacks (peer-beta 429 → approximation), silently
degrading live DCF prices. Here dcf_prices are read from cache and left
untouched; only scorecard_metrics and the scorecard CSV columns change.

Updates per ticker:
  - static/data/{T}_data.json → scorecard_metrics replaced
  - outputs.csv → Auto_Score, Floor_Cap, ROIC, FCF_NI, PE_Current, PFCF_Current
    (PE/PFCF rescaled to the CSV's current price, which the daily price cron
    keeps fresh, via ratio = csv_price / cached_profile_price)
  - GG/EM prices, upsides, financial columns: NOT touched.

After running: python _rerender_reports.py && python _score_audit.py

Usage:
    python _batch1_rescore.py            # all tickers with cached JSONs
    python _batch1_rescore.py LCID SNAP  # specific tickers
"""

import builtins
import csv
import io
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook

import fmp_3statementv6 as mdl
import csv_schema as _schema
from _rescore_offline import _inject_ratios_cache

_ROOT     = os.path.dirname(os.path.abspath(__file__))
CSV_PATH  = os.path.join(_ROOT, "outputs.csv")
DATA_DIR  = os.path.join(_ROOT, "static", "data")

TIER_PTS = {"HIGH": 10, "MOD-HIGH": 7, "MOD": 7, "MOD-LOW": 3, "LOW": 0}


def _f(v, dp=4):
    return "" if v is None else f"{v:.{dp}f}"


def _num(v):
    try:
        return float(v) if v not in ("", None) else None
    except ValueError:
        return None


def rescore(ticker: str, row: dict):
    path = os.path.join(DATA_DIR, f"{ticker}_data.json")
    if not os.path.exists(path):
        return None, "no cached JSON"
    with open(path, "r", encoding="utf-8") as f:
        cached = json.load(f)

    is_data = cached.get("is_data") or []
    bs_data = cached.get("bs_data") or []
    cf_data = cached.get("cf_data") or []
    if not (is_data and bs_data and cf_data):
        return None, "incomplete cached financials"

    profile      = cached.get("profile") or {}
    years        = cached.get("years") or []
    analyst_ests = cached.get("analyst_ests") or []
    prev_sm      = cached.get("scorecard_metrics") or {}
    dp           = cached.get("dcf_prices") or {}

    bc_manual  = (row.get("Manual_Clarity") or "").strip() or None
    ltp_manual = (row.get("Manual_LTP") or "").strip() or None

    _inject_ratios_cache(ticker, prev_sm)

    logs = []
    _orig_print = builtins.print
    builtins.print = lambda *a, **k: logs.append(" ".join(str(x) for x in a))
    try:
        bank_credit = mdl.fetch_bank_credit_data(ticker)   # EDGAR, not FMP
        wb = Workbook()
        _, sm = mdl.build_scorecard(
            wb, ticker, is_data, bs_data, cf_data, years,
            biz_clarity=bc_manual, ltp=ltp_manual,
            dcf_gg_price=dp.get("gg_price"),
            evs_regime=bool(dp.get("evs_regime")),
            bank_credit=bank_credit,
            analyst_ests=analyst_ests,
            profile=profile,
            fx_to_usd=dp.get("fx_to_usd"),
        )
    except Exception as e:
        builtins.print = _orig_print
        import traceback
        traceback.print_exc()
        return None, f"engine error: {e}"
    finally:
        builtins.print = _orig_print

    # Display score: quant-only, or qual-adjusted when BC/LTP present (Rule 10)
    auto_score     = sm.get("auto_score") or 0
    auto_score_raw = sm.get("auto_score_raw") or 0
    W       = sm.get("weights") or {}
    bc_pts  = TIER_PTS.get(bc_manual, 0)  * float(W.get("BC", 2.5))   / 10
    ltp_pts = TIER_PTS.get(ltp_manual, 0) * float(W.get("LTP", 10.0)) / 10
    adj = round((auto_score_raw + bc_pts + ltp_pts) / 10, 1)
    fc  = sm.get("floor_cap")
    if fc is not None:
        adj = min(adj, float(fc))
    display = adj if (bc_manual or ltp_manual) else auto_score

    # PE/PFCF rescale to the CSV's (fresh) price basis
    csv_px    = _num(row.get("Price"))
    cached_px = _num(profile.get("price"))
    ratio = (csv_px / cached_px) if (csv_px and cached_px and cached_px > 0) else 1.0

    old = {
        "score":  row.get("Auto_Score"), "roic": row.get("ROIC"),
        "fcfni":  row.get("FCF_NI"),     "pe":   row.get("PE_Current"),
        "pfcf":   row.get("PFCF_Current"), "cap": row.get("Floor_Cap"),
    }
    row["Auto_Score"]   = "" if display is None else str(display)
    row["Floor_Cap"]    = "" if fc is None else str(fc)
    row["ROIC"]         = _f(sm.get("roic"))
    row["FCF_NI"]       = _f(sm.get("fcf_ni"))
    row["PE_Current"]   = _f(sm.get("pe_current") * ratio
                             if sm.get("pe_current") is not None else None, 1)
    row["PFCF_Current"] = _f(sm.get("pfcf_current") * ratio
                             if sm.get("pfcf_current") is not None else None, 1)

    # Persist new scorecard metrics; dcf_prices remain untouched
    cached["scorecard_metrics"] = sm
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cached, f)

    new = {
        "score": row["Auto_Score"], "roic": row["ROIC"],
        "fcfni": row["FCF_NI"],     "pe":   row["PE_Current"],
        "pfcf":  row["PFCF_Current"], "cap": row["Floor_Cap"],
    }
    return (old, new), None


def main():
    only = {t.upper() for t in sys.argv[1:]}

    with open(CSV_PATH, "r", encoding="utf-8") as f:
        content = _schema.migrate(f.read())
    rows = list(csv.DictReader(io.StringIO(content)))
    cols = _schema.COLUMNS

    changed, skipped = [], []
    for row in rows:
        t = (row.get("Ticker") or "").strip().upper()
        if not t or (only and t not in only):
            continue
        res, err = rescore(t, row)
        if err:
            skipped.append((t, err))
            continue
        old, new = res
        flag = "  *" if old["score"] != new["score"] else ""
        print(f"{t:6s} score {old['score']:>4} -> {new['score']:>4}{flag}   "
              f"ROIC {old['roic'][:6]:>6} -> {new['roic'][:6]:>6}   "
              f"FCF/NI {old['fcfni'][:7]:>7} -> {new['fcfni'][:7]:>7}   "
              f"P/E {old['pe']:>6} -> {new['pe']:>6}")
        changed.append(t)

    with open(CSV_PATH, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, lineterminator="\n")
        w.writeheader()
        w.writerows(rows)

    print(f"\nRe-scored {len(changed)} tickers; skipped {len(skipped)}:")
    for t, e in skipped:
        print(f"  {t}: {e}")
    print("\nNext: python _rerender_reports.py && python _score_audit.py")


if __name__ == "__main__":
    main()
