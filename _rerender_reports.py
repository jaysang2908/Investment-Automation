"""
_rerender_reports.py
Re-renders every HTML report from the cached static/data/*.json files.
No FMP API calls — reads all financial data from disk.
Run after changing report_bridge.py or Report_Template.html.

Score integrity rule: outputs.csv Auto_Score is the authoritative displayed
score. This script recomputes adj_score from cached raw values (including
provisional LTP contribution) and writes it back to outputs.csv so the
dashboard and report hero always stay in sync.
"""
import csv, io, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data_store import load_ticker_data
from report_bridge import build_report_data, render_html_report
import csv_schema as _schema

_ROOT    = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(_ROOT, "outputs.csv")
RPT_DIR  = os.path.join(_ROOT, "static", "reports")

TIER_PTS = {"HIGH": 10, "MOD-HIGH": 7, "MOD": 7, "MOD-LOW": 3, "LOW": 0}

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


def _write_csv_score(ticker, new_score, csv_rows):
    """Patch Auto_Score for ticker in the CSV rows dict and write to disk."""
    if ticker not in csv_rows:
        return
    csv_rows[ticker]["Auto_Score"] = str(new_score)
    if not os.path.exists(CSV_PATH):
        return
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        content = _schema.migrate(f.read())
    reader  = csv.DictReader(io.StringIO(content))
    fields  = reader.fieldnames or []
    all_rows = list(reader)
    for row in all_rows:
        if row.get("Ticker", "").strip().upper() == ticker:
            row["Auto_Score"] = str(new_score)
    with open(CSV_PATH, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(all_rows)


def _provisional_ltp_tier(sm, is_data):
    """Derive provisional LTP tier from cached scorecard signals.
    Same logic as build_scorecard() and report_bridge.py legacy fallback.
    Returns tier string (HIGH/MOD-HIGH/MOD-LOW) or None.
    """
    if sm.get("ltp_provisional_tier"):
        return sm["ltp_provisional_tier"]
    bucket  = sm.get("sector_bucket", "")
    cagr    = sm.get("rev_cagr") or 0.0
    evs     = bool(sm.get("evs_regime"))
    gm = 0.0
    if is_data:
        _is0 = is_data[-1]
        _rev = _is0.get("revenue") or 0
        _gp  = _is0.get("grossProfit") or 0
        gm   = _gp / _rev if _rev > 0 else 0.0
    # Banks always MOD-LOW — FCF/EBIT noise can spuriously trigger EVS regime
    if bucket == "bank":
        return "MOD-LOW"
    if evs:
        return "HIGH"
    if bucket == "tech_growth":
        return "HIGH" if (cagr > 0.25 and gm > 0.60) else "MOD-HIGH"
    if bucket == "stable_compounder":
        return "MOD-HIGH" if ((gm > 0.50 and cagr > 0.07) or gm > 0.38) else "MOD-LOW"
    return "MOD-LOW"


def _compute_adj_score(sm, is_data, bc_manual, ltp_manual):
    """Compute adj_score incorporating provisional LTP when no user input.

    When the user has supplied ltp_manual: use their tier.
    Otherwise: use engine-provisional or bridge-derived provisional tier.
    Returns (adj_score_0_10, ltp_tier_used, provisional_flag).
    """
    auto_raw  = sm.get("auto_score_raw")
    if auto_raw is None:
        # Legacy cache: auto_score > 10 means raw scale
        cached_auto = sm.get("auto_score") or 0
        auto_raw = cached_auto * 87.5 / 10.0 if cached_auto <= 10 else cached_auto
    auto_raw  = float(auto_raw or 0)
    W         = sm.get("weights") or {}
    w_bc      = float(W.get("BC",  2.5))
    w_ltp     = float(W.get("LTP", 10.0))
    floor_cap = sm.get("floor_cap")

    bc_pts = TIER_PTS.get((bc_manual or "").strip().upper(), 0) * w_bc / 10

    # LTP: user-supplied takes precedence; fall back to provisional
    ltp_tier      = (ltp_manual or "").strip().upper() or None
    provisional   = False
    if not ltp_tier:
        ltp_tier    = _provisional_ltp_tier(sm, is_data)
        provisional = True

    # Normalise: engine uses 4-tier; old user dropdowns used HIGH/MOD/LOW
    _ltp_map = {"HIGH": "HIGH", "MOD-HIGH": "MOD-HIGH", "MOD": "MOD-HIGH",
                "MOD-LOW": "MOD-LOW", "LOW": "LOW"}
    ltp_tier_norm = _ltp_map.get(ltp_tier, "MOD-LOW") if ltp_tier else "MOD-LOW"
    ltp_pts = TIER_PTS.get(ltp_tier_norm, 0) * w_ltp / 10

    adj_raw   = auto_raw + bc_pts + ltp_pts
    adj_score = round(min(10.0, adj_raw / 87.5 * 10), 1)
    if floor_cap is not None:
        adj_score = min(adj_score, float(floor_cap))
    return adj_score, ltp_tier_norm, provisional


def _normalise_sm(sm_orig, csv_auto, csv_cap):
    """Return a patched copy of scorecard_metrics with normalised score fields."""
    sm = dict(sm_orig)

    # ── auto_score ────────────────────────────────────────────────────────────
    cached_auto = sm.get("auto_score")
    if cached_auto is not None and cached_auto > 10:
        raw_val = cached_auto
        sm["auto_score"]     = round(raw_val / 87.5 * 10, 1)
        if sm.get("auto_score_raw") is None:
            sm["auto_score_raw"] = raw_val
    # Override with provided value (recomputed adj_score below)
    if csv_auto is not None:
        sm["auto_score"] = csv_auto

    # ── floor_cap ─────────────────────────────────────────────────────────────
    cached_cap = sm.get("floor_cap")
    if cached_cap is not None and cached_cap > 10:
        sm["floor_cap"] = round(cached_cap / 87.5 * 10, 1)
    if csv_cap is not None:
        sm["floor_cap"] = csv_cap

    return sm


def main():
    csv_rows = _read_csv()
    data_dir = os.path.join(_ROOT, "static", "data")
    json_files = [f for f in os.listdir(data_dir) if f.endswith("_data.json")]
    tickers = sorted(f.replace("_data.json", "") for f in json_files)

    print(f"Re-rendering {len(tickers)} reports from cache (no FMP calls)\n")
    ok = 0; fail = 0

    for ticker in tickers:
        cached = load_ticker_data(ticker)
        if not cached:
            print(f"  {ticker:<6}  SKIP  (no cache)")
            continue

        csv_row    = csv_rows.get(ticker, {})
        bc_manual  = csv_row.get("Manual_Clarity") or None
        ltp_manual = csv_row.get("Manual_LTP") or None

        try:
            csv_cap = float(csv_row.get("Floor_Cap") or 0) or None
        except Exception:
            csv_cap = None

        sm_raw  = cached.get("scorecard_metrics") or {}
        is_data = cached.get("is_data") or []

        # Recompute adj_score from raw values + provisional LTP.
        # This ensures the hero score includes the provisional LTP contribution
        # even for cached reports that predate the feature.
        adj_score, ltp_used, prov = _compute_adj_score(sm_raw, is_data, bc_manual, ltp_manual)

        sm = _normalise_sm(sm_raw, adj_score, csv_cap)

        # Write updated score back to outputs.csv so dashboard stays in sync
        _write_csv_score(ticker, adj_score, csv_rows)

        try:
            profile      = cached.get("profile") or {}
            bs_data      = cached.get("bs_data") or []
            cf_data      = cached.get("cf_data") or []
            years        = cached.get("years") or []
            wacc_val     = cached.get("wacc_val")
            dcf_prices   = cached.get("dcf_prices") or {}
            analyst_ests = cached.get("analyst_ests") or []
            price_history = cached.get("price_history")
            consensus_pt  = cached.get("consensus_pt")

            current_price = float(profile.get("price") or 0) or None
            market_cap    = float(profile.get("mktCap") or profile.get("marketCap") or 0) or None

            report_data = build_report_data(
                ticker=ticker, profile=profile,
                is_data=is_data, bs_data=bs_data, cf_data=cf_data, years=years,
                wacc_val=wacc_val, dcf_prices=dcf_prices,
                scorecard_metrics=sm, manual_rating=None,
                current_price=current_price, market_cap=market_cap,
                biz_clarity=bc_manual, ltp=ltp_manual,
                adj_score=adj_score, analyst_ests=analyst_ests,
                price_history=price_history, consensus_pt=consensus_pt,
            )
            html = render_html_report(report_data)
            out_path = os.path.join(RPT_DIR, f"{ticker}_report.html")
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(html)
            prov_tag = f"  LTP={ltp_used}{'★' if prov else ''}" if prov else ""
            print(f"  {ticker:<6}  OK   display={adj_score}{prov_tag}  ({len(html):,} bytes)")
            ok += 1
        except Exception as e:
            import traceback
            print(f"  {ticker:<6}  FAIL  {e}")
            traceback.print_exc()
            fail += 1

    print(f"\n{ok} rendered OK  |  {fail} failed")


if __name__ == "__main__":
    main()
