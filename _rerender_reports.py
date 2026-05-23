"""
_rerender_reports.py
Re-renders every HTML report from the cached static/data/*.json files.
No FMP API calls — reads all financial data from disk.
Run after changing report_bridge.py or Report_Template.html.

Score integrity rule: outputs.csv Auto_Score is the authoritative displayed
score. This script injects it (after normalising old raw-scale values) into
scorecard_metrics so the report hero ALWAYS matches the dashboard.
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


def _normalise_sm(sm_orig, csv_auto, csv_cap):
    """Return a patched copy of scorecard_metrics with normalised score fields.

    Old cached data stored auto_score on the raw 0-87.5 scale; new data uses
    normalised 0-10.  Detection: if auto_score > 10 it is on the raw scale.
    The CSV Auto_Score is always the normalised display value and is used as
    the authoritative override.
    """
    sm = dict(sm_orig)

    # ── auto_score ────────────────────────────────────────────────────────────
    cached_auto = sm.get("auto_score")
    if cached_auto is not None and cached_auto > 10:
        # Old raw format — normalise and stash raw value for adj_score maths
        raw_val = cached_auto
        sm["auto_score"]     = round(raw_val / 87.5 * 10, 1)
        if sm.get("auto_score_raw") is None:
            sm["auto_score_raw"] = raw_val
    # Always override with CSV value — it is the authoritative display score
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
            csv_auto = float(csv_row.get("Auto_Score") or 0) or None
        except Exception:
            csv_auto = None
        try:
            csv_cap = float(csv_row.get("Floor_Cap") or 0) or None
        except Exception:
            csv_cap = None

        sm_raw = cached.get("scorecard_metrics") or {}
        sm     = _normalise_sm(sm_raw, csv_auto, csv_cap)

        # ── display score: CSV is authoritative ──────────────────────────────
        # When quals are present, the CSV Auto_Score already stores the correct
        # adj_score from the original run (Rule 10).  Use it directly rather
        # than recomputing from cached raw values, which may differ by one
        # rounding step on old-format caches (auto_score_raw=None).
        adj_score = csv_auto if (bc_manual or ltp_manual) else None

        try:
            profile      = cached.get("profile") or {}
            is_data      = cached.get("is_data") or []
            bs_data      = cached.get("bs_data") or []
            cf_data      = cached.get("cf_data") or []
            years        = cached.get("years") or []
            wacc_val     = cached.get("wacc_val")
            dcf_prices   = cached.get("dcf_prices") or {}
            analyst_ests = cached.get("analyst_ests") or []
            # Reuse cached price_history and consensus_pt so re-render makes
            # ZERO FMP calls (previously every re-render burned 2 per ticker).
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
            displayed = adj_score if adj_score is not None else csv_auto
            print(f"  {ticker:<6}  OK   display={displayed}  ({len(html):,} bytes)")
            ok += 1
        except Exception as e:
            import traceback
            print(f"  {ticker:<6}  FAIL  {e}")
            traceback.print_exc()
            fail += 1

    print(f"\n{ok} rendered OK  |  {fail} failed")


if __name__ == "__main__":
    main()
