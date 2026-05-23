"""
_smoke_test_muff.py — Smoke test the upgraded MUFF technical analysis pipeline.

Runs the full speculative engine + report bridge on mock data, prints scorecard
output, verifies all template variables resolve, and writes the rendered HTML.
"""
import io
import os
import re
import sys
import traceback

# Force UTF-8 stdout so the ↑/↓/● glyphs don't blow up on the Windows cp1252 console.
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import speculative_engine as eng
import speculative_report_bridge as bridge


def run_mock_test():
    print("=" * 72)
    print("MOCK DATA PIPELINE TEST")
    print("=" * 72)
    data = eng._mock_data("MOCK")
    sc = eng.build_speculative_scorecard(
        data,
        narrative_theme="AI / Machine Learning",
        narrative_strength="HIGH",
        catalyst_desc="Q3 2026 product launch with $500M contract win",
        catalyst_timing="near",
    )
    sn = eng.build_scenario_model(data, sc, hold_months=6)

    print(f"\nTotal score: {sc['total_score']}/120  |  Verdict: {sc['verdict']}")
    print(f"\nIndividual signals:")
    for key, s in sc["scores"].items():
        print(f"  {key:<20}  {s['tier']:<5}  {s['pts']:>2} pts")

    print(f"\n--- Technical & Parabolic signal detail ---")
    ts = sc["scores"]["technical_setup"]
    print(f"  tier: {ts['tier']}  pts: {ts['pts']}")
    print(f"  note: {ts['note']}")

    rd = bridge.build_speculative_report_data("MOCK", data, sc, sn)
    print(f"\nPrecursor dashboard: {rd['PRECURSOR_FIRING']}/8 firing")
    print(f"Precursor verdict:   {rd['PRECURSOR_VERDICT']}")
    print(f"Precursor cls:       {rd['PRECURSOR_VERDICT_CLS']}")

    html = bridge.render_speculative_report(rd)
    unresolved = re.findall(r"\{\{[A-Z_]+\}\}", html)
    print(f"\nUnresolved template vars: {len(unresolved)}")
    if unresolved:
        print(f"  first 10: {sorted(set(unresolved))[:10]}")

    out_path = "static/reports/MOCK_speculative_report.html"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\nWrote: {out_path}  ({len(html):,} bytes)")

    # Verify the dashboard HTML is present
    if "precursor-grid" in html and "precursor-tile" in html:
        print("Precursor grid HTML present in output.")
    else:
        print("WARN: precursor grid HTML missing from output!")


def run_live_test(ticker: str = "AMD"):
    """Try a live FMP + Polygon fetch to verify the new TA fields populate."""
    print("\n" + "=" * 72)
    print(f"LIVE DATA TEST — {ticker}")
    print("=" * 72)
    api_key  = os.environ.get("FMP_API_KEY", "")
    poly_key = os.environ.get("POLYGON_API_KEY", "")
    if not api_key or not poly_key:
        print("Skipping live test — FMP_API_KEY or POLYGON_API_KEY env vars not set.")
        return

    try:
        data = eng.fetch_speculative_data(ticker, api_key, poly_key, mock=False)
        print(f"\nKey technical fields for {ticker}:")
        fields = [
            "current_price", "rsi_14", "macd_hist", "ema_21", "ema_50",
            "high_52w", "pct_from_52w_high",
            "bb_squeeze", "bb_expansion", "bb_width_pct",
            "obv_rising", "obv_divergence", "obv_slope_pct",
            "acc_dist_tier", "up_vol_share",
            "roc_20d", "roc_accelerating",
            "atr_14", "atr_pct_price",
        ]
        for f in fields:
            v = data.get(f)
            if isinstance(v, float):
                print(f"  {f:<22} = {v:.4f}")
            else:
                print(f"  {f:<22} = {v}")
    except Exception as e:
        print(f"Live test failed: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    try:
        run_mock_test()
    except Exception as e:
        print(f"\nMOCK TEST FAILED: {e}")
        traceback.print_exc()
        sys.exit(1)

    # Live test is optional — will skip silently if no keys
    run_live_test()
