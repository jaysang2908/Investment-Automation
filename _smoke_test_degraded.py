"""
_smoke_test_degraded.py — Verify graceful degradation when TA data is missing.

Simulates the scenario where Polygon API is unreachable (no OHLCV bars), so all
the new technical indicators end up as None. The pipeline should still produce a
clean report with the precursor dashboard showing "Setup Not Present".
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import speculative_engine as eng
import speculative_report_bridge as bridge

data = eng._mock_data("NODATA")

# Wipe everything that comes from Polygon OHLCV
ta_fields = [
    "bb_squeeze", "bb_expansion", "bb_width_pct", "bb_pos_pct", "bb_upper", "bb_lower", "bb_mid",
    "obv_rising", "obv_divergence", "obv_slope_pct",
    "acc_dist_tier", "up_vol_share",
    "roc_20d", "roc_accelerating",
    "atr_14", "atr_pct_price",
    "macd", "macd_hist", "macd_signal",
    "ema_21", "ema_50",
    "rsi_14",
    "high_52w", "low_52w", "pct_from_52w_high",
    "price_vs_50ma", "price_vs_200ma",
    "vol_10d_avg", "vol_50d_avg",
]
for k in ta_fields:
    data[k] = None

sc = eng.build_speculative_scorecard(data)
ts = sc["scores"]["technical_setup"]
print(f"TA scorer with no data:  tier={ts['tier']}  pts={ts['pts']}")
print(f"  note: {ts['note']}")

sn = eng.build_scenario_model(data, sc, hold_months=6)
rd = bridge.build_speculative_report_data("NODATA", data, sc, sn)
print(f"\nPrecursor firing: {rd['PRECURSOR_FIRING']} / 8")
print(f"Precursor verdict: {rd['PRECURSOR_VERDICT']}")
print(f"Precursor cls: {rd['PRECURSOR_VERDICT_CLS']}")

# Verify the HTML still renders cleanly
html = bridge.render_speculative_report(rd)
import re
unresolved = re.findall(r"\{\{[A-Z_]+\}\}", html)
print(f"\nUnresolved template vars: {len(unresolved)}")
print(f"HTML length: {len(html):,} bytes")
print(f"Precursor grid present: {'precursor-grid' in html}")
print(f"All 8 tiles present: {html.count('precursor-tile') == 8}")
