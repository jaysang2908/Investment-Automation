"""Test harness — show before/after scoring impact of #1-#9 changes.

Loads cached financial data per ticker, calls build_scorecard (with light FMP
calls for ratios), then computes both:
  • OLD score: discrete TIER_PTS lookup (10/7/3/0) with default weights
  • NEW score: continuous score with regime-aware weights

Limitations vs production runs:
  • Cached data lacks weightedAverageShsOut / goodwill / dividendsPaid /
    commonStockRepurchased — share-dilution / goodwill indicators in #6 and
    capital-returns in #5 will show degraded values for these tickers.
  • Each ticker triggers ~6 FMP calls (ratios + profile + 4 peer ratios).
"""
import json, sys, copy, os
sys.path.insert(0, os.path.dirname(__file__))

from openpyxl import Workbook
import fmp_3statementv6 as mdl

OLD_TIER_PTS = {"HIGH": 10, "MOD-HIGH": 7, "MOD-LOW": 3, "LOW": 0}
OLD_W = {"BC": 2.5, "Moat": 10.0, "LTP": 10.0, "Mgmt": 7.5,
         "RevCAGR": 10.0, "FCFNI": 10.0, "CapRet": 5.0, "ROIC": 7.5,
         "Lev": 5.0, "EBITInt": 7.5, "Exec": 5.0,
         "PE": 10.0, "PFCF": 10.0}

CRIT_KEYS = [
    ("Moat",     "tier_moat",     "score_moat"),
    ("Mgmt",     "tier_mgmt",     "score_mgmt"),
    ("RevCAGR",  "tier_rev_cagr", "score_rev_cagr"),
    ("FCFNI",    "tier_fcf_ni",   "score_fcf_ni"),
    ("CapRet",   "tier_cap_ret",  "score_cap_ret"),
    ("ROIC",     "tier_roic",     "score_roic"),
    ("Lev",      "tier_leverage", "score_leverage"),
    ("EBITInt",  "tier_ebit_int", "score_ebit_int"),
    ("Exec",     "tier_exec",     "score_exec"),
    ("PE",       "tier_pe",       "score_pe"),
    ("PFCF",     "tier_pfcf",     "score_pfcf"),
]

def load_cached(ticker):
    p = os.path.join(os.path.dirname(__file__), "static", "data", f"{ticker}_data.json")
    return json.load(open(p))

def run_one(ticker):
    print(f"\n{'='*78}\nTICKER: {ticker}\n{'='*78}")
    d = load_cached(ticker)
    is_data, bs_data, cf_data = d["is_data"], d["bs_data"], d["cf_data"]
    years   = d["years"]
    dcf_gg  = (d.get("dcf_prices") or {}).get("gg_price")
    evs     = bool((d.get("dcf_prices") or {}).get("evs_regime"))

    wb = Workbook()
    _, m = mdl.build_scorecard(
        wb, ticker, is_data, bs_data, cf_data, years,
        dcf_gg_price=dcf_gg, evs_regime=evs,
    )

    new_w = m.get("weights") or OLD_W
    print(f"\n  Sector bucket: {m.get('sector_bucket')}  is_bank={m.get('is_bank')}  evs={m.get('evs_regime')}")
    print(f"  Floor cap (soft, #7): {m.get('floor_cap')}")
    print()
    print(f"  {'Criterion':<10} {'Tier':<10} | OLD pts (TIER_PTS)  | NEW score (continuous)  | D score | Weight (old->new)")
    print(f"  {'-'*100}")
    old_total = 0.0
    new_total = 0.0
    for label, tk, sk in CRIT_KEYS:
        tier   = m.get(tk)
        score  = m.get(sk)
        old_pts = OLD_TIER_PTS.get(tier or "", 0) if tier else 0
        old_w   = OLD_W[label]
        new_w_v = new_w[label]
        # Old total uses default weights; new uses regime-aware
        old_contrib = old_pts / 10.0 * old_w
        new_contrib = (score or 0) / 10.0 * new_w_v if (score is not None and new_w_v > 0) else 0
        old_total += old_contrib
        new_total += new_contrib
        d_score = (score or 0) - old_pts if score is not None else 0
        tier_str = tier or "N/A"
        score_str = f"{score:.2f}" if score is not None else "N/A"
        print(f"  {label:<10} {tier_str:<10} | {old_pts:>5}  x{old_w:>5.1f} = {old_contrib:>5.2f} | "
              f"{score_str:>5}  x{new_w_v:>5.1f} = {new_contrib:>5.2f} | {d_score:+6.2f} | "
              f"{old_w:>5.1f} -> {new_w_v:>5.1f}")

    print(f"  {'-'*100}")
    print(f"  {'TOTAL (auto-only, max 87.5)':<35} OLD: {old_total:>6.2f} / 87.5  ({old_total/87.5*100:.1f}%)")
    print(f"  {'engine auto_score':<35}                NEW: {m.get('auto_score'):>6.2f} / 87.5  ({(m.get('auto_score') or 0)/87.5*100:.1f}%)")
    print(f"  Δ total: {(m.get('auto_score') or 0) - old_total:+.2f} pts")
    if m.get("floor_cap"):
        print(f"  Soft cap (#7) applied: {m['floor_cap']:.1f} → final auto_score capped at min(raw, cap)")

    # Show key new annotations
    print(f"\n  Key notes (new annotations):")
    for label, tk, sk in CRIT_KEYS:
        nk = tk.replace("tier_", "")
        # Look for the corresponding note key (built from CRITERIA tuple, not metrics)
    # Just print the criteria notes from tier→note paths if they exist
    print(f"    [Note text only available via building Excel sheet — see Excel file for full notes]")
    return m

def main():
    print("Before/after scoring impact — Investment-Automation scorecard #1-#9")
    print("Engine: continuous scoring + through-cycle + QoE + smarter cap returns")
    print("        + mgmt indicators + soft cap + DCF anchor + regime weights")
    for t in ["AAPL", "F", "JPM"]:
        try:
            run_one(t)
        except Exception as e:
            print(f"\n!! {t} failed: {type(e).__name__}: {e}")
            import traceback; traceback.print_exc()

if __name__ == "__main__":
    main()
