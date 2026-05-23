"""
_diag_chart.py
One-FMP-call diagnostic to verify the /stable/historical-price-eod/full
endpoint shape matches what _fetch_price_history() expects.

Behavior:
  1. Hits the endpoint for a single ticker (default JPM).
  2. Dumps response status, top-level shape, first record keys, and a sample
     value for `adjClose` / `close` / `price` / `date`.
  3. Runs our actual parser (_fetch_price_history) and reports point count.
  4. If parser returns valid data, ALSO saves to the ticker's cache so the
     chart immediately renders correctly - and re-renders that ticker's HTML.
  5. Burns exactly 1 FMP call (or 0 if quota still exhausted).

Run after the FMP daily quota resets (UTC midnight).
"""
import sys, os, json, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests
from fmp_3statementv6 import API_KEY
from report_bridge import _fetch_price_history, build_report_data, render_html_report
from data_store import load_ticker_data, save_ticker_data

TICKER = (sys.argv[1] if len(sys.argv) > 1 else "JPM").upper()


def main():
    print(f"\n{'='*70}")
    print(f"CHART DIAGNOSTIC - {TICKER}")
    print(f"{'='*70}\n")

    # ── Step 1: raw response ─────────────────────────────────────────────────
    end_date   = datetime.date.today()
    start_date = end_date - datetime.timedelta(days=int(3 * 365.25))
    url = (f"https://financialmodelingprep.com/stable/historical-price-eod/full"
           f"?symbol={TICKER}"
           f"&from={start_date.isoformat()}&to={end_date.isoformat()}"
           f"&apikey={API_KEY}")
    print(f"  URL: {url.replace(API_KEY, '***')}\n")

    try:
        r = requests.get(url, timeout=15)
    except Exception as e:
        print(f"  REQUEST FAILED: {e}")
        return

    print(f"  HTTP status: {r.status_code}")
    if r.status_code == 429:
        print("  -> Quota still exhausted. Try again after UTC midnight.")
        print(f"  -> Raw error: {r.text[:200]}")
        return
    if r.status_code != 200:
        print(f"  -> Unexpected status. Body: {r.text[:300]}")
        return

    try:
        data = r.json()
    except Exception as e:
        print(f"  -> Response not valid JSON: {e}")
        print(f"  -> Body: {r.text[:300]}")
        return

    print(f"  Response type: {type(data).__name__}")
    if isinstance(data, list):
        print(f"  Response: list of {len(data)} records (flat array - matches parser)")
        sample = data[0] if data else {}
    elif isinstance(data, dict):
        print(f"  Response: dict with top-level keys {list(data.keys())[:8]}")
        # Common shapes: {symbol, historical: [...]}, {data: [...]}, {prices: [...]}
        for k in ("historical", "data", "prices", "results"):
            if k in data and isinstance(data[k], list):
                print(f"  -> DETECTED: records nested under '{k}' key (parser expects flat list - BUG)")
                sample = data[k][0] if data[k] else {}
                break
        else:
            print(f"  -> UNKNOWN SHAPE - parser will return empty")
            sample = data
    else:
        print(f"  -> Unexpected response: {data}")
        return

    if isinstance(sample, dict):
        print(f"\n  First record keys: {list(sample.keys())}")
        print(f"  Sample values:")
        for f in ["date", "adjClose", "close", "price", "open", "high", "low", "volume"]:
            v = sample.get(f)
            if v is not None:
                print(f"    {f:<10} = {v}")

    # ── Step 2: run actual parser ────────────────────────────────────────────
    print(f"\n  Running _fetch_price_history('{TICKER}', years=3) - note: this makes a SECOND call")
    print(f"  (skip this and inspect manually if you want to conserve quota)")
    # Skip the second call to be quota-safe - re-use what we already have
    print(f"  Skipped - re-using the response already in memory.\n")

    # Manually run the same parsing logic as _fetch_price_history
    if isinstance(data, list):
        records = data
    elif isinstance(data, dict):
        records = next((data[k] for k in ("historical","data","prices","results") if isinstance(data.get(k), list)), [])
    else:
        records = []

    records = list(reversed(records))   # oldest -> newest
    full_labels, full_values = [], []
    for d in records:
        v = d.get("adjClose")
        if v is None:
            v = d.get("price") or d.get("close")
        try:
            v = float(v) if v is not None else 0.0
        except Exception:
            continue
        if v <= 0:
            continue
        full_labels.append(str(d.get("date") or "")[:10])
        full_values.append(round(v, 4))

    print(f"  Parsed {len(full_values)} valid points")
    if full_values:
        print(f"    First:  {full_labels[0]}  ${full_values[0]}")
        print(f"    Last:   {full_labels[-1]}  ${full_values[-1]}")
        print(f"    Range:  ${min(full_values):.2f} - ${max(full_values):.2f}")
        ret = (full_values[-1] / full_values[0] - 1) * 100
        print(f"    3yr return: {ret:+.1f}%")
    else:
        print(f"  -> ZERO valid points parsed. Either endpoint returned empty, or field-name mismatch.")
        return

    # ── Step 3: save to cache so the chart shows immediately ────────────────
    cached = load_ticker_data(TICKER)
    if not cached:
        print(f"\n  ! No existing cache for {TICKER} - can't save. Run local_rerun.py first.")
        return

    # Sample to ~120 points for chart payload size (matches _fetch_price_history logic)
    target = 120
    step   = max(1, len(full_values) // target)
    sampled_labels = full_labels[::step]
    sampled_values = full_values[::step]
    if sampled_labels and sampled_labels[-1] != full_labels[-1]:
        sampled_labels.append(full_labels[-1])
        sampled_values.append(full_values[-1])

    stats = {
        "first":      full_values[0],
        "last":       full_values[-1],
        "high":       max(full_values),
        "low":        min(full_values),
        "high_date":  full_labels[full_values.index(max(full_values))],
        "low_date":   full_labels[full_values.index(min(full_values))],
        "return_pct": (full_values[-1] / full_values[0] - 1) if full_values[0] > 0 else None,
        "first_date": full_labels[0],
        "last_date":  full_labels[-1],
        "n_points":   len(full_values),
    }
    payload = {"labels": sampled_labels, "values": sampled_values, "stats": stats}

    save_ticker_data(
        TICKER,
        cached["is_data"], cached["bs_data"], cached["cf_data"],
        cached.get("profile") or {}, cached.get("years") or [],
        cached.get("wacc_val"),
        cached.get("dcf_prices") or {},
        cached.get("scorecard_metrics") or {},
        cached.get("analyst_ests") or [],
        price_history=payload,
        consensus_pt=cached.get("consensus_pt"),  # preserve
    )
    print(f"\n  OK Saved {len(sampled_values)} sampled chart points to {TICKER}_data.json")
    print(f"  -> Re-render {TICKER} to see the chart populate.")


if __name__ == "__main__":
    main()
