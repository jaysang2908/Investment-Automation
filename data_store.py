"""
data_store.py
Saves and loads raw FMP financial data per ticker to static/data/TICKER_data.json.
The DCF calculator reads these files so it never needs to re-hit FMP for
tickers that already have reports.
"""
import os, json, datetime

DATA_DIR = os.path.join(os.path.dirname(__file__), "static", "data")

def save_ticker_data(ticker, is_data, bs_data, cf_data, profile, years,
                     wacc_val, dcf_prices, scorecard_metrics, analyst_ests=None,
                     anomalies=None, price_history=None, consensus_pt=None):
    """Save ticker data to disk.

    price_history (optional): dict {"labels": [iso dates], "values": [floats],
                                    "stats": {...}} from _fetch_price_history.
        Persisted so re-renders don't waste a fresh FMP call on every render.
    consensus_pt  (optional): dict {"consensus_pt": float, "consensus_pt_high":
                                    float, "consensus_pt_low": float,
                                    "analyst_count": int} from _fetch_analyst_data.
        Persisted for the same reason.

    Existing fields (price_history, consensus_pt) are preserved across saves
    when the new call doesn't supply them — so partial updates don't blow
    away previously-cached data.
    """
    os.makedirs(DATA_DIR, exist_ok=True)
    path = os.path.join(DATA_DIR, f"{ticker}_data.json")

    # Preserve previously-cached price_history / consensus_pt when caller
    # didn't pass them (e.g. _rescore_banks.py only updates scorecard)
    existing = {}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                existing = json.load(f) or {}
        except Exception:
            existing = {}

    payload = {
        "ticker":            ticker,
        "fetched":           datetime.date.today().isoformat(),
        "profile":           profile,
        "years":             years,
        "is_data":           is_data,
        "bs_data":           bs_data,
        "cf_data":           cf_data,
        "wacc_val":          wacc_val,
        "dcf_prices":        dcf_prices or {},
        "scorecard_metrics": scorecard_metrics or {},
        "analyst_ests":      analyst_ests or [],
        "anomalies":         anomalies or [],
        "price_history":     price_history if price_history is not None else existing.get("price_history"),
        "consensus_pt":      consensus_pt  if consensus_pt  is not None else existing.get("consensus_pt"),
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f)

def load_ticker_data(ticker):
    path = os.path.join(DATA_DIR, f"{ticker}_data.json")
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)
