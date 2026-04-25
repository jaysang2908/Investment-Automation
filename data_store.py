"""
data_store.py
Saves and loads raw FMP financial data per ticker to static/data/TICKER_data.json.
The DCF calculator reads these files so it never needs to re-hit FMP for
tickers that already have reports.
"""
import os, json, datetime

DATA_DIR = os.path.join(os.path.dirname(__file__), "static", "data")

def save_ticker_data(ticker, is_data, bs_data, cf_data, profile, years,
                     wacc_val, dcf_prices, scorecard_metrics, analyst_ests=None):
    os.makedirs(DATA_DIR, exist_ok=True)
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
    }
    path = os.path.join(DATA_DIR, f"{ticker}_data.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f)

def load_ticker_data(ticker):
    path = os.path.join(DATA_DIR, f"{ticker}_data.json")
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)
