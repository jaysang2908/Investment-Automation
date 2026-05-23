"""
rerun_tickers.py
Re-runs selected tickers through the live Render app to pick up latest code changes.
Each call regenerates the HTML report and updates outputs.csv (+ pushes to GitHub).

Run: python rerun_tickers.py
"""

import time
import requests

BASE_URL = "https://investment-automation.onrender.com"

TICKERS = [
    "NFLX", "NVDA", "ADBE", "COST", "V", "MSFT",
    "AMD", "AAPL", "JNJ", "META", "TSM", "KO",
]

DELAY_SECONDS = 8  # pause between calls to avoid hammering FMP API


def run():
    total = len(TICKERS)
    passed, failed = [], []

    for i, ticker in enumerate(TICKERS, 1):
        print(f"[{i}/{total}] Generating {ticker}...", end=" ", flush=True)
        try:
            r = requests.post(
                f"{BASE_URL}/generate",
                json={"ticker": ticker},
                timeout=120,
            )
            if r.status_code == 200:
                data = r.json()
                rid = data.get("report_id", "?")
                score = data.get("auto_score", "?")
                print(f"OK — report_id={rid}, score={score}")
                passed.append(ticker)
            else:
                print(f"FAILED — HTTP {r.status_code}: {r.text[:120]}")
                failed.append(ticker)
        except Exception as e:
            print(f"ERROR — {e}")
            failed.append(ticker)

        if i < total:
            time.sleep(DELAY_SECONDS)

    print(f"\n{'='*50}")
    print(f"Done: {len(passed)}/{total} succeeded.")
    if failed:
        print(f"Failed: {', '.join(failed)}")
    else:
        print("All tickers updated successfully.")
    print(f"\nDashboard: {BASE_URL}/dashboard.html")


if __name__ == "__main__":
    run()
