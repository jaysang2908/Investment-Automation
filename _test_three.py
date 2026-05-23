"""Quick regression test for the 3 scorecard changes."""
import time, requests

BASE = "https://investment-automation.onrender.com"
TICKERS = ["NVDA", "NFLX", "META"]
DELAY = 70

before = {"NVDA": 8.7, "NFLX": 9.8, "META": 7.4}
results = {}

for i, ticker in enumerate(TICKERS, 1):
    print(f"[{i}/3] POSTing {ticker}...", flush=True)
    try:
        r = requests.post(f"{BASE}/generate", json={"ticker": ticker}, timeout=150)
        if r.status_code == 200:
            d = r.json()
            score = d.get("auto_score", "?")
            rid   = d.get("report_id", "?")
            print(f"  OK  score={score}  report_id={rid}")
            results[ticker] = score
        else:
            print(f"  FAIL HTTP {r.status_code}: {r.text[:200]}")
            results[ticker] = None
    except Exception as e:
        print(f"  ERROR: {e}")
        results[ticker] = None

    if i < len(TICKERS):
        print(f"  Sleeping {DELAY}s...", flush=True)
        time.sleep(DELAY)

print()
print("=" * 55)
print(f"  {'Ticker':<6}  {'Before':>7}  {'After':>7}  {'Delta':>7}  Notes")
print(f"  {'-'*70}")

notes = {
    "NVDA": "exec-risk fix (CAGR 100%), PEG~0.38 -> +3pt abs-PE",
    "NFLX": "DCF contamination removed",
    "META": "PEG~1.38 -> +1.5pt abs-PE boost",
}

for t in TICKERS:
    b = before[t]
    a = results.get(t)
    if a is not None:
        try:
            delta = float(a) - float(b)
            print(f"  {t:<6}  {b:>7.1f}  {float(a):>7.1f}  {delta:>+7.1f}  {notes[t]}")
        except Exception:
            print(f"  {t:<6}  {b:>7.1f}  {str(a):>7}  {'N/A':>7}  {notes[t]}")
    else:
        print(f"  {t:<6}  {b:>7.1f}  {'FAIL':>7}  {'N/A':>7}  {notes[t]}")

print("=" * 55)
