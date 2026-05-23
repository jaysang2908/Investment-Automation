import time, requests

BASE   = "https://investment-automation.onrender.com"
TICKERS = ["COST", "JNJ", "JPM", "BAC"]
DELAY  = 70

before = {"COST": 7.6, "JNJ": 7.9, "JPM": 5.9, "BAC": 5.0}
results = {}

for i, ticker in enumerate(TICKERS, 1):
    print(f"[{i}/{len(TICKERS)}] {ticker}...", flush=True)
    try:
        r = requests.post(f"{BASE}/generate", json={"ticker": ticker}, timeout=150)
        if r.status_code == 200:
            d = r.json()
            score = d.get("auto_score", "?")
            print(f"  OK  score={score}  report_id={d.get('report_id','?')}")
            results[ticker] = score
        else:
            print(f"  FAIL HTTP {r.status_code}: {r.text[:150]}")
            results[ticker] = None
    except Exception as e:
        print(f"  ERROR: {e}")
        results[ticker] = None

    if i < len(TICKERS):
        print(f"  Sleeping {DELAY}s...", flush=True)
        time.sleep(DELAY)

print()
print("=" * 60)
print(f"  {'Ticker':<6}  {'Before':>7}  {'After':>7}  {'Delta':>7}")
print(f"  {'-'*40}")
for t in TICKERS:
    b = before[t]
    a = results.get(t)
    if a is not None:
        try:
            delta = float(a) - b
            print(f"  {t:<6}  {b:>7.1f}  {float(a):>7.1f}  {delta:>+7.1f}")
        except Exception:
            print(f"  {t:<6}  {b:>7.1f}  {str(a):>7}  {'N/A':>7}")
    else:
        print(f"  {t:<6}  {b:>7.1f}  {'FAIL':>7}  {'N/A':>7}")
print("=" * 60)
