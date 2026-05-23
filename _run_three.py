"""Temporary script: re-run AAPL, AMD, V on Render."""
import time, requests

BASE_URL = "https://investment-automation.onrender.com"
TICKERS = ["AAPL", "AMD", "V"]
DELAY = 70

results = {}
for i, t in enumerate(TICKERS, 1):
    print(f"[{i}/3] Running {t}...", flush=True)
    try:
        r = requests.post(f"{BASE_URL}/generate", json={"ticker": t}, timeout=180)
        if r.status_code == 200:
            d = r.json()
            score    = d.get("auto_score")
            price    = d.get("price")
            pt       = d.get("pt")
            gg_up    = d.get("gg_upside")
            em_up    = d.get("em_upside")
            mcap     = d.get("market_cap_b")
            roic     = d.get("roic")
            cagr     = d.get("rev_cagr")
            fcfni    = d.get("fcf_ni")
            de       = d.get("d_ebitda")
            pe       = d.get("pe_current")
            pe5      = d.get("pe_5yr")
            pf       = d.get("pfcf_current")
            pf5      = d.get("pfcf_5yr")
            print(f"  OK  score={score}  price={price}  pt={pt}")
            print(f"  gg_up={gg_up}  em_up={em_up}  mcap={mcap}")
            print(f"  roic={roic}  cagr={cagr}  fcfni={fcfni}  de={de}")
            print(f"  pe={pe}  pe5={pe5}  pf={pf}  pf5={pf5}")
            results[t] = d
        else:
            print(f"  FAILED HTTP {r.status_code}: {r.text[:300]}")
            results[t] = {"error": r.status_code}
    except Exception as e:
        print(f"  ERROR: {e}")
        results[t] = {"error": str(e)}
    if i < len(TICKERS):
        print(f"  Waiting {DELAY}s...", flush=True)
        time.sleep(DELAY)

print("\nDone.")
print("Summary:", {t: results[t].get("auto_score", results[t].get("error")) for t in TICKERS})
