import requests
r = requests.post("https://investment-automation.onrender.com/generate",
                  json={"ticker": "ADBE"}, timeout=150)
print(f"HTTP {r.status_code}")
if r.status_code == 200:
    d = r.json()
    print(f"score={d.get('auto_score')}  report_id={d.get('report_id')}")
else:
    print(r.text[:400])
