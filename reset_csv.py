"""Reset outputs.csv to current schema header (wipes all rows)."""
import base64, requests, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import csv_schema as _schema

token = os.environ["GITHUB_TOKEN"]
headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
api = "https://api.github.com/repos/jaysang2908/Investment-Automation/contents/outputs.csv"

r = requests.get(api, headers=headers, params={"ref": "main"})
sha = r.json().get("sha")

payload = {
    "message": "Reset outputs.csv to current schema",
    "branch": "main",
    "content": base64.b64encode(_schema.HEADER.encode()).decode(),
}
if sha:
    payload["sha"] = sha
print(requests.put(api, headers=headers, json=payload).status_code)
