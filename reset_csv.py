import base64, requests, os

token = os.environ["GITHUB_TOKEN"]
headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
api = "https://api.github.com/repos/jaysang2908/Investment-Automation/contents/outputs.csv"

r = requests.get(api, headers=headers, params={"ref": "main"})
sha = r.json()["sha"]

new = "Ticker,Price,MktCap_B,ROIC,Rev_CAGR,FCF_NI,D_EBITDA,PE_Current,PE_5yr,PFCF_Current,PFCF_5yr,Auto_Score,Floor_Cap,Date\n"
payload = {
    "message": "Reset outputs.csv with correct header",
    "branch": "main",
    "sha": sha,
    "content": base64.b64encode(new.encode()).decode(),
}
print(requests.put(api, headers=headers, json=payload).status_code)
