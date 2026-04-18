"""
backfill_heatmap.py
-------------------
Reads existing local Excel models and writes rows to outputs.csv on GitHub.
No FMP API calls — uses only the already-generated .xlsx files.

Run: python backfill_heatmap.py

Requires:
  - GITHUB_TOKEN env variable (or paste directly below)
  - pip install openpyxl requests
"""

import base64
import datetime
import glob
import os
import re
import sys
import requests
import openpyxl

# ── Config ────────────────────────────────────────────────────────────────────
GITHUB_TOKEN  = os.environ.get("GITHUB_TOKEN", "")   # or paste token here
GITHUB_REPO   = "jaysang2908/Investment-Automation"
GITHUB_BRANCH = "main"
FOLDER        = os.path.dirname(os.path.abspath(__file__))

CSV_HEADER = (
    "Ticker,Price,MktCap_B,ROIC,Rev_CAGR,FCF_NI,D_EBITDA,"
    "PE_Current,PE_5yr,PFCF_Current,PFCF_5yr,"
    "Auto_Score,Floor_Cap,Date\n"
)

# Scorecard criteria: (label_substring, weight)
CRITERIA_WEIGHTS = {
    "Moat Profile":              10.0,
    "Management":                 7.5,
    "Revenue 3yr CAGR":          10.0,
    "Cash Quality":              10.0,
    "Capital Returns":            5.0,
    "ROIC":                       7.5,
    "Credit Risk":                5.0,
    "Interest Cover":             7.5,
    "Execution Risk":             5.0,
    "Valuation vs Median  (P/E)":10.0,
    "Valuation vs Median  (P/FCF)":10.0,
}
TIER_VAL = {"HIGH": 10, "MOD-HIGH": 7, "MOD-LOW": 3, "LOW": 0}

# ── GitHub helpers ────────────────────────────────────────────────────────────
GH_HEADERS = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept":        "application/vnd.github.v3+json",
}
GH_API = f"https://api.github.com/repos/{GITHUB_REPO}/contents/outputs.csv"


def _read_csv():
    r = requests.get(GH_API, headers=GH_HEADERS, params={"ref": GITHUB_BRANCH}, timeout=8)
    if r.status_code == 200:
        info = r.json()
        return info["sha"], base64.b64decode(info["content"]).decode()
    return None, CSV_HEADER


def _write_csv(sha, content):
    payload = {
        "message": "Backfill heatmap from local Excel models",
        "branch":  GITHUB_BRANCH,
        "content": base64.b64encode(content.encode()).decode(),
    }
    if sha:
        payload["sha"] = sha
    r = requests.put(GH_API, headers=GH_HEADERS, json=payload, timeout=15)
    if r.status_code not in (200, 201):
        print(f"  ERROR: {r.status_code} — {r.json().get('message','')}")
    else:
        print("  Written to GitHub.")


# ── Excel parsing helpers ─────────────────────────────────────────────────────
def _pct(s):
    """Parse '37.5%' → 0.375, or None."""
    if not s:
        return None
    m = re.search(r"([\d.]+)%", str(s))
    return round(float(m.group(1)) / 100, 4) if m else None


def _num(s):
    """Parse first float from a string, or None."""
    if not s:
        return None
    m = re.search(r"([\d.]+)", str(s))
    return float(m.group(1)) if m else None


def _parse_de(s):
    """Parse D/EBITDA note: '2.1x' → 2.1, 'Net cash ...' → 0.0"""
    s = str(s or "")
    if "net cash" in s.lower():
        return 0.0
    m = re.search(r"([\d.]+)x", s)
    return float(m.group(1)) if m else None


def _parse_val(s):
    """Parse 'Current 51.7x  |  5yr avg 44.9x ...' → (51.7, 44.9)"""
    s = str(s or "")
    cur  = re.search(r"[Cc]urrent\s+([\d.]+)x", s)
    avg  = re.search(r"5yr\s+avg\s+([\d.]+)x", s)
    return (
        float(cur.group(1)) if cur else None,
        float(avg.group(1)) if avg else None,
    )


def _parse_fcf_ni(s):
    """Parse FCF/NI note: '97%' or 'FCF/NI 97% ...' → 0.97"""
    return _pct(s)


def parse_excel(path):
    """Extract all heatmap fields from one Excel workbook."""
    wb = openpyxl.load_workbook(path, data_only=True)
    result = {}

    # ── Scorecard tab ─────────────────────────────────────────────────────────
    if "Scorecard" not in wb.sheetnames:
        return None
    sc = wb["Scorecard"]

    # Scan rows: build {stripped_label: (col_D_value, col_E_tier)}
    label_map = {}
    floor_cap = None
    for row in sc.iter_rows(min_col=1, max_col=5, values_only=True):
        a = str(row[0] or "").strip()
        d = str(row[3] or "").strip()
        e = str(row[4] or "").strip() if row[4] else None
        if a:
            label_map[a] = (d, e)
        # Detect floor gate row
        if "HARD FLOOR GATE" in a.upper() and "CAPPED AT" in a.upper():
            m = re.search(r"capped at\s+(\d+)", a, re.IGNORECASE)
            if m:
                floor_cap = int(m.group(1))

    def _get(label_substr):
        for k, v in label_map.items():
            if label_substr.lower() in k.lower():
                return v
        return ("", None)

    # Parse KPIs
    result["roic"]     = _pct(_get("ROIC")[0])
    result["rev_cagr"] = _pct(_get("Revenue 3yr CAGR")[0])
    result["fcf_ni"]   = _parse_fcf_ni(_get("Cash Quality")[0])
    result["d_ebitda"] = _parse_de(_get("Credit Risk")[0])

    pe_note   = _get("Valuation vs Median  (P/E)")[0]
    pfcf_note = _get("Valuation vs Median  (P/FCF)")[0]
    result["pe_current"],   result["pe_5yr_avg"]   = _parse_val(pe_note)
    result["pfcf_current"], result["pfcf_5yr_avg"] = _parse_val(pfcf_note)

    # Compute auto_score
    scored = []
    for label_substr, weight in CRITERIA_WEIGHTS.items():
        _, tier = _get(label_substr)
        if tier and tier.upper() in TIER_VAL:
            scored.append((TIER_VAL[tier.upper()], weight))

    if scored:
        auto_score = round(sum((s / 10) * w for s, w in scored), 1)
        if floor_cap is not None:
            auto_score = min(auto_score, floor_cap)
        result["auto_score"] = auto_score
    else:
        result["auto_score"] = None

    result["floor_cap"] = floor_cap

    # ── DCF tab — price and shares ────────────────────────────────────────────
    price  = None
    shares = None
    if "DCF" in wb.sheetnames:
        dcf = wb["DCF"]
        for row in dcf.iter_rows(min_col=1, max_col=2, values_only=True):
            a = str(row[0] or "").strip()
            b = row[1]
            if "current market price" in a.lower() and b is not None:
                try:
                    price = float(b)
                except (TypeError, ValueError):
                    pass
            if "shares outstanding" in a.lower() and "diluted" in a.lower() and b is not None:
                try:
                    shares = float(b)   # in millions
                except (TypeError, ValueError):
                    pass

    result["price"]   = round(price, 2) if price else None
    mkt_cap_b = round(price * shares / 1000, 2) if price and shares else None
    result["mkt_cap_b"] = mkt_cap_b

    return result


def _f(v, dp=4):
    return "" if v is None else f"{v:.{dp}f}"


# ── Main ──────────────────────────────────────────────────────────────────────
def run():
    if not GITHUB_TOKEN:
        sys.exit("ERROR: GITHUB_TOKEN not set.")

    # Find all model files
    files = sorted(glob.glob(os.path.join(FOLDER, "*_FinancialModel_*.xlsx")))
    if not files:
        sys.exit("No *_FinancialModel_*.xlsx files found.")

    print(f"Found {len(files)} Excel files.")
    print("Reading outputs.csv from GitHub...")
    sha, content = _read_csv()

    existing = set()
    for line in content.splitlines()[1:]:
        if line.strip():
            existing.add(line.split(",")[0].strip())
    print(f"Already present: {sorted(existing) or 'none'}\n")

    today      = datetime.date.today().isoformat()
    rows_added = 0

    for path in files:
        fname  = os.path.basename(path)
        ticker = fname.split("_")[0]

        if ticker in existing:
            print(f"  Skipping {ticker} — already in CSV")
            continue

        print(f"Processing {ticker} ({fname})...")
        try:
            m = parse_excel(path)
            if m is None:
                print(f"  No Scorecard tab — skipping.")
                continue

            row = ",".join([
                ticker,
                _f(m.get("price"),       2),
                _f(m.get("mkt_cap_b"),   2),
                _f(m.get("roic")),
                _f(m.get("rev_cagr")),
                _f(m.get("fcf_ni")),
                _f(m.get("d_ebitda"),    2),
                _f(m.get("pe_current"),  1),
                _f(m.get("pe_5yr_avg"),  1),
                _f(m.get("pfcf_current"),1),
                _f(m.get("pfcf_5yr_avg"),1),
                "" if m.get("auto_score") is None else str(m["auto_score"]),
                "" if m.get("floor_cap")  is None else str(m["floor_cap"]),
                today,
            ]) + "\n"

            content += row
            rows_added += 1
            print(f"  auto_score={m.get('auto_score')}  price={m.get('price')}  "
                  f"mkt_cap={m.get('mkt_cap_b')}B  roic={m.get('roic')}")

        except Exception as e:
            import traceback
            print(f"  ERROR: {e}")
            traceback.print_exc()

    if rows_added == 0:
        print("\nNo new rows to write.")
        return

    print(f"\nWriting {rows_added} new row(s) to GitHub...")
    _write_csv(sha, content)
    print("Done.")


if __name__ == "__main__":
    run()
