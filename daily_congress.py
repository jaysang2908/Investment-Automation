"""
daily_congress.py — Fetch congressional stock trade disclosures for covered tickers.

Data source: Financial Modeling Prep (FMP) /stable/senate-latest-trading and
/stable/house-latest-trading endpoints. Uses the same FMP_API_KEY already
configured for the main engine.

Filters to covered tickers in outputs.csv. Saves last LOOKBACK_DAYS of trades.
Pushes result to GitHub so it survives Render redeploys.

Runs as a Render cron (04:00 UTC daily).
"""

import os
import csv
import json
import base64
import logging
import datetime
import requests

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s")
log = logging.getLogger(__name__)

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
OUTPUTS_CSV = os.path.join(SCRIPT_DIR, "outputs.csv")
OUTPUT_JSON = os.path.join(SCRIPT_DIR, "static", "data", "congress_cache.json")

FMP_API_KEY   = os.environ.get("FMP_API_KEY", "")
GITHUB_TOKEN  = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO   = os.environ.get("GITHUB_REPO", "jaysang2908/Investment-Automation")
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH", "main")

LOOKBACK_DAYS = 180  # keep 6 months; dashboard shows last 30

FMP_BASE = "https://financialmodelingprep.com"
HEADERS  = {"User-Agent": "Mozilla/5.0 (compatible; investment-research-bot/1.0)"}


# ── GitHub helpers ────────────────────────────────────────────────────────────

def _gh_headers():
    return {"Authorization": f"token {GITHUB_TOKEN}",
            "Accept": "application/vnd.github.v3+json"}


def _push_to_github(local_path: str, repo_path: str, commit_msg: str) -> None:
    if not GITHUB_TOKEN:
        log.info("No GITHUB_TOKEN — skipping push for %s", repo_path)
        return
    try:
        with open(local_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode()
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{repo_path}"
        r = requests.get(url, headers=_gh_headers(), params={"ref": GITHUB_BRANCH}, timeout=8)
        sha = r.json().get("sha") if r.status_code == 200 else None
        payload = {"message": commit_msg, "branch": GITHUB_BRANCH, "content": content_b64}
        if sha:
            payload["sha"] = sha
        requests.put(url, headers=_gh_headers(), json=payload, timeout=15)
        log.info("GitHub push OK — %s", repo_path)
    except Exception as e:
        log.warning("GitHub push failed for %s: %s", repo_path, e)


# ── Data loading ──────────────────────────────────────────────────────────────

def load_covered() -> set:
    covered = set()
    try:
        with open(OUTPUTS_CSV, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                t = (row.get("Ticker") or "").strip().upper()
                if t:
                    covered.add(t)
    except FileNotFoundError:
        log.warning("outputs.csv not found")
    return covered


def _norm_type(raw: str) -> str:
    r = raw.lower().strip()
    if "purchase" in r or "buy" in r:
        return "purchase"
    if "sale" in r or "sell" in r or "sold" in r:
        return "sale"
    if "exchange" in r:
        return "exchange"
    return r


def _norm_date(raw: str) -> str:
    """Normalise various date formats to YYYY-MM-DD."""
    raw = (raw or "").strip()[:10]
    if not raw:
        return ""
    # Already ISO
    if len(raw) == 10 and raw[4] == "-":
        return raw
    # MM/DD/YYYY
    try:
        return datetime.datetime.strptime(raw, "%m/%d/%Y").strftime("%Y-%m-%d")
    except ValueError:
        pass
    return raw


# ── FMP fetchers ──────────────────────────────────────────────────────────────

def _fmp_fetch(endpoint: str, params: dict) -> list:
    """GET an FMP endpoint; return parsed JSON list or []."""
    if not FMP_API_KEY:
        return []
    params["apikey"] = FMP_API_KEY
    try:
        r = requests.get(f"{FMP_BASE}{endpoint}", params=params,
                         headers=HEADERS, timeout=30)
        if r.status_code == 429:
            log.warning("FMP rate-limited on %s — quota may be exhausted for today", endpoint)
            return []
        if r.status_code == 403:
            log.warning("FMP 403 on %s — endpoint may require a higher plan", endpoint)
            return []
        r.raise_for_status()
        data = r.json()
        if isinstance(data, list):
            return data
        if isinstance(data, dict) and "data" in data:
            return data["data"]
        return []
    except Exception as e:
        log.warning("FMP fetch failed for %s: %s", endpoint, e)
        return []


def _parse_fmp_trade(tx: dict, chamber: str, covered: set, cutoff: str) -> dict | None:
    """Map an FMP congressional trade record to our internal schema."""
    # FMP uses 'symbol' for ticker
    ticker = (tx.get("symbol") or tx.get("ticker") or "").strip().upper()
    if not ticker or len(ticker) > 5:
        return None
    if ticker not in covered:
        return None

    tx_date = _norm_date(tx.get("transactionDate") or tx.get("transaction_date") or "")
    if not tx_date or tx_date < cutoff:
        return None

    # Name: FMP senate uses firstName/lastName; house uses representative
    first = (tx.get("firstName") or tx.get("first_name") or "").strip()
    last  = (tx.get("lastName")  or tx.get("last_name")  or "").strip()
    name  = (tx.get("representative") or tx.get("senator") or
             f"{first} {last}".strip() or "")

    return {
        "chamber":    chamber,
        "name":       name.strip(),
        "party":      (tx.get("party") or "").strip(),
        "ticker":     ticker,
        "asset":      (tx.get("assetDescription") or tx.get("asset_description") or "").strip(),
        "tx_type":    _norm_type(tx.get("type") or ""),
        "amount":     (tx.get("amount") or "").strip(),
        "tx_date":    tx_date,
        "disclosure": _norm_date(tx.get("disclosureDate") or tx.get("disclosure_date") or ""),
    }


def fetch_senate_fmp(covered: set, cutoff: str) -> list:
    trades = []
    data = _fmp_fetch("/stable/senate-latest-trading", {"limit": 3000})
    log.info("Senate (FMP): fetched %d raw records", len(data))
    for tx in data:
        t = _parse_fmp_trade(tx, "Senate", covered, cutoff)
        if t:
            trades.append(t)
    return trades


def fetch_house_fmp(covered: set, cutoff: str) -> list:
    trades = []
    data = _fmp_fetch("/stable/house-latest-trading", {"limit": 3000})
    log.info("House (FMP): fetched %d raw records", len(data))
    for tx in data:
        t = _parse_fmp_trade(tx, "House", covered, cutoff)
        if t:
            trades.append(t)
    return trades


# ── Summary builder ───────────────────────────────────────────────────────────

def build_ticker_summary(trades: list) -> dict:
    """Per-ticker aggregation: purchase/sale counts + most recent date."""
    summary = {}
    for t in trades:
        sym = t["ticker"]
        if sym not in summary:
            summary[sym] = {"purchases": 0, "sales": 0, "last_date": ""}
        if t["tx_type"] == "purchase":
            summary[sym]["purchases"] += 1
        elif t["tx_type"] == "sale":
            summary[sym]["sales"] += 1
        if t["tx_date"] > summary[sym]["last_date"]:
            summary[sym]["last_date"] = t["tx_date"]
    return summary


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    log.info("=== Congressional trades fetch  %s ===", datetime.date.today())

    if not FMP_API_KEY:
        log.error("FMP_API_KEY not set — cannot fetch congressional data")
        return

    covered = load_covered()
    if not covered:
        log.warning("No covered tickers in outputs.csv — nothing to filter against")
        return

    cutoff = (datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)).isoformat()
    log.info("Covered tickers: %d  |  Cutoff: %s", len(covered), cutoff)

    senate_trades = fetch_senate_fmp(covered, cutoff)
    house_trades  = fetch_house_fmp(covered, cutoff)
    all_trades    = sorted(senate_trades + house_trades,
                           key=lambda x: x.get("tx_date") or "", reverse=True)

    ticker_summary = build_ticker_summary(all_trades)

    log.info("Senate: %d  House: %d  Total: %d  Tickers: %d",
             len(senate_trades), len(house_trades),
             len(all_trades), len(ticker_summary))

    payload = {
        "generated":       datetime.date.today().isoformat(),
        "lookback_days":   LOOKBACK_DAYS,
        "trade_count":     len(all_trades),
        "ticker_count":    len(ticker_summary),
        "trades":          all_trades,
        "ticker_summary":  ticker_summary,
    }

    os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    log.info("Saved → %s", OUTPUT_JSON)
    _push_to_github(OUTPUT_JSON, "static/data/congress_cache.json",
                    f"congress: {len(all_trades)} trades across {len(ticker_summary)} tickers")

    print(f"Done: {len(all_trades)} trades ({len(senate_trades)} Senate, "
          f"{len(house_trades)} House) across {len(ticker_summary)} tickers.")


if __name__ == "__main__":
    run()
