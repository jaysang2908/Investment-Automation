"""
daily_congress.py — Fetch congressional stock trade disclosures for covered tickers.

Data sources (free, no API key — STOCK Act disclosure aggregators):
  House:  house-stock-watcher-data.s3-us-west-2.amazonaws.com/data/all_transactions.json
  Senate: senate-stock-watcher-data.s3-us-west-2.amazonaws.com/data/all_transactions.json

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

GITHUB_TOKEN  = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO   = os.environ.get("GITHUB_REPO", "jaysang2908/Investment-Automation")
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH", "main")

LOOKBACK_DAYS = 180  # keep 6 months in cache; dashboard shows last 30

HOUSE_URL  = "https://house-stock-watcher-data.s3-us-west-2.amazonaws.com/data/all_transactions.json"
SENATE_URL = "https://senate-stock-watcher-data.s3-us-west-2.amazonaws.com/data/all_transactions.json"

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; investment-research-bot/1.0)"}


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


def fetch_house(covered: set, cutoff: str) -> list:
    trades = []
    try:
        r = requests.get(HOUSE_URL, headers=HEADERS, timeout=30)
        r.raise_for_status()
        data = r.json()
        log.info("House: fetched %d total transactions", len(data))
        for tx in data:
            ticker = (tx.get("ticker") or "").strip().upper()
            # Skip non-stock entries (options, bonds, mutual funds with no ticker)
            if not ticker or len(ticker) > 5 or ticker.startswith("--"):
                continue
            if ticker not in covered:
                continue
            tx_date = (tx.get("transaction_date") or "").strip()[:10]
            if not tx_date or tx_date < cutoff:
                continue
            trades.append({
                "chamber":    "House",
                "name":       (tx.get("representative") or "").strip(),
                "party":      "",
                "ticker":     ticker,
                "asset":      (tx.get("asset_description") or "").strip(),
                "tx_type":    _norm_type(tx.get("type") or ""),
                "amount":     (tx.get("amount") or "").strip(),
                "tx_date":    tx_date,
                "disclosure": (tx.get("disclosure_date") or "").strip()[:10],
            })
    except Exception as e:
        log.warning("House fetch failed: %s", e)
    return trades


def fetch_senate(covered: set, cutoff: str) -> list:
    trades = []
    try:
        r = requests.get(SENATE_URL, headers=HEADERS, timeout=30)
        r.raise_for_status()
        data = r.json()
        log.info("Senate: fetched %d total transactions", len(data))
        for tx in data:
            ticker = (tx.get("ticker") or "").strip().upper()
            if not ticker or len(ticker) > 5 or ticker.startswith("--"):
                continue
            if ticker not in covered:
                continue
            tx_date = (tx.get("transaction_date") or "").strip()[:10]
            if not tx_date or tx_date < cutoff:
                continue
            first = (tx.get("first_name") or "").strip()
            last  = (tx.get("last_name")  or "").strip()
            trades.append({
                "chamber":    "Senate",
                "name":       f"{first} {last}".strip(),
                "party":      "",
                "ticker":     ticker,
                "asset":      (tx.get("asset_description") or "").strip(),
                "tx_type":    _norm_type(tx.get("type") or ""),
                "amount":     (tx.get("amount") or "").strip(),
                "tx_date":    tx_date,
                "disclosure": (tx.get("disclosure_date") or "").strip()[:10],
            })
    except Exception as e:
        log.warning("Senate fetch failed: %s", e)
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

    covered = load_covered()
    if not covered:
        log.warning("No covered tickers in outputs.csv — nothing to filter against")
        return

    cutoff = (datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)).isoformat()
    log.info("Covered tickers: %d  |  Cutoff: %s", len(covered), cutoff)

    house_trades  = fetch_house(covered, cutoff)
    senate_trades = fetch_senate(covered, cutoff)
    all_trades    = sorted(house_trades + senate_trades,
                           key=lambda x: x.get("tx_date") or "", reverse=True)

    ticker_summary = build_ticker_summary(all_trades)

    log.info("House: %d  Senate: %d  Total: %d  Tickers: %d",
             len(house_trades), len(senate_trades),
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

    print(f"Done: {len(all_trades)} trades ({len(house_trades)} House, "
          f"{len(senate_trades)} Senate) across {len(ticker_summary)} tickers.")


if __name__ == "__main__":
    run()
