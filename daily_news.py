"""
daily_news.py — Fetch latest news for all tickers with reports via FMP API.

Reads tickers dynamically from static/reports/*_report.html.
Saves to static/data/news_cache.json.
Designed to run standalone (cron / scheduler) — never called from request handlers.
"""

import os
import sys
import json
import glob
import logging
import datetime

import requests

# ── Resolve paths relative to this script ─────────────────────────────────────
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR  = os.path.join(SCRIPT_DIR, "static", "reports")
CACHE_PATH   = os.path.join(SCRIPT_DIR, "static", "data", "news_cache.json")

# ── API key — same source as the financial model ──────────────────────────────
sys.path.insert(0, SCRIPT_DIR)
import fmp_3statementv6 as mdl

API_KEY = os.environ.get("FMP_API_KEY", mdl.API_KEY)

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s")
log = logging.getLogger(__name__)


def discover_tickers() -> list[str]:
    """Scan static/reports/ for *_report.html and return sorted ticker list."""
    pattern = os.path.join(REPORTS_DIR, "*_report.html")
    tickers = []
    for path in glob.glob(pattern):
        fname = os.path.basename(path)
        ticker = fname.replace("_report.html", "")
        if ticker:
            tickers.append(ticker)
    return sorted(set(tickers))


def fetch_news(tickers: list[str], limit: int = 100) -> list[dict]:
    """Single FMP API call for all tickers. Returns list of article dicts."""
    if not tickers:
        log.warning("No tickers found — nothing to fetch.")
        return []

    url = "https://financialmodelingprep.com/stable/news"
    params = {
        "tickers": ",".join(tickers),
        "limit": limit,
        "apikey": API_KEY,
    }

    try:
        resp = requests.get(url, params=params, timeout=30)
    except requests.RequestException as e:
        log.error("Network error fetching news: %s", e)
        return []

    if resp.status_code in (429, 402):
        log.warning("FMP rate-limit / payment issue (HTTP %d). Skipping.", resp.status_code)
        return []

    if resp.status_code != 200:
        log.error("FMP returned HTTP %d: %s", resp.status_code, resp.text[:300])
        return []

    data = resp.json()
    if not isinstance(data, list):
        log.error("Unexpected response shape: %s", type(data))
        return []

    # Normalise to consistent keys
    articles = []
    for item in data:
        articles.append({
            "title":         item.get("title", ""),
            "url":           item.get("url", ""),
            "publishedDate": item.get("publishedDate", ""),
            "site":          item.get("site") or item.get("source", ""),
            "text":          item.get("text", ""),
            "symbol":        item.get("symbol") or item.get("ticker", ""),
            "image":         item.get("image", ""),
        })

    return articles


def run():
    tickers = discover_tickers()
    log.info("Tickers (%d): %s", len(tickers), ", ".join(tickers))

    articles = fetch_news(tickers)

    cache = {
        "fetched":  datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S"),
        "tickers":  tickers,
        "articles": articles,
    }

    os.makedirs(os.path.dirname(CACHE_PATH), exist_ok=True)
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)

    print(f"Fetched {len(articles)} articles for {len(tickers)} tickers")


if __name__ == "__main__":
    run()
