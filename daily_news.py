"""
daily_news.py — Fetch latest news for all tickers via Yahoo Finance RSS.

Reads tickers dynamically from static/reports/*_report.html.
Saves to static/data/news_cache.json.
Designed to run standalone (cron / scheduler) — never called from request handlers.

No API key required. Uses Yahoo Finance public RSS feeds.
"""

import os
import sys
import json
import glob
import logging
import datetime
import xml.etree.ElementTree as ET

import requests

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR  = os.path.join(SCRIPT_DIR, "static", "reports")
CACHE_PATH   = os.path.join(SCRIPT_DIR, "static", "data", "news_cache.json")

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s")
log = logging.getLogger(__name__)


def discover_tickers() -> list:
    """Scan static/reports/ for *_report.html and return sorted ticker list."""
    pattern = os.path.join(REPORTS_DIR, "*_report.html")
    tickers = []
    for path in glob.glob(pattern):
        fname = os.path.basename(path)
        ticker = fname.replace("_report.html", "")
        if ticker:
            tickers.append(ticker)
    return sorted(set(tickers))


def fetch_yahoo_rss(ticker: str, limit: int = 8) -> list:
    """Fetch news articles for a single ticker via Yahoo Finance RSS."""
    url = f"https://feeds.finance.yahoo.com/rss/2.0/headline?s={ticker}&region=US&lang=en-US"
    try:
        resp = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
        if resp.status_code != 200:
            log.warning("%s: Yahoo RSS returned %d", ticker, resp.status_code)
            return []

        root = ET.fromstring(resp.text)
        articles = []
        ns = {"media": "http://search.yahoo.com/mrss/"}

        for item in root.findall(".//item")[:limit]:
            title       = item.findtext("title", "").strip()
            link        = item.findtext("link", "").strip()
            pub_date    = item.findtext("pubDate", "").strip()
            description = item.findtext("description", "").strip()
            source_el   = item.find("source")
            source      = source_el.text.strip() if source_el is not None and source_el.text else "Yahoo Finance"

            # Normalise pubDate → ISO format
            published = pub_date
            try:
                from email.utils import parsedate_to_datetime
                published = parsedate_to_datetime(pub_date).isoformat()
            except Exception:
                pass

            articles.append({
                "title":         title,
                "url":           link,
                "publishedDate": published,
                "site":          source,
                "text":          description,
                "symbol":        ticker,
                "image":         "",
            })

        return articles

    except Exception as e:
        log.warning("%s: RSS error — %s", ticker, e)
        return []


def run():
    tickers = discover_tickers()
    log.info("Tickers (%d): %s", len(tickers), ", ".join(tickers))

    all_articles = []
    for ticker in tickers:
        articles = fetch_yahoo_rss(ticker)
        log.info("  %s: %d articles", ticker, len(articles))
        all_articles.extend(articles)

    # Sort newest first
    def _sort_key(a):
        try:
            return a["publishedDate"]
        except Exception:
            return ""
    all_articles.sort(key=_sort_key, reverse=True)

    cache = {
        "fetched":  datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S"),
        "tickers":  tickers,
        "articles": all_articles,
    }

    os.makedirs(os.path.dirname(CACHE_PATH), exist_ok=True)
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)

    print(f"Fetched {len(all_articles)} articles for {len(tickers)} tickers -> {CACHE_PATH}")


if __name__ == "__main__":
    run()
