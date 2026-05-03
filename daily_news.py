"""
daily_news.py — Fetch latest news via FMP stock_news API (primary) + Yahoo Finance RSS (fallback).

Reads tickers dynamically from static/reports/*_report.html.
Saves to static/data/news_cache.json.
Designed to run standalone (cron / scheduler) — never called from request handlers.

FMP batch endpoint: one call for all tickers, up to 200 articles.
Yahoo RSS:          per-ticker fallback, used when FMP returns nothing for a ticker.
"""

import os
import sys
import json
import glob
import logging
import datetime
import xml.etree.ElementTree as ET

import requests

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR = os.path.join(SCRIPT_DIR, "static", "reports")
CACHE_PATH  = os.path.join(SCRIPT_DIR, "static", "data", "news_cache.json")

# FMP API key — read from environment (same var as server.py), fall back to hardcoded key.
try:
    import fmp_3statementv6 as _mdl
    _FALLBACK_KEY = _mdl.API_KEY
except Exception:
    _FALLBACK_KEY = ""
FMP_API_KEY = os.environ.get("FMP_API_KEY", _FALLBACK_KEY)

FMP_NEWS_LIMIT = 200    # max articles per batch call
YAHOO_LIMIT    = 6      # articles per ticker from Yahoo (fallback only)

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s")
log = logging.getLogger(__name__)


def discover_tickers() -> list:
    """Scan static/reports/ for *_report.html and return sorted ticker list."""
    pattern = os.path.join(REPORTS_DIR, "*_report.html")
    tickers = []
    for path in glob.glob(pattern):
        fname  = os.path.basename(path)
        ticker = fname.replace("_report.html", "")
        if ticker:
            tickers.append(ticker)
    return sorted(set(tickers))


# ── FMP news ──────────────────────────────────────────────────────────────────

def fetch_fmp_news(tickers: list, limit: int = FMP_NEWS_LIMIT) -> list:
    """Batch-fetch news from FMP /stable/news/stock for all tickers at once.

    Returns normalised article dicts with source_type='fmp'.
    Falls back to the v3 endpoint if the stable one fails.
    """
    if not FMP_API_KEY or not tickers:
        return []

    tickers_str = ",".join(tickers)
    articles = []

    # Try stable endpoint first, fall back to v3.
    endpoints = [
        (f"https://financialmodelingprep.com/stable/news/stock"
         f"?tickers={tickers_str}&limit={limit}&apikey={FMP_API_KEY}"),
        (f"https://financialmodelingprep.com/api/v3/stock_news"
         f"?tickers={tickers_str}&limit={limit}&apikey={FMP_API_KEY}"),
    ]

    for url in endpoints:
        try:
            resp = requests.get(url, timeout=12, headers={"User-Agent": "Mozilla/5.0"})
            if resp.status_code != 200:
                log.warning("FMP news HTTP %d — %s", resp.status_code, url[:80])
                continue
            data = resp.json()
            if not isinstance(data, list):
                log.warning("FMP news unexpected response type: %s", type(data))
                continue

            for item in data:
                sym  = (item.get("symbol") or item.get("tickers") or "").upper()
                # FMP sometimes returns comma-separated tickers; take the first match.
                if "," in sym:
                    matched = [t for t in sym.split(",") if t.strip() in tickers]
                    sym = matched[0] if matched else sym.split(",")[0].strip()

                pub = item.get("publishedDate") or item.get("date") or ""
                # FMP dates may be "2024-01-15 10:30:00" — normalise to ISO.
                try:
                    if pub and " " in pub:
                        pub = pub.replace(" ", "T")
                except Exception:
                    pass

                articles.append({
                    "title":         (item.get("title") or "").strip(),
                    "url":           (item.get("url") or "").strip(),
                    "publishedDate": pub,
                    "site":          (item.get("site") or "FMP / Newswire").strip(),
                    "text":          (item.get("text") or item.get("description") or "").strip(),
                    "symbol":        sym,
                    "image":         item.get("image") or "",
                    "source_type":   "fmp",
                })

            if articles:
                log.info("FMP batch: %d articles for %d tickers", len(articles), len(tickers))
                return articles

        except Exception as e:
            log.warning("FMP news error (%s): %s", url[:60], e)
            continue

    log.warning("FMP news: no articles returned from either endpoint")
    return []


# ── Yahoo RSS (fallback) ──────────────────────────────────────────────────────

def fetch_yahoo_rss(ticker: str, limit: int = YAHOO_LIMIT) -> list:
    """Fetch news for a single ticker via Yahoo Finance RSS. Returns normalised dicts."""
    url = f"https://feeds.finance.yahoo.com/rss/2.0/headline?s={ticker}&region=US&lang=en-US"
    try:
        resp = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
        if resp.status_code != 200:
            return []

        root     = ET.fromstring(resp.text)
        articles = []

        for item in root.findall(".//item")[:limit]:
            title    = item.findtext("title", "").strip()
            link     = item.findtext("link", "").strip()
            pub_date = item.findtext("pubDate", "").strip()
            desc     = item.findtext("description", "").strip()
            src_el   = item.find("source")
            source   = src_el.text.strip() if src_el is not None and src_el.text else "Yahoo Finance"

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
                "text":          desc,
                "symbol":        ticker,
                "image":         "",
                "source_type":   "yahoo",
            })

        return articles

    except Exception as e:
        log.warning("%s: Yahoo RSS error — %s", ticker, e)
        return []


# ── Merge + deduplicate ───────────────────────────────────────────────────────

def merge_and_dedup(fmp_articles: list, yahoo_articles: list) -> list:
    """Merge FMP and Yahoo articles. FMP takes priority; Yahoo fills gaps.

    Deduplicated by URL (case-insensitive). Yahoo articles are only kept for
    tickers that got zero FMP coverage.
    """
    seen_urls   = set()
    fmp_tickers = set(a["symbol"].upper() for a in fmp_articles if a.get("symbol"))

    merged = []
    for art in fmp_articles:
        key = (art.get("url") or "").lower().strip()
        if key and key not in seen_urls:
            seen_urls.add(key)
            merged.append(art)

    # Yahoo fallback: only for tickers with no FMP coverage, or if FMP returned nothing.
    for art in yahoo_articles:
        sym = (art.get("symbol") or "").upper()
        if sym in fmp_tickers:
            continue   # FMP already covered this ticker
        key = (art.get("url") or "").lower().strip()
        if key and key not in seen_urls:
            seen_urls.add(key)
            merged.append(art)

    return merged


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    tickers = discover_tickers()
    if not tickers:
        log.warning("No tickers found in static/reports/. Generate at least one report first.")
        return

    log.info("Tickers (%d): %s", len(tickers), ", ".join(tickers))

    # 1. FMP batch fetch (all tickers in one call)
    fmp_articles = fetch_fmp_news(tickers)

    # 2. Yahoo fallback for any ticker that got nothing from FMP
    fmp_covered  = set(a["symbol"].upper() for a in fmp_articles)
    missing      = [t for t in tickers if t not in fmp_covered]
    yahoo_articles = []
    if missing:
        log.info("Yahoo fallback for %d tickers: %s", len(missing), ", ".join(missing))
        for ticker in missing:
            arts = fetch_yahoo_rss(ticker)
            log.info("  %s: %d Yahoo articles", ticker, len(arts))
            yahoo_articles.extend(arts)

    # 3. Merge + deduplicate
    all_articles = merge_and_dedup(fmp_articles, yahoo_articles)

    # 4. Sort newest first
    def _sort_key(a):
        try:
            return a.get("publishedDate") or ""
        except Exception:
            return ""
    all_articles.sort(key=_sort_key, reverse=True)

    # 5. Stats summary
    n_fmp   = sum(1 for a in all_articles if a.get("source_type") == "fmp")
    n_yahoo = sum(1 for a in all_articles if a.get("source_type") == "yahoo")
    log.info("Total: %d articles  (FMP: %d, Yahoo: %d)", len(all_articles), n_fmp, n_yahoo)

    # 6. Write cache
    cache = {
        "fetched":   datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S"),
        "tickers":   tickers,
        "articles":  all_articles,
        "stats":     {"fmp": n_fmp, "yahoo": n_yahoo, "total": len(all_articles)},
    }

    os.makedirs(os.path.dirname(CACHE_PATH), exist_ok=True)
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)

    print(f"Saved {len(all_articles)} articles "
          f"(FMP: {n_fmp}, Yahoo fallback: {n_yahoo}) → {CACHE_PATH}")


if __name__ == "__main__":
    run()
