"""
daily_news.py — Fetch latest news via Google News RSS (primary), Seeking Alpha RSS
(secondary), and Yahoo Finance RSS (fallback).

Reads tickers dynamically from static/reports/*_report.html.
Saves to static/data/news_cache.json.
Designed to run standalone (cron / scheduler) — never called from request handlers.

All sources are free-tier, no API keys required.
  Google News RSS:   one call per ticker, up to GOOGLE_LIMIT articles
  Seeking Alpha RSS: one call per ticker, up to SA_LIMIT articles
  Yahoo Finance RSS: one call per ticker, up to YAHOO_LIMIT articles
"""

import os
import re
import sys
import json
import glob
import time
import logging
import datetime
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime

import requests

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR = os.path.join(SCRIPT_DIR, "static", "reports")
CACHE_PATH  = os.path.join(SCRIPT_DIR, "static", "data", "news_cache.json")

HEADERS      = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
GOOGLE_LIMIT = 10
SA_LIMIT     = 10
YAHOO_LIMIT  = 10
REQUEST_TIMEOUT = 12

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


def _parse_rfc822_date(raw: str) -> str:
    """Convert RFC 822 pubDate string to ISO 8601. Returns raw string on failure."""
    try:
        return parsedate_to_datetime(raw).isoformat()
    except Exception:
        return raw


def fetch_google_news_rss(ticker: str, limit: int = GOOGLE_LIMIT) -> list:
    """Fetch per-ticker news from Google News RSS. No API key required."""
    url = (
        f"https://news.google.com/rss/search"
        f"?q={ticker}+stock&hl=en-US&gl=US&ceid=US:en"
    )
    try:
        r = requests.get(url, timeout=REQUEST_TIMEOUT, headers=HEADERS)
        if r.status_code != 200:
            log.warning("%s: Google News HTTP %d", ticker, r.status_code)
            return []

        root     = ET.fromstring(r.text)
        articles = []

        for item in root.findall(".//item")[:limit]:
            title   = item.findtext("title", "").strip()
            link    = item.findtext("link", "").strip()
            pub     = _parse_rfc822_date(item.findtext("pubDate", ""))
            src_el  = item.find("source")
            site    = src_el.text.strip() if src_el is not None and src_el.text else "Google News"
            articles.append({
                "title":         title,
                "url":           link,
                "publishedDate": pub,
                "site":          site,
                "text":          "",   # Google News RSS has no article body
                "symbol":        ticker,
                "image":         "",
                "source_type":   "google",
            })

        return articles

    except Exception as e:
        log.warning("%s: Google News error — %s", ticker, e)
        return []


def fetch_seeking_alpha_rss(ticker: str, limit: int = SA_LIMIT) -> list:
    """Fetch per-ticker news from Seeking Alpha RSS. No API key required."""
    url = f"https://seekingalpha.com/api/sa/combined/{ticker}.xml"
    try:
        r = requests.get(url, timeout=REQUEST_TIMEOUT, headers=HEADERS)
        if r.status_code != 200:
            log.warning("%s: Seeking Alpha HTTP %d", ticker, r.status_code)
            return []

        root     = ET.fromstring(r.text)
        articles = []

        for item in root.findall(".//item")[:limit]:
            title = item.findtext("title", "").strip()
            link  = item.findtext("link", "").strip()
            pub   = _parse_rfc822_date(item.findtext("pubDate", ""))
            desc  = item.findtext("description", "").strip()

            desc_text = re.sub(r"<[^>]+>", "", desc).strip()

            articles.append({
                "title":         title,
                "url":           link,
                "publishedDate": pub,
                "site":          "Seeking Alpha",
                "text":          desc_text[:300],
                "symbol":        ticker,
                "image":         "",
                "source_type":   "seeking_alpha",
            })

        return articles

    except Exception as e:
        log.warning("%s: Seeking Alpha error — %s", ticker, e)
        return []


def fetch_yahoo_rss(ticker: str, limit: int = YAHOO_LIMIT) -> list:
    """Fetch per-ticker news from Yahoo Finance RSS. No API key required."""
    url = f"https://feeds.finance.yahoo.com/rss/2.0/headline?s={ticker}&region=US&lang=en-US"
    try:
        r = requests.get(url, timeout=REQUEST_TIMEOUT, headers=HEADERS)
        if r.status_code != 200:
            return []

        root     = ET.fromstring(r.text)
        articles = []

        for item in root.findall(".//item")[:limit]:
            title    = item.findtext("title", "").strip()
            link     = item.findtext("link", "").strip()
            pub_date = item.findtext("pubDate", "").strip()
            desc     = item.findtext("description", "").strip()
            src_el   = item.find("source")
            source   = src_el.text.strip() if src_el is not None and src_el.text else "Yahoo Finance"

            articles.append({
                "title":         title,
                "url":           link,
                "publishedDate": _parse_rfc822_date(pub_date),
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


def fetch_all_for_ticker(ticker: str) -> list:
    """Fetch and deduplicate articles from all sources for one ticker."""
    seen_urls = set()
    merged    = []

    for articles in [
        fetch_google_news_rss(ticker),
        fetch_seeking_alpha_rss(ticker),
        fetch_yahoo_rss(ticker),
    ]:
        for art in articles:
            key = (art.get("url") or "").lower().strip()
            if key and key not in seen_urls:
                seen_urls.add(key)
                merged.append(art)

    return merged


def run():
    tickers = discover_tickers()
    if not tickers:
        log.warning("No tickers found in static/reports/. Generate at least one report first.")
        return

    log.info("Fetching news for %d tickers: %s", len(tickers), ", ".join(tickers))

    all_articles = []
    stats = {"google": 0, "seeking_alpha": 0, "yahoo": 0, "total": 0}

    for i, ticker in enumerate(tickers):
        articles = fetch_all_for_ticker(ticker)
        for art in articles:
            stats[art.get("source_type", "other")] = stats.get(art.get("source_type", "other"), 0) + 1
        all_articles.extend(articles)
        log.info("  %s: %d articles", ticker, len(articles))
        if i < len(tickers) - 1:
            time.sleep(0.5)   # polite delay between tickers

    # Sort newest first
    all_articles.sort(key=lambda a: a.get("publishedDate") or "", reverse=True)

    stats["total"] = len(all_articles)
    log.info(
        "Total: %d articles  (Google: %d, SeekingAlpha: %d, Yahoo: %d)",
        stats["total"], stats.get("google", 0),
        stats.get("seeking_alpha", 0), stats.get("yahoo", 0),
    )

    cache = {
        "fetched":  datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S"),
        "tickers":  tickers,
        "articles": all_articles,
        "stats":    stats,
    }

    os.makedirs(os.path.dirname(CACHE_PATH), exist_ok=True)
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)

    print(
        f"Saved {stats['total']} articles "
        f"(Google: {stats.get('google',0)}, "
        f"SeekingAlpha: {stats.get('seeking_alpha',0)}, "
        f"Yahoo: {stats.get('yahoo',0)}) -> {CACHE_PATH}"
    )

    _push_to_github(CACHE_PATH, "static/data/news_cache.json",
                    f"news: refresh {len(tickers)} tickers ({stats['total']} articles)")


def _push_to_github(local_path: str, repo_path: str, commit_msg: str) -> None:
    """Push a local file to GitHub so it survives Render redeploys."""
    import base64
    token  = os.environ.get("GITHUB_TOKEN", "")
    repo   = os.environ.get("GITHUB_REPO", "jaysang2908/Investment-Automation")
    branch = os.environ.get("GITHUB_BRANCH", "main")

    if not token:
        log.info("GITHUB_TOKEN not set — skipping GitHub push for %s", repo_path)
        return

    try:
        with open(local_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode()

        gh_api = f"https://api.github.com/repos/{repo}/contents/{repo_path}"
        gh_hdr = {
            "Authorization": f"token {token}",
            "Accept":        "application/vnd.github.v3+json",
        }

        r = requests.get(gh_api, headers=gh_hdr, params={"ref": branch}, timeout=8)
        sha = r.json().get("sha") if r.status_code == 200 else None

        payload = {"message": commit_msg, "branch": branch, "content": content_b64}
        if sha:
            payload["sha"] = sha

        requests.put(gh_api, headers=gh_hdr, json=payload, timeout=15)
        log.info("GitHub push OK — %s", repo_path)

    except Exception as e:
        log.warning("GitHub push failed for %s: %s", repo_path, e)


if __name__ == "__main__":
    run()
