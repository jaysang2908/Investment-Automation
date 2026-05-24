"""
daily_long_term_bets.py — Daily batch runner for Long Term Bets ticker analysis.

Reads watchlist.txt (S&P 500 or custom list), skips tickers already in
outputs.csv, and processes the next BATCH_SIZE tickers by POSTing to the
live /generate endpoint.

FMP retry logic:
  - Up to MAX_RETRIES attempts per ticker with exponential backoff
  - 429 / quota errors: back off and retry
  - Empty data (no financials): mark as permanently skipped
  - Other failures: leave in queue, retry next run

Designed to run as a Render cron (02:00 UTC daily). Laptop does not need to be on.
"""

import os
import csv
import json
import time
import logging
import datetime
import requests

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s")
log = logging.getLogger(__name__)

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
WATCHLIST    = os.path.join(SCRIPT_DIR, "watchlist.txt")
OUTPUTS_CSV  = os.path.join(SCRIPT_DIR, "outputs.csv")
SKIP_FILE    = os.path.join(SCRIPT_DIR, "skipped_tickers.txt")

RENDER_URL   = os.environ.get("RENDER_INTERNAL_URL",
               "https://investment-automation.onrender.com")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

BATCH_SIZE   = int(os.environ.get("LTB_BATCH_SIZE", "10"))
MAX_RETRIES  = 3
REQUEST_TIMEOUT = 180   # /generate can take up to 2 min on cold FMP


# ── Helpers ───────────────────────────────────────────────────────────────────

def load_watchlist() -> list:
    """Return tickers from watchlist.txt, ignoring comments and blanks."""
    tickers = []
    try:
        with open(WATCHLIST, encoding="utf-8") as f:
            for line in f:
                t = line.split("#")[0].strip().upper()
                if t:
                    tickers.append(t)
    except FileNotFoundError:
        log.error("watchlist.txt not found at %s", WATCHLIST)
    return tickers


def load_completed() -> set:
    """Return tickers already in outputs.csv (any row = completed)."""
    done = set()
    try:
        with open(OUTPUTS_CSV, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                t = (row.get("Ticker") or row.get("ticker") or "").strip().upper()
                if t:
                    done.add(t)
    except FileNotFoundError:
        pass
    return done


def load_skipped() -> set:
    """Return permanently-skipped tickers (no FMP data available)."""
    skipped = set()
    try:
        with open(SKIP_FILE, encoding="utf-8") as f:
            for line in f:
                t = line.split("#")[0].strip().upper()
                if t:
                    skipped.add(t)
    except FileNotFoundError:
        pass
    return skipped


def mark_skipped(ticker: str, reason: str) -> None:
    """Append a ticker to skipped_tickers.txt so it's never retried."""
    with open(SKIP_FILE, "a", encoding="utf-8") as f:
        f.write(f"{ticker}  # {reason}  {datetime.date.today()}\n")
    log.info("Marked %s as permanently skipped: %s", ticker, reason)


def generate_ticker(ticker: str) -> dict:
    """
    POST to /generate for one ticker. Returns dict with keys:
      success (bool), status_code (int), error (str), permanent_fail (bool)
    """
    url     = f"{RENDER_URL}/generate"
    payload = {"ticker": ticker, "password": APP_PASSWORD}

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            log.info("  %s  attempt %d/%d", ticker, attempt, MAX_RETRIES)
            r = requests.post(url, json=payload, timeout=REQUEST_TIMEOUT)

            if r.status_code == 200:
                data  = r.json()
                score = data.get("auto_score", "?")
                log.info("  %s  OK — score=%s", ticker, score)
                return {"success": True, "status_code": 200, "score": score}

            body = ""
            try:
                body = r.json().get("error", r.text[:200])
            except Exception:
                body = r.text[:200]

            # No financial data → permanent skip, no point retrying
            if r.status_code == 400 and "No financial data" in body:
                return {"success": False, "status_code": 400,
                        "error": body, "permanent_fail": True}

            # Rate-limited or server error → back off and retry
            delay = 30 * attempt
            log.warning("  %s  HTTP %d (%s) — waiting %ds before retry",
                         ticker, r.status_code, body[:80], delay)
            if attempt < MAX_RETRIES:
                time.sleep(delay)

        except requests.exceptions.Timeout:
            log.warning("  %s  timeout on attempt %d", ticker, attempt)
            if attempt < MAX_RETRIES:
                time.sleep(20 * attempt)

        except Exception as e:
            log.warning("  %s  error on attempt %d: %s", ticker, attempt, e)
            if attempt < MAX_RETRIES:
                time.sleep(10 * attempt)

    return {"success": False, "status_code": 0,
            "error": "max retries exceeded", "permanent_fail": False}


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    log.info("=== Long Term Bets daily batch  %s ===", datetime.date.today())

    watchlist = load_watchlist()
    completed = load_completed()
    skipped   = load_skipped()

    log.info("Watchlist: %d  Completed: %d  Skipped: %d",
             len(watchlist), len(completed), len(skipped))

    # Queue = watchlist tickers not yet done and not permanently skipped
    queue = [t for t in watchlist if t not in completed and t not in skipped]
    log.info("Queue (remaining): %d tickers", len(queue))

    if not queue:
        log.info("Queue empty — all watchlist tickers completed or skipped.")
        return

    batch = queue[:BATCH_SIZE]
    log.info("Processing batch of %d: %s", len(batch), ", ".join(batch))

    results = {"success": [], "retry_tomorrow": [], "permanent_skip": []}

    for i, ticker in enumerate(batch, 1):
        log.info("[%d/%d] %s", i, len(batch), ticker)
        result = generate_ticker(ticker)

        if result["success"]:
            results["success"].append(ticker)
        elif result.get("permanent_fail"):
            mark_skipped(ticker, result.get("error", "no data"))
            results["permanent_skip"].append(ticker)
        else:
            results["retry_tomorrow"].append(ticker)

        # Pace calls — /generate makes ~5 FMP calls; give FMP a breath
        if i < len(batch):
            time.sleep(10)

    # ── Summary ───────────────────────────────────────────────────────────────
    log.info("")
    log.info("=== Batch complete ===")
    log.info("Success (%d):         %s", len(results["success"]),
             ", ".join(results["success"]) or "—")
    log.info("Retry tomorrow (%d):  %s", len(results["retry_tomorrow"]),
             ", ".join(results["retry_tomorrow"]) or "—")
    log.info("Permanently skipped (%d): %s", len(results["permanent_skip"]),
             ", ".join(results["permanent_skip"]) or "—")
    log.info("Remaining in queue after this run: %d",
             len(queue) - len(batch))

    print(
        f"Done: {len(results['success'])} OK, "
        f"{len(results['retry_tomorrow'])} retry, "
        f"{len(results['permanent_skip'])} skipped. "
        f"Queue remaining: {len(queue) - len(batch)}"
    )


if __name__ == "__main__":
    run()
