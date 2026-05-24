"""
daily_long_term_bets.py — Daily batch runner for Long Term Bets ticker analysis.

Reads watchlist.txt, skips completed/rejected tickers, processes the next
BATCH_SIZE tickers by POSTing to the live /generate endpoint, then writes
score-change alerts for the dashboard.

Rejection rules (added to skipped_tickers.txt, pushed to GitHub):
  - No FMP data available        → permanent skip
  - auto_score < REJECT_THRESHOLD → too low quality to re-run

Stale refresh: tickers last run >STALE_DAYS ago are re-queued at lower
priority than new (never-run) tickers.

Runs as a Render cron (02:00 UTC daily). Laptop does not need to be on.
"""

import os
import re
import csv
import json
import time
import base64
import logging
import datetime
import requests

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s")
log = logging.getLogger(__name__)

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
WATCHLIST    = os.path.join(SCRIPT_DIR, "watchlist.txt")
OUTPUTS_CSV  = os.path.join(SCRIPT_DIR, "outputs.csv")
SKIP_FILE    = os.path.join(SCRIPT_DIR, "skipped_tickers.txt")
HISTORY_CSV  = os.path.join(SCRIPT_DIR, "score_history.csv")
ALERTS_JSON  = os.path.join(SCRIPT_DIR, "static", "data", "score_alerts.json")

RENDER_URL      = os.environ.get("RENDER_INTERNAL_URL",
                  "https://investment-automation.onrender.com")
APP_PASSWORD    = os.environ.get("APP_PASSWORD", "")
GITHUB_TOKEN    = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO     = os.environ.get("GITHUB_REPO", "jaysang2908/Investment-Automation")
GITHUB_BRANCH   = os.environ.get("GITHUB_BRANCH", "main")

BATCH_SIZE       = int(os.environ.get("LTB_BATCH_SIZE", "10"))
MAX_RETRIES      = 3
REQUEST_TIMEOUT  = 180    # /generate can take ~2 min on cold FMP
STALE_DAYS       = 90     # re-run tickers older than this
REJECT_THRESHOLD = 2.5    # auto_score (0-10) below this → skip permanently
ALERT_THRESHOLD  = 0.5    # score change (0-10) that triggers a dashboard alert


# ── GitHub helpers ────────────────────────────────────────────────────────────

def _gh_headers():
    return {"Authorization": f"token {GITHUB_TOKEN}",
            "Accept": "application/vnd.github.v3+json"}


def _fetch_file_from_github(repo_path: str) -> str | None:
    """Return decoded file content from GitHub, or None if not found."""
    if not GITHUB_TOKEN:
        return None
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{repo_path}"
        r = requests.get(url, headers=_gh_headers(),
                         params={"ref": GITHUB_BRANCH}, timeout=8)
        if r.status_code == 200:
            return base64.b64decode(r.json()["content"]).decode("utf-8")
    except Exception as e:
        log.warning("GitHub fetch failed for %s: %s", repo_path, e)
    return None


def _push_file_to_github(local_path: str, repo_path: str, commit_msg: str) -> None:
    """Push a local file to GitHub so it survives Render redeploys."""
    if not GITHUB_TOKEN:
        log.info("No GITHUB_TOKEN — skipping push for %s", repo_path)
        return
    try:
        with open(local_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode()
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{repo_path}"
        r = requests.get(url, headers=_gh_headers(),
                         params={"ref": GITHUB_BRANCH}, timeout=8)
        sha = r.json().get("sha") if r.status_code == 200 else None
        payload = {"message": commit_msg, "branch": GITHUB_BRANCH,
                   "content": content_b64}
        if sha:
            payload["sha"] = sha
        requests.put(url, headers=_gh_headers(), json=payload, timeout=15)
        log.info("GitHub push OK — %s", repo_path)
    except Exception as e:
        log.warning("GitHub push failed for %s: %s", repo_path, e)


# ── Watchlist / completed / skipped ──────────────────────────────────────────

def load_watchlist() -> list:
    tickers = []
    try:
        with open(WATCHLIST, encoding="utf-8") as f:
            for line in f:
                t = line.split("#")[0].strip().upper()
                if t:
                    tickers.append(t)
    except FileNotFoundError:
        log.error("watchlist.txt not found")
    return tickers


def load_completed() -> dict:
    """Return {ticker: last_run_date} for all tickers in outputs.csv."""
    done = {}
    try:
        with open(OUTPUTS_CSV, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                t = (row.get("Ticker") or "").strip().upper()
                d = (row.get("Date") or "").strip()
                if t:
                    done[t] = d
    except FileNotFoundError:
        pass
    return done


def load_skipped() -> set:
    """Return permanently-skipped tickers. Prefers GitHub copy (most current)."""
    skipped = set()

    # Try to pull the latest from GitHub first
    remote = _fetch_file_from_github("skipped_tickers.txt")
    if remote is not None:
        for line in remote.splitlines():
            t = line.split("#")[0].strip().upper()
            if t:
                skipped.add(t)
        # Write to local disk so mark_skipped() can append and re-push
        with open(SKIP_FILE, "w", encoding="utf-8") as f:
            f.write(remote if remote.endswith("\n") else remote + "\n")
        log.info("Loaded %d skipped tickers from GitHub", len(skipped))
        return skipped

    # Fall back to local file (first run or no token)
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
    """Append ticker to skipped_tickers.txt and push to GitHub."""
    with open(SKIP_FILE, "a", encoding="utf-8") as f:
        f.write(f"{ticker}  # {reason}  {datetime.date.today()}\n")
    log.info("Marked %s as permanently skipped: %s", ticker, reason)
    _push_file_to_github(SKIP_FILE, "skipped_tickers.txt",
                         f"skip: {ticker} ({reason})")


# ── Generate ──────────────────────────────────────────────────────────────────

def generate_ticker(ticker: str) -> dict:
    """
    POST to /generate. Returns:
      success (bool), score (float|None), error (str), permanent_fail (bool)
    """
    url     = f"{RENDER_URL}/generate"
    payload = {"ticker": ticker, "password": APP_PASSWORD}

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            log.info("  %s  attempt %d/%d", ticker, attempt, MAX_RETRIES)
            r = requests.post(url, json=payload, timeout=REQUEST_TIMEOUT)

            if r.status_code == 200:
                data  = r.json()
                score = data.get("auto_score")
                log.info("  %s  OK — auto_score=%s", ticker, score)
                return {"success": True, "score": score, "permanent_fail": False}

            body = ""
            try:
                body = r.json().get("error", r.text[:200])
            except Exception:
                body = r.text[:200]

            if r.status_code == 400 and "No financial data" in body:
                return {"success": False, "score": None,
                        "error": "no_fmp_data", "permanent_fail": True}

            delay = 30 * attempt
            log.warning("  %s  HTTP %d (%s) — retrying in %ds",
                        ticker, r.status_code, body[:80], delay)
            if attempt < MAX_RETRIES:
                time.sleep(delay)

        except requests.exceptions.Timeout:
            log.warning("  %s  timeout on attempt %d", ticker, attempt)
            if attempt < MAX_RETRIES:
                time.sleep(20 * attempt)
        except Exception as e:
            log.warning("  %s  error: %s", ticker, e)
            if attempt < MAX_RETRIES:
                time.sleep(10 * attempt)

    return {"success": False, "score": None,
            "error": "max_retries_exceeded", "permanent_fail": False}


# ── Score change alerts ───────────────────────────────────────────────────────

def compute_score_alerts() -> list:
    """
    Read score_history.csv and return tickers where the two most recent
    runs differ by >= ALERT_THRESHOLD (on 0-10 normalized scale).
    Raw scores in the CSV are on a 0-87.5 scale; normalize by dividing by 8.75.
    """
    history = {}   # ticker -> list of (date, raw_auto_score)
    try:
        with open(HISTORY_CSV, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                t = (row.get("Ticker") or "").strip().upper()
                d = (row.get("Date") or "").strip()
                s = row.get("Auto_Score") or ""
                if t and d and s:
                    try:
                        history.setdefault(t, []).append((d, float(s)))
                    except ValueError:
                        pass
    except FileNotFoundError:
        return []

    alerts = []
    for ticker, entries in history.items():
        if len(entries) < 2:
            continue
        entries.sort(key=lambda x: x[0])
        prev_date, prev_raw = entries[-2]
        curr_date, curr_raw = entries[-1]
        prev = round(prev_raw / 8.75, 2)
        curr = round(curr_raw / 8.75, 2)
        change = round(curr - prev, 2)
        if abs(change) >= ALERT_THRESHOLD:
            alerts.append({
                "ticker":    ticker,
                "prev_score": prev,
                "curr_score": curr,
                "change":    change,
                "prev_date": prev_date,
                "curr_date": curr_date,
            })

    alerts.sort(key=lambda x: abs(x["change"]), reverse=True)
    return alerts


def write_score_alerts(alerts: list) -> None:
    os.makedirs(os.path.dirname(ALERTS_JSON), exist_ok=True)
    payload = {
        "generated": datetime.date.today().isoformat(),
        "threshold": ALERT_THRESHOLD,
        "alerts":    alerts,
    }
    with open(ALERTS_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)

    if alerts:
        log.info("Score alerts (%d): %s", len(alerts),
                 ", ".join(f"{a['ticker']} {a['change']:+.1f}" for a in alerts[:5]))

    _push_file_to_github(ALERTS_JSON, "static/data/score_alerts.json",
                         f"alerts: {len(alerts)} score changes detected")


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    log.info("=== Long Term Bets daily batch  %s ===", datetime.date.today())

    watchlist = load_watchlist()
    completed = load_completed()   # {ticker: date_str}
    skipped   = load_skipped()

    log.info("Watchlist: %d  Completed: %d  Skipped: %d",
             len(watchlist), len(completed), len(skipped))

    today     = datetime.date.today()
    stale_cutoff = today - datetime.timedelta(days=STALE_DAYS)

    # New = never run, not skipped
    new_queue = [t for t in watchlist
                 if t not in completed and t not in skipped]

    # Stale = run but >STALE_DAYS ago, not skipped
    stale_queue = []
    for t in watchlist:
        if t in skipped or t not in completed:
            continue
        try:
            last = datetime.date.fromisoformat(completed[t])
            if last < stale_cutoff:
                stale_queue.append(t)
        except (ValueError, TypeError):
            pass

    log.info("New queue: %d  Stale queue: %d", len(new_queue), len(stale_queue))

    # New tickers first, stale fills remaining slots
    batch = (new_queue + stale_queue)[:BATCH_SIZE]

    if not batch:
        log.info("Nothing to process — all tickers up to date.")
        # Still compute alerts in case previous runs changed scores
        alerts = compute_score_alerts()
        write_score_alerts(alerts)
        return

    log.info("Batch (%d): %s", len(batch), ", ".join(batch))
    results = {"success": [], "retry_tomorrow": [], "permanent_skip": []}

    for i, ticker in enumerate(batch, 1):
        log.info("[%d/%d] %s", i, len(batch), ticker)
        result = generate_ticker(ticker)

        if result["success"]:
            score = result.get("score")
            # Reject low-quality tickers — not worth re-running in future
            if score is not None and float(score) < REJECT_THRESHOLD:
                mark_skipped(ticker,
                             f"low_quality (score={score:.1f} < {REJECT_THRESHOLD})")
            results["success"].append(ticker)

        elif result.get("permanent_fail"):
            mark_skipped(ticker, result.get("error", "no_data"))
            results["permanent_skip"].append(ticker)

        else:
            results["retry_tomorrow"].append(ticker)

        if i < len(batch):
            time.sleep(10)

    # Compute and write score change alerts after the batch
    alerts = compute_score_alerts()
    write_score_alerts(alerts)

    log.info("")
    log.info("=== Batch complete ===")
    log.info("Success (%d):          %s", len(results["success"]),
             ", ".join(results["success"]) or "—")
    log.info("Retry tomorrow (%d):   %s", len(results["retry_tomorrow"]),
             ", ".join(results["retry_tomorrow"]) or "—")
    log.info("Permanently skipped (%d): %s", len(results["permanent_skip"]),
             ", ".join(results["permanent_skip"]) or "—")
    remaining = len(new_queue) - len([t for t in batch if t in new_queue])
    log.info("New tickers remaining: %d", remaining)

    print(
        f"Done: {len(results['success'])} OK, "
        f"{len(results['retry_tomorrow'])} retry, "
        f"{len(results['permanent_skip'])} skipped. "
        f"Alerts: {len(alerts)}. New remaining: {remaining}"
    )


if __name__ == "__main__":
    run()
