"""
daily_prices.py — refresh dashboard prices from free EOD data (zero FMP quota).

Problem this solves: outputs.csv rows carry the price captured on the ticker's
last full model run. Between runs the dashboard was presenting prices (and the
GG/EM upsides derived from them) that could be 5+ weeks stale, as if current.

What it does, per ticker row in outputs.csv:
  1. Fetch the latest close from Yahoo Finance's chart API (free, no key).
  2. Update Price.
  3. Scale MktCap_B by the price ratio (share count unchanged between runs;
     a >25% jump is logged as a possible split — see S-001).
  4. Recompute GG_Upside / EM_Upside against the stored GG_Price / EM_Price
     (fair values only move on a full re-run; upside moves with price daily).
  5. Scale PE_Current / PFCF_Current by the price ratio (price/EPS and
     price/FCFps are linear in price; denominators only change on re-run).
  Date column is NOT touched — it remains the model run date.

It also stamps static/data/price_refresh.json ({asof, source, updated}) which
the dashboard reads to display "Prices as of <date> close".

Modes:
  python daily_prices.py           — cron mode: reads the LATEST outputs.csv
                                     from GitHub, pushes updated CSV + sidecar
                                     back (Render redeploys on push, which is
                                     how the web container picks up the data).
                                     Requires GITHUB_TOKEN + GITHUB_REPO.
  python daily_prices.py --local   — reads/writes the local files only
                                     (for a local run before a manual commit).

Free-tier etiquette: ~55 sequential Yahoo requests with a 0.4s spacing, once
per weekday after US close (see render.yaml cron).
"""

import base64
import csv
import datetime
import io
import json
import os
import sys
import time

import requests

_UA = {"User-Agent": "Mozilla/5.0 (investment-automation daily price refresh)"}
_YAHOO = "https://query1.finance.yahoo.com/v8/finance/chart/{sym}?range=1d&interval=1d"

GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO  = os.environ.get("GITHUB_REPO", "")   # e.g. "jaysang2908/Investment-Automation"
_GH_API      = "https://api.github.com/repos/{repo}/contents/{path}"

_ROOT = os.path.dirname(os.path.abspath(__file__))
_CSV_LOCAL     = os.path.join(_ROOT, "outputs.csv")
_SIDECAR_LOCAL = os.path.join(_ROOT, "static", "data", "price_refresh.json")


def _fetch_close(ticker: str):
    """Return (price, asof_date_iso) from Yahoo, or (None, None) on any miss."""
    try:
        r = requests.get(_YAHOO.format(sym=ticker), headers=_UA, timeout=10)
        meta = r.json()["chart"]["result"][0]["meta"]
        px = meta.get("regularMarketPrice")
        ts = meta.get("regularMarketTime")
        if not px or px <= 0:
            return None, None
        asof = (datetime.datetime.utcfromtimestamp(ts).date().isoformat()
                if ts else None)
        return float(px), asof
    except Exception as e:
        print(f"  {ticker}: Yahoo fetch failed ({e})")
        return None, None


def _gh_get(path: str):
    """Fetch (text, sha) of a repo file via the contents API."""
    url = _GH_API.format(repo=GITHUB_REPO, path=path)
    r = requests.get(url, headers={"Authorization": f"token {GITHUB_TOKEN}",
                                   "Accept": "application/vnd.github+json"},
                     timeout=15)
    r.raise_for_status()
    j = r.json()
    return base64.b64decode(j["content"]).decode("utf-8"), j["sha"]


def _gh_put(path: str, text: str, sha, msg: str):
    url = _GH_API.format(repo=GITHUB_REPO, path=path)
    body = {"message": msg,
            "content": base64.b64encode(text.encode("utf-8")).decode("ascii")}
    if sha:
        body["sha"] = sha
    r = requests.put(url, headers={"Authorization": f"token {GITHUB_TOKEN}",
                                   "Accept": "application/vnd.github+json"},
                     json=body, timeout=20)
    r.raise_for_status()


def _num(v):
    try:
        return float(v) if v not in ("", None) else None
    except ValueError:
        return None


def refresh_csv_text(content: str):
    """Apply price refresh to CSV text. Returns (new_text, updated, asof)."""
    reader = csv.DictReader(io.StringIO(content))
    cols   = reader.fieldnames or []
    rows   = list(reader)

    updated, latest_asof = 0, None
    for row in rows:
        t = (row.get("Ticker") or "").strip().upper()
        if not t:
            continue
        px, asof = _fetch_close(t)
        time.sleep(0.4)
        if px is None:
            continue
        latest_asof = max(latest_asof or "", asof or "")

        old_px = _num(row.get("Price"))
        row["Price"] = f"{px:.2f}"

        if old_px and old_px > 0:
            ratio = px / old_px
            if abs(ratio - 1) > 0.25:
                print(f"  [S-001-style warning] {t}: price moved "
                      f"{(ratio - 1):+.0%} since last run - possible split or "
                      f"data issue; MktCap/PE scaling may be unreliable.")
            for col in ("MktCap_B", "PE_Current", "PFCF_Current"):
                v = _num(row.get(col))
                if v is not None:
                    dp = 2 if col == "MktCap_B" else 1
                    row[col] = f"{v * ratio:.{dp}f}"

        # Upsides re-anchor to the stored fair values regardless of old price
        # (this also repairs rows that were missing Price, e.g. AMD).
        for px_col, up_col in (("GG_Price", "GG_Upside"), ("EM_Price", "EM_Upside")):
            fv = _num(row.get(px_col))
            if fv is not None and fv > 0:
                row[up_col] = f"{fv / px - 1:.4f}"
        updated += 1
        print(f"  {t}: {old_px if old_px else 'n/a'} -> {px:.2f}")

    out = io.StringIO()
    w = csv.DictWriter(out, fieldnames=cols, lineterminator="\n")
    w.writeheader()
    w.writerows(rows)
    return out.getvalue(), updated, latest_asof


def main():
    local_only = "--local" in sys.argv
    use_github = bool(GITHUB_TOKEN and GITHUB_REPO) and not local_only

    if use_github:
        content, sha = _gh_get("outputs.csv")
    else:
        with open(_CSV_LOCAL, "r", encoding="utf-8") as f:
            content = f.read()
        sha = None

    print(f"Refreshing prices for {content.count(chr(10)) - 1} rows "
          f"({'GitHub' if use_github else 'local'} mode)...")
    new_text, updated, asof = refresh_csv_text(content)
    if not updated:
        print("No prices refreshed — leaving files untouched.")
        return

    sidecar = json.dumps({
        "asof":    asof or datetime.date.today().isoformat(),
        "source":  "Yahoo Finance EOD",
        "updated": updated,
        "run_utc": datetime.datetime.utcnow().isoformat(timespec="seconds"),
    }, indent=2)

    if use_github:
        # Push CSV with the fetched SHA; on a 409/422 (concurrent scheduled run
        # pushed between our GET and PUT) refetch and reapply once.
        try:
            _gh_put("outputs.csv", new_text, sha,
                    f"prices: daily EOD refresh ({updated} tickers, {asof})")
        except requests.HTTPError:
            print("  Conflict on push - refetching latest and reapplying...")
            content, sha = _gh_get("outputs.csv")
            new_text, updated, asof = refresh_csv_text(content)
            _gh_put("outputs.csv", new_text, sha,
                    f"prices: daily EOD refresh ({updated} tickers, {asof})")
        try:
            _, sc_sha = _gh_get("static/data/price_refresh.json")
        except Exception:
            sc_sha = None
        _gh_put("static/data/price_refresh.json", sidecar, sc_sha,
                f"prices: refresh marker {asof}")
    else:
        with open(_CSV_LOCAL, "w", encoding="utf-8", newline="") as f:
            f.write(new_text)
        os.makedirs(os.path.dirname(_SIDECAR_LOCAL), exist_ok=True)
        with open(_SIDECAR_LOCAL, "w", encoding="utf-8") as f:
            f.write(sidecar)

    print(f"Done: {updated} tickers updated, prices as of {asof}.")


if __name__ == "__main__":
    main()
