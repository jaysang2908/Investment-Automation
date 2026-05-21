"""
speculative_engine.py — Engine for the "Maybe Unsafe Financial Freedom" speculative tool.

Scores 12 dimensions across two categories:

  Auto-scored (FMP + yfinance + free public APIs):
    1.  Price Momentum        — RSI + price vs 50d/200d MA
    2.  Volume Signal         — recent volume surge vs 30d baseline
    3.  Short Interest        — short float %, days to cover (squeeze fuel)
    4.  Analyst Revisions     — estimate revision direction + upgrade/downgrade balance
    5.  Float Size            — small float/mktcap = explosive potential
    6.  Insider Buying        — open-market purchases last 90 days
    7.  Downside Floor        — net cash position, debt load
    8.  Options Activity      — put/call ratio via yfinance
    9.  Technical Setup       — MACD crossover, 52-week high proximity, EMA stack alignment
    10. Social / Trend Momentum — Google Trends search spike + StockTwits sentiment

  Manual (user-supplied):
    11. Narrative Theme       — dropdown selection + strength rating
    12. Catalyst Quality      — user-described catalyst + timing tier

Scoring: HIGH = 10 pts · MOD = 5 pts · LOW = 0 pts  (max 120)

Verdicts (75% / ~54% / 45% / 30% of 120):
    ≥ 90 → Moonshot Conviction
    ≥ 65 → Strong Speculative
    ≥ 54 → Speculative Play
    ≥ 36 → High Risk
     < 36 → Pass
"""

import datetime
import io

import requests as _req

# ── Constants ──────────────────────────────────────────────────────────────────

TIER_PTS = {"HIGH": 10, "MOD": 5, "LOW": 0}

POLYGON_API_KEY = ""   # set from env in server.py

NARRATIVE_THEMES = [
    "AI / Machine Learning",
    "Defence & Aerospace",
    "Biotech Catalyst",
    "GLP-1 / Weight Loss",
    "Energy Transition",
    "Nuclear / SMR",
    "Crypto-Adjacent",
    "Turnaround / Restructuring",
    "Supply Chain Re-shoring",
    "Space / Deep Tech",
    "Other",
]

VERDICTS = [
    (90, "Moonshot Conviction"),
    (65, "Strong Speculative"),
    (54, "Speculative Play"),
    (36, "High Risk"),
    (0,  "Pass"),
]

MAX_SCORE = 120


def _verdict(score: float) -> str:
    for threshold, label in VERDICTS:
        if score >= threshold:
            return label
    return "Pass"


# ── Mock data (used when FMP quota is exhausted or MOCK_MODE=True) ─────────────

def _mock_data(ticker: str) -> dict:
    """Return plausible dummy signals so the full pipeline renders without API calls."""
    return {
        "ticker":         ticker,
        "mock":           True,
        "current_price":  42.50,
        "market_cap":     3_800_000_000,
        "float_shares":   85_000_000,
        "shares_out":     92_000_000,
        "net_debt":       -120_000_000,          # net cash
        "trailing_rev":   680_000_000,
        "fwd_rev_1yr":    820_000_000,
        "fwd_rev_2yr":    1_010_000_000,
        "current_ev":     3_680_000_000,
        # Momentum
        "rsi_14":         61.4,
        "price_vs_50ma":  0.07,                  # +7% above 50d MA
        "price_vs_200ma": 0.22,                  # +22% above 200d MA
        # Volume
        "vol_10d_avg":    4_200_000,
        "vol_50d_avg":    2_800_000,
        # Short interest
        "short_float_pct": 18.5,
        "days_to_cover":   5.2,
        # Analyst revisions
        "upgrades_90d":    4,
        "downgrades_90d":  1,
        "est_rev_pct":     0.12,                 # +12% estimate revision
        # Insider buying (last 90d open-market)
        "insider_buys":    3,
        "insider_buy_usd": 820_000,
        "insider_sells":   0,
        # Options
        "put_call_ratio":  0.42,
        "call_oi":         45_000,
        "put_oi":          18_900,
        # Balance sheet
        "cash":            380_000_000,
        "total_debt":      260_000_000,
        "ebitda_ttm":      95_000_000,
        # Social / trend momentum
        "trend_ratio":      2.4,
        "best_theme_ratio": 3.1,
        "best_theme_kw":    "AI stocks",
        "all_gt_ratios":    {"MOCK": 2.4, "AI stocks": 3.1, "HBM memory": 2.8},
        "trend_values":     [30,28,32,35,40,38,42,45,55,60,72,80,85,72],
        # yfinance news sentiment
        "news_count":        8,
        "news_7d":           5,
        "news_bullish":      5,
        "news_bearish":      1,
        "news_bullish_pct":  0.83,
        # Reddit sentiment
        "reddit_mentions":   38,
        "reddit_24h":        14,
        "reddit_bullish":    24,
        "reddit_bearish":    6,
        "reddit_bullish_pct": 0.80,
        # Technical setup
        "macd":            1.42,          # MACD line
        "macd_signal":     0.88,          # signal line
        "macd_hist":       0.54,          # histogram (positive = bullish)
        "ema_21":          40.10,
        "ema_50":          38.60,
        "high_52w":        46.80,
        "low_52w":         22.30,
        "pct_from_52w_high": -0.093,      # -9.3% from 52w high
        "company_name":    f"{ticker} Corp.",
        "sector":          "Technology",
        "profile":         {},
    }


# ── Polygon helpers ────────────────────────────────────────────────────────────

def _fetch_polygon_ohlcv(ticker: str, polygon_key: str) -> list | None:
    """Fetch ~500 days of daily OHLCV bars from Polygon, sorted ascending.

    Returns a list of dicts with keys: o, h, l, c, v, t (timestamp ms).
    Returns None if the key is missing or the request fails.
    """
    if not polygon_key:
        return None
    try:
        from datetime import date, timedelta
        end   = date.today().strftime("%Y-%m-%d")
        start = (date.today() - timedelta(days=720)).strftime("%Y-%m-%d")
        url   = (f"https://api.polygon.io/v2/aggs/ticker/{ticker}"
                 f"/range/1/day/{start}/{end}")
        r = _req.get(url, params={"adjusted": "true", "sort": "asc",
                                  "limit": 500, "apiKey": polygon_key},
                     timeout=15)
        return r.json().get("results") or None
    except Exception:
        return None


def _fetch_polygon_ticker_details(ticker: str, polygon_key: str) -> dict:
    """Fetch ticker reference data from Polygon (market cap, shares outstanding)."""
    if not polygon_key:
        return {}
    try:
        r = _req.get(f"https://api.polygon.io/v3/reference/tickers/{ticker}",
                     params={"apiKey": polygon_key}, timeout=10)
        return r.json().get("results") or {}
    except Exception:
        return {}


def _fetch_polygon_news(ticker: str, polygon_key: str) -> dict:
    """Fetch recent news from Polygon with built-in per-ticker sentiment.

    Polygon's insights array carries a sentiment field ('positive' / 'negative' /
    'neutral') pre-scored by their NLP model — cleaner than keyword matching.
    """
    if not polygon_key:
        return {"news_count": None, "news_bullish_pct": None,
                "news_error": "No Polygon API key"}
    try:
        r = _req.get("https://api.polygon.io/v2/reference/news",
                     params={"ticker": ticker, "limit": 20, "apiKey": polygon_key},
                     timeout=10)
        articles = r.json().get("results") or []
        if not articles:
            return {"news_count": 0, "news_7d": 0, "news_bullish": 0,
                    "news_bearish": 0, "news_bullish_pct": None}

        cutoff_7d = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=7)
        bullish = bearish = recent_7d = 0
        for article in articles:
            insights      = article.get("insights") or []
            ticker_insight = next((i for i in insights if i.get("ticker") == ticker), None)
            if ticker_insight:
                s = ticker_insight.get("sentiment", "neutral")
                if s == "positive":
                    bullish += 1
                elif s == "negative":
                    bearish += 1
            pub = article.get("published_utc", "")
            try:
                ts = datetime.datetime.fromisoformat(pub.replace("Z", "+00:00"))
                if ts > cutoff_7d:
                    recent_7d += 1
            except Exception:
                pass

        total_scored = bullish + bearish
        return {
            "news_count":       len(articles),
            "news_7d":          recent_7d,
            "news_bullish":     bullish,
            "news_bearish":     bearish,
            "news_bullish_pct": (bullish / total_scored) if total_scored else None,
        }
    except Exception as e:
        return {"news_count": None, "news_bullish_pct": None, "news_error": str(e)}


# ── Sentiment keyword sets (used by Reddit scorer) ────────────────────────────

_BULL_WORDS = frozenset([
    "beat", "beats", "upgrade", "upgrades", "surge", "surges", "soar", "soars",
    "strong", "buy", "outperform", "record", "win", "wins", "breakthrough",
    "deal", "contract", "raises", "raised", "rally", "rallies", "bullish",
    "growth", "profit", "gains", "exceeded", "exceeds", "upside", "expansion",
    "accelerat", "positive", "boost", "boosts", "momentum", "milestone",
])
_BEAR_WORDS = frozenset([
    "miss", "misses", "downgrade", "downgrades", "fall", "falls", "drop", "drops",
    "weak", "sell", "underperform", "concern", "concerns", "warn", "warns",
    "warning", "cut", "cuts", "loss", "losses", "decline", "disappoint",
    "disappoints", "bearish", "negative", "shortfall", "slowdown", "risk",
    "risks", "lawsuit", "probe", "investigation", "recall", "layoff", "layoffs",
])


def _text_sentiment(text: str) -> int:
    """Return +1 (bullish), -1 (bearish), 0 (neutral) for a piece of text."""
    lower = text.lower()
    bull  = any(w in lower for w in _BULL_WORDS)
    bear  = any(w in lower for w in _BEAR_WORDS)
    if bull and not bear:
        return 1
    if bear and not bull:
        return -1
    return 0


def _fetch_reddit_sentiment(ticker: str) -> dict:
    """Search r/wallstreetbets, r/stocks, r/investing for ticker mentions (last 7 days).

    Scores each post via keyword sentiment on title + body text, and uses the
    Reddit upvote_ratio as a secondary signal (≥0.75 = bullish crowd, <0.35 = bearish).

    Requires env vars: REDDIT_CLIENT_ID, REDDIT_CLIENT_SECRET.
    Optional: REDDIT_USER_AGENT (defaults to 'InvestmentResearch/1.0').
    Returns None fields gracefully when credentials are absent.
    """
    import os
    client_id     = os.environ.get("REDDIT_CLIENT_ID", "")
    client_secret = os.environ.get("REDDIT_CLIENT_SECRET", "")
    if not client_id or not client_secret:
        return {
            "reddit_mentions": None, "reddit_bullish_pct": None,
            "reddit_error": "Missing REDDIT_CLIENT_ID / REDDIT_CLIENT_SECRET env vars",
        }

    try:
        import praw
        user_agent = os.environ.get("REDDIT_USER_AGENT", "InvestmentResearch/1.0")
        reddit = praw.Reddit(
            client_id=client_id,
            client_secret=client_secret,
            user_agent=user_agent,
        )

        subreddits = ["wallstreetbets", "stocks", "investing", "StockMarket"]
        cutoff_24h = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=24)

        all_posts: list = []
        for sub_name in subreddits:
            try:
                results = reddit.subreddit(sub_name).search(
                    ticker, time_filter="week", limit=15, sort="new"
                )
                all_posts.extend(list(results))
            except Exception:
                pass

        if not all_posts:
            return {"reddit_mentions": 0, "reddit_24h": 0,
                    "reddit_bullish": 0, "reddit_bearish": 0, "reddit_bullish_pct": None}

        bullish = bearish = recent_24h = 0
        for post in all_posts:
            text          = (post.title or "") + " " + (getattr(post, "selftext", "") or "")
            kw_sentiment  = _text_sentiment(text)
            upvote_ratio  = float(getattr(post, "upvote_ratio", 0.5) or 0.5)
            # Bullish if keyword OR crowd upvote signal; bearish if keyword AND downvoted
            if kw_sentiment > 0 or upvote_ratio >= 0.75:
                bullish += 1
            elif kw_sentiment < 0 or upvote_ratio < 0.35:
                bearish += 1
            try:
                ts = datetime.datetime.fromtimestamp(post.created_utc, tz=datetime.timezone.utc)
                if ts > cutoff_24h:
                    recent_24h += 1
            except Exception:
                pass

        total_scored  = bullish + bearish
        bullish_pct   = (bullish / total_scored) if total_scored > 0 else None

        return {
            "reddit_mentions":    len(all_posts),
            "reddit_24h":         recent_24h,
            "reddit_bullish":     bullish,
            "reddit_bearish":     bearish,
            "reddit_bullish_pct": bullish_pct,
        }
    except Exception as e:
        return {"reddit_mentions": None, "reddit_bullish_pct": None, "reddit_error": str(e)}


# ── Live data fetch ────────────────────────────────────────────────────────────

def fetch_speculative_data(ticker: str, api_key: str, polygon_key: str = "",
                           mock: bool = False, narrative_theme: str = "",
                           custom_terms: list[str] | None = None) -> dict:
    """Fetch all signals needed to score the speculative scorecard.

    Falls back to mock data if `mock=True` or on any network failure, so the
    pipeline always has something to work with.
    """
    if mock:
        return _mock_data(ticker)

    data: dict = {"ticker": ticker, "mock": False}

    base = "https://financialmodelingprep.com/stable"

    def _get(path, params=None):
        try:
            p = {"apikey": api_key}
            if params:
                p.update(params)
            r = _req.get(f"{base}/{path}", params=p, timeout=10)
            return r.json()
        except Exception:
            return None

    # ── Profile (FMP — name, sector, price) ──────────────────────────────────
    prof_raw = _get(f"profile?symbol={ticker}")
    profile  = (prof_raw[0] if isinstance(prof_raw, list) and prof_raw else prof_raw or {})
    data["profile"]       = profile
    data["company_name"]  = profile.get("companyName") or ticker
    data["sector"]        = profile.get("sector") or "Unknown"
    data["current_price"] = float(profile.get("price") or 0) or None

    # ── Market cap / shares — Polygon ticker details (more reliable than FMP) ─
    poly_details          = _fetch_polygon_ticker_details(ticker, polygon_key)
    data["market_cap"]    = float(poly_details.get("market_cap") or 0) or None
    data["shares_out"]    = float(poly_details.get("weighted_shares_outstanding") or 0) or None
    data["float_shares"]  = float(poly_details.get("share_class_shares_outstanding") or 0) or None
    # Fallback to FMP profile if Polygon didn't return market cap
    if not data["market_cap"]:
        data["market_cap"]  = float(profile.get("mktCap") or 0) or None
    if not data["shares_out"]:
        data["shares_out"]  = float(profile.get("sharesOutstanding") or 0) or None

    # ── Balance sheet (latest) ────────────────────────────────────────────────
    bs_raw = _get(f"balance-sheet-statement?symbol={ticker}&limit=1")
    bs     = (bs_raw[0] if isinstance(bs_raw, list) and bs_raw else {})
    cash       = float(bs.get("cashAndCashEquivalents") or 0)
    total_debt = float(bs.get("totalDebt") or 0)
    data["cash"]       = cash
    data["total_debt"] = total_debt
    data["net_debt"]   = total_debt - cash       # positive = net debt, negative = net cash

    # ── Income statement (latest, for trailing revenue + EBITDA) ─────────────
    is_raw = _get(f"income-statement?symbol={ticker}&limit=1")
    is_    = (is_raw[0] if isinstance(is_raw, list) and is_raw else {})
    data["trailing_rev"] = float(is_.get("revenue") or 0) or None
    data["ebitda_ttm"]   = float(is_.get("ebitda") or 0) or None

    # ── EV ────────────────────────────────────────────────────────────────────
    mktcap   = data["market_cap"] or 0
    net_debt = data["net_debt"]   or 0
    data["current_ev"] = mktcap + net_debt

    # ── Analyst estimates (fwd revenue) ───────────────────────────────────────
    ae_raw = _get(f"analyst-estimates?symbol={ticker}&period=annual&limit=3")
    fwd_revs = []
    if isinstance(ae_raw, list):
        today_yr = datetime.date.today().year
        for e in sorted(ae_raw, key=lambda x: x.get("date", "")):
            yr = int(str(e.get("date", "0"))[:4])
            if yr >= today_yr:
                rev = float(e.get("estimatedRevenueAvg") or e.get("revenueAvg") or 0)
                if rev:
                    fwd_revs.append(rev)
    data["fwd_rev_1yr"] = fwd_revs[0] if len(fwd_revs) > 0 else None
    data["fwd_rev_2yr"] = fwd_revs[1] if len(fwd_revs) > 1 else None

    # ── OHLCV + all technical indicators — Polygon + pandas-ta ────────────────
    # Polygon returns full bars (O/H/L/C/V), sorted ascending. pandas-ta then
    # computes RSI/MACD/EMA locally — no extra API calls needed.
    data["price_vs_50ma"]  = None
    data["price_vs_200ma"] = None
    data["vol_10d_avg"]    = None
    data["vol_50d_avg"]    = None
    data["rsi_14"]         = None
    data["macd"]           = data["macd_signal"] = data["macd_hist"] = None
    data["ema_21"]         = data["ema_50"]       = None
    data["high_52w"]       = data["low_52w"]      = data["pct_from_52w_high"] = None

    poly_bars = _fetch_polygon_ohlcv(ticker, polygon_key)
    if poly_bars and len(poly_bars) >= 50:
        closes_asc = [float(b["c"]) for b in poly_bars]
        highs_asc  = [float(b["h"]) for b in poly_bars]
        lows_asc   = [float(b["l"]) for b in poly_bars]
        vols_asc   = [float(b["v"]) for b in poly_bars]

        # Descending lists for simple slice-based MA / volume calcs
        closes_desc = list(reversed(closes_asc))
        vols_desc   = list(reversed(vols_asc))

        if len(closes_desc) >= 200:
            cur   = closes_desc[0]
            ma50  = sum(closes_desc[:50])  / 50
            ma200 = sum(closes_desc[:200]) / 200
            data["price_vs_50ma"]  = (cur / ma50  - 1) if ma50  else None
            data["price_vs_200ma"] = (cur / ma200 - 1) if ma200 else None

        if len(vols_desc) >= 50:
            data["vol_10d_avg"] = sum(vols_desc[:10]) / 10
            data["vol_50d_avg"] = sum(vols_desc[:50]) / 50

        # 52-week high / low — use actual H/L bars from Polygon (not just close)
        year_bars = min(len(poly_bars), 252)
        year_highs = highs_asc[-year_bars:]
        year_lows  = lows_asc[-year_bars:]
        data["high_52w"] = max(year_highs)
        data["low_52w"]  = min(year_lows)
        if data["high_52w"] > 0 and closes_desc:
            data["pct_from_52w_high"] = (closes_desc[0] / data["high_52w"]) - 1

        # ta library — RSI/MACD/EMA computed from Polygon closes, no extra API calls
        try:
            import pandas as pd
            import ta as ta_lib

            close_s = pd.Series(closes_asc)

            rsi_v = ta_lib.momentum.RSIIndicator(close=close_s, window=14).rsi().iloc[-1]
            data["rsi_14"] = float(rsi_v) if pd.notna(rsi_v) else None

            macd_obj = ta_lib.trend.MACD(close=close_s)
            macd_v   = macd_obj.macd().iloc[-1]
            hist_v   = macd_obj.macd_diff().iloc[-1]
            sig_v    = macd_obj.macd_signal().iloc[-1]
            data["macd"]        = float(macd_v) if pd.notna(macd_v) else None
            data["macd_hist"]   = float(hist_v) if pd.notna(hist_v) else None
            data["macd_signal"] = float(sig_v)  if pd.notna(sig_v)  else None

            for period, key in [(21, "ema_21"), (50, "ema_50")]:
                try:
                    ema_v = ta_lib.trend.EMAIndicator(close=close_s, window=period).ema_indicator().iloc[-1]
                    data[key] = float(ema_v) if pd.notna(ema_v) else None
                except Exception:
                    pass
                # Manual fallback if library returns NaN
                if data[key] is None and len(closes_asc) >= period:
                    alpha = 2.0 / (period + 1)
                    ema = sum(closes_asc[:period]) / period
                    for v in closes_asc[period:]:
                        ema = alpha * v + (1 - alpha) * ema
                    data[key] = ema
        except Exception as e:
            data["ta_error"] = str(e)

    # ── Short interest ────────────────────────────────────────────────────────
    si_raw = _get(f"short-interest?symbol={ticker}")
    data["short_float_pct"] = None
    data["days_to_cover"]   = None
    if isinstance(si_raw, list) and si_raw:
        si = si_raw[0]
        data["short_float_pct"] = float(si.get("shortPercentOfFloat") or si.get("shortPercent") or 0) or None
        if data["short_float_pct"]:
            data["short_float_pct"] *= 100   # convert 0.18 → 18.0
        data["days_to_cover"] = float(si.get("daysToCover") or si.get("shortRatio") or 0) or None

    # ── Upgrades / downgrades (last 90 days) ──────────────────────────────────
    cutoff = (datetime.date.today() - datetime.timedelta(days=90)).isoformat()
    ud_raw = _get(f"upgrades-downgrades?symbol={ticker}&limit=50")
    ups = downs = 0
    ud_list = ud_raw if isinstance(ud_raw, list) else (ud_raw.get("data") or [] if isinstance(ud_raw, dict) else [])
    if not ud_list and isinstance(ud_raw, dict):
        data["ud_error"] = ud_raw.get("message") or str(ud_raw)
    for e in ud_list:
        if (e.get("publishedDate") or e.get("date") or "") < cutoff:
            continue
        action = (e.get("action") or e.get("newGrade") or e.get("gradingCompany") or "").lower()
        grade  = (e.get("newGrade") or "").lower()
        if any(w in action or w in grade for w in ("upgrade", "buy", "outperform", "overweight", "positive")):
            ups += 1
        elif any(w in action or w in grade for w in ("downgrade", "sell", "underperform", "underweight", "negative")):
            downs += 1
    data["upgrades_90d"]   = ups
    data["downgrades_90d"] = downs

    # ── Analyst estimate revision (compare latest to 3mo-ago estimate) ────────
    data["est_rev_pct"] = None
    if isinstance(ae_raw, list) and len(ae_raw) >= 2:
        try:
            new_est = float(ae_raw[0].get("estimatedRevenueAvg") or 0)
            old_est = float(ae_raw[1].get("estimatedRevenueAvg") or 0)
            if old_est:
                data["est_rev_pct"] = (new_est - old_est) / abs(old_est)
        except Exception:
            pass

    # ── Insider trading (last 90 days, open-market only) ──────────────────────
    cutoff_ins = (datetime.date.today() - datetime.timedelta(days=90)).isoformat()
    ins_raw = _get(f"insider-trading?symbol={ticker}&limit=50")
    buys = sells = 0
    buy_usd = 0.0
    if isinstance(ins_raw, list):
        for e in ins_raw:
            if (e.get("transactionDate") or "") < cutoff_ins:
                continue
            ttype = e.get("transactionType") or ""
            if ttype == "P-Purchase":
                buys += 1
                buy_usd += float(e.get("value") or 0)
            elif ttype == "S-Sale":
                sells += 1
    data["insider_buys"]    = buys
    data["insider_buy_usd"] = buy_usd
    data["insider_sells"]   = sells

    # ── Options activity (yfinance — no free alternative for options) ─────────
    data["put_call_ratio"] = None
    data["call_oi"]        = None
    data["put_oi"]         = None
    try:
        import yfinance as yf
        yt   = yf.Ticker(ticker)
        exps = yt.options
        if exps:
            chain   = yt.option_chain(exps[0])
            call_oi = chain.calls["openInterest"].sum()
            put_oi  = chain.puts["openInterest"].sum()
            data["call_oi"]        = int(call_oi)
            data["put_oi"]         = int(put_oi)
            data["put_call_ratio"] = (put_oi / call_oi) if call_oi else None
    except Exception as e:
        data["options_error"] = str(e)

    # ── News sentiment — Polygon NLP (replaces yfinance keyword scraping) ─────
    pn = _fetch_polygon_news(ticker, polygon_key)
    data["news_count"]       = pn.get("news_count")
    data["news_7d"]          = pn.get("news_7d")
    data["news_bullish"]     = pn.get("news_bullish")
    data["news_bearish"]     = pn.get("news_bearish")
    data["news_bullish_pct"] = pn.get("news_bullish_pct")

    # ── Reddit sentiment (PRAW — requires REDDIT_CLIENT_ID + REDDIT_CLIENT_SECRET) ─
    rd = _fetch_reddit_sentiment(ticker)
    data["reddit_mentions"]    = rd.get("reddit_mentions")
    data["reddit_24h"]         = rd.get("reddit_24h")
    data["reddit_bullish"]     = rd.get("reddit_bullish")
    data["reddit_bearish"]     = rd.get("reddit_bearish")
    data["reddit_bullish_pct"] = rd.get("reddit_bullish_pct")

    return data


# ── Scoring functions ──────────────────────────────────────────────────────────

def _score_momentum(data: dict) -> tuple[str, int, str]:
    rsi   = data.get("rsi_14")
    vs50  = data.get("price_vs_50ma")
    vs200 = data.get("price_vs_200ma")

    if rsi is None:
        return "MOD", 5, "RSI unavailable — defaulting to neutral"

    above_50  = vs50  is not None and vs50  > 0
    above_200 = vs200 is not None and vs200 > 0

    # Ideal speculative zone: RSI 50-75, above both MAs (trending, not extended)
    if 50 <= rsi <= 75 and above_50 and above_200:
        return "HIGH", 10, f"RSI {rsi:.0f} in momentum zone; +{vs50*100:.0f}% vs 50d MA, +{vs200*100:.0f}% vs 200d MA"
    if rsi > 75 and above_50 and above_200:
        return "MOD", 5, f"RSI {rsi:.0f} extended (overbought risk) — strong trend but pullback possible"
    if 40 <= rsi < 50 and (above_50 or above_200):
        return "MOD", 5, f"RSI {rsi:.0f} recovering — early-stage setup, watch for breakout above 50"
    if rsi > 50 and (above_50 or above_200):
        return "MOD", 5, f"RSI {rsi:.0f} positive but only above one MA — partial momentum"
    return "LOW", 0, f"RSI {rsi:.0f} — no clear upward momentum, price below key moving averages"


def _score_volume(data: dict) -> tuple[str, int, str]:
    v10 = data.get("vol_10d_avg")
    v50 = data.get("vol_50d_avg")

    if not v10 or not v50 or v50 == 0:
        return "MOD", 5, "Volume data unavailable — defaulting to neutral"

    ratio = v10 / v50
    if ratio >= 1.75:
        return "HIGH", 10, f"10d avg volume is {ratio:.1f}x the 50d baseline — strong institutional accumulation signal"
    if ratio >= 1.20:
        return "MOD", 5, f"10d avg volume is {ratio:.1f}x the 50d baseline — elevated but not unusual"
    return "LOW", 0, f"10d avg volume is {ratio:.1f}x the 50d baseline — no unusual volume activity"


def _score_short_interest(data: dict) -> tuple[str, int, str]:
    sf  = data.get("short_float_pct")
    dtc = data.get("days_to_cover")

    if sf is None:
        return "MOD", 5, "Short interest data unavailable — defaulting to neutral"

    if sf >= 15 and (dtc is None or dtc >= 3):
        dtc_str = f", {dtc:.1f} days to cover" if dtc else ""
        return "HIGH", 10, f"{sf:.1f}% short float{dtc_str} — significant short squeeze fuel"
    if sf >= 8:
        dtc_str = f", {dtc:.1f} days to cover" if dtc else ""
        return "MOD", 5, f"{sf:.1f}% short float{dtc_str} — moderate short interest"
    return "LOW", 0, f"{sf:.1f}% short float — low short interest, limited squeeze potential"


def _score_analyst_revisions(data: dict) -> tuple[str, int, str]:
    ups   = data.get("upgrades_90d", 0) or 0
    downs = data.get("downgrades_90d", 0) or 0
    rev   = data.get("est_rev_pct")

    total = ups + downs
    if total == 0 and rev is None:
        return "MOD", 5, "No recent analyst actions or estimate data available"

    upgrade_bias = ups > downs
    strong_rev   = rev is not None and rev > 0.05    # >5% revision up
    any_neg_rev  = rev is not None and rev < -0.05

    if upgrade_bias and strong_rev:
        rev_str = f"+{rev*100:.0f}%" if rev else ""
        return "HIGH", 10, f"{ups} upgrades vs {downs} downgrades (90d); estimates revised {rev_str} — strong positive momentum"
    if upgrade_bias or strong_rev:
        rev_str = f"{rev*100:+.0f}%" if rev else ""
        return "MOD", 5, f"{ups} upgrades vs {downs} downgrades (90d){'; estimates ' + rev_str if rev else ''}"
    if any_neg_rev or downs > ups:
        rev_str = f"{rev*100:+.0f}%" if rev else ""
        return "LOW", 0, f"{ups} upgrades vs {downs} downgrades (90d){'; estimates ' + rev_str if rev else ''} — negative analyst momentum"
    return "MOD", 5, f"{ups} upgrades vs {downs} downgrades (90d) — neutral analyst activity"


def _score_float(data: dict) -> tuple[str, int, str]:
    mktcap      = data.get("market_cap")
    float_sh    = data.get("float_shares")
    shares_out  = data.get("shares_out")
    cur_price   = data.get("current_price")

    float_val = None
    if float_sh and cur_price:
        float_val = float_sh * cur_price

    # Use market cap as primary signal (easier to get consistently)
    if mktcap:
        mktcap_b = mktcap / 1e9
        if mktcap_b < 1.5:
            size_tier = "HIGH"
            size_note = f"${mktcap_b:.1f}B market cap (micro/small cap) — high potential for explosive moves"
        elif mktcap_b < 10:
            size_tier = "MOD"
            size_note = f"${mktcap_b:.1f}B market cap (small/mid cap) — meaningful but not micro-cap explosive"
        else:
            size_tier = "LOW"
            size_note = f"${mktcap_b:.1f}B market cap (large cap) — needs massive catalysts to move significantly"

        if float_val:
            float_b = float_val / 1e9
            size_note += f"; float ~${float_b:.1f}B"
        return size_tier, TIER_PTS[size_tier], size_note

    return "MOD", 5, "Market cap data unavailable — defaulting to neutral"


def _score_insider(data: dict) -> tuple[str, int, str]:
    buys    = data.get("insider_buys", 0) or 0
    buy_usd = data.get("insider_buy_usd", 0) or 0
    sells   = data.get("insider_sells", 0) or 0

    if buys == 0 and sells == 0:
        return "MOD", 5, "No insider transactions in the last 90 days"
    if buys >= 2 and buy_usd >= 200_000 and buys > sells:
        return "HIGH", 10, f"{buys} open-market insider purchase(s), ${buy_usd/1e3:.0f}K total (90d) — strong smart-money signal"
    if buys >= 1 and buys >= sells:
        return "MOD", 5, f"{buys} insider purchase(s), ${buy_usd/1e3:.0f}K total vs {sells} sale(s) — modest positive signal"
    return "LOW", 0, f"Net insider selling ({sells} sales vs {buys} buys, 90d) — negative signal"


def _score_downside_floor(data: dict) -> tuple[str, int, str]:
    net_debt  = data.get("net_debt")     # positive = net debt, negative = net cash
    ebitda    = data.get("ebitda_ttm")
    cash      = data.get("cash", 0) or 0
    tot_debt  = data.get("total_debt", 0) or 0

    if net_debt is None:
        return "MOD", 5, "Balance sheet data unavailable"

    if net_debt <= 0:
        # Net cash position
        net_cash_m = abs(net_debt) / 1e6
        return "HIGH", 10, f"Net cash of ${net_cash_m:.0f}M — downside protected, no debt cliff risk"

    # Net debt — check leverage
    if ebitda and ebitda > 0:
        lev = net_debt / ebitda
        if lev < 2.0:
            return "MOD", 5, f"Net debt ${net_debt/1e6:.0f}M, D/EBITDA {lev:.1f}x — manageable leverage"
        if lev < 4.0:
            return "LOW", 0, f"Net debt ${net_debt/1e6:.0f}M, D/EBITDA {lev:.1f}x — elevated leverage, catalyst must arrive before debt pressure"
        return "LOW", 0, f"Net debt ${net_debt/1e6:.0f}M, D/EBITDA {lev:.1f}x — high leverage, binary risk if catalyst fails"

    if tot_debt > cash * 3:
        return "LOW", 0, f"Debt ${tot_debt/1e6:.0f}M significantly exceeds cash ${cash/1e6:.0f}M — material downside risk"
    return "MOD", 5, f"Cash ${cash/1e6:.0f}M vs debt ${tot_debt/1e6:.0f}M — moderate balance sheet"


def _score_options(data: dict) -> tuple[str, int, str]:
    pc  = data.get("put_call_ratio")
    c_oi = data.get("call_oi")
    p_oi = data.get("put_oi")

    if pc is None:
        err = data.get("options_error", "")
        note = f"Options data unavailable — {err}" if err else "Options data unavailable (no listed options or data error)"
        return "MOD", 5, note

    if pc < 0.45:
        return "HIGH", 10, f"Put/Call ratio {pc:.2f} — heavy call-side positioning, smart money expressing bullish conviction"
    if pc < 0.75:
        return "MOD", 5, f"Put/Call ratio {pc:.2f} — moderately bullish options positioning"
    if pc > 1.20:
        return "LOW", 0, f"Put/Call ratio {pc:.2f} — put-heavy positioning, market hedging against downside"
    return "MOD", 5, f"Put/Call ratio {pc:.2f} — roughly neutral options positioning"


def _score_social_trend(data: dict) -> tuple[str, int, str]:
    """Score social/trend momentum across two sub-signals:

      1. Polygon news sentiment  — bullish vs bearish ratio from Polygon NLP (last 20 articles)
      2. Reddit sentiment        — mention count + bullish/bearish ratio across WSB/stocks/investing

    HIGH = both sub-signals bullish · MOD = 1 bullish or all unavailable · LOW = 0 bullish with ≥1 bearish
    """
    bullish_signals = []
    notes = []

    # Sub-signal 1: Polygon news sentiment
    news_bp    = data.get("news_bullish_pct")
    news_bull  = data.get("news_bullish", 0) or 0
    news_bear  = data.get("news_bearish", 0) or 0
    news_count = data.get("news_count") or 0
    news_7d    = data.get("news_7d") or 0

    if news_bp is not None:
        activity = f"{news_7d} articles in last 7d" if news_7d else f"{news_count} total"
        if news_bp >= 0.65:
            bullish_signals.append(True)
            notes.append(f"News: {news_bp*100:.0f}% bullish ({news_bull}+ / {news_bear}−, {activity})")
        elif news_bp >= 0.40:
            notes.append(f"News: {news_bp*100:.0f}% bullish — mixed sentiment ({activity})")
        else:
            bullish_signals.append(False)
            notes.append(f"News: {news_bp*100:.0f}% bullish — bearish headline bias ({activity})")
    elif news_count == 0:
        notes.append("News: no recent articles found")
    else:
        notes.append("News: unavailable")

    # Sub-signal 2: Reddit sentiment (WSB + stocks + investing) — INVERTED SCORING
    # Contrarian logic: extreme crowd bullishness on retail forums means you're late.
    # What we want is UNDER-THE-RADAR positioning, not confirmation that WSB already piled in.
    rd_bp    = data.get("reddit_bullish_pct")
    rd_bull  = data.get("reddit_bullish", 0) or 0
    rd_bear  = data.get("reddit_bearish", 0) or 0
    rd_total = data.get("reddit_mentions") or 0
    rd_24h   = data.get("reddit_24h") or 0

    if rd_bp is not None:
        activity = f"{rd_24h} posts in last 24h" if rd_24h else f"{rd_total} posts this week"
        if rd_bp >= 0.75 and rd_total >= 15:
            # Crowded retail trade — late entry risk
            bullish_signals.append(False)
            notes.append(f"Reddit: {rd_bp*100:.0f}% bullish ({rd_bull}+ / {rd_bear}−, {activity}) — crowded retail trade, asymmetry eroded")
        elif rd_total < 8:
            # Under the radar — early positioning opportunity
            bullish_signals.append(True)
            notes.append(f"Reddit: low visibility ({rd_total} mentions this week) — not yet retail-crowded, potential early setup")
        elif rd_bp < 0.35:
            # Crowd is skeptical — confirms weak setup signal
            bullish_signals.append(False)
            notes.append(f"Reddit: {rd_bp*100:.0f}% bullish — crowd skeptical ({activity})")
        else:
            # Moderate interest, mixed sentiment — neutral; don't penalise or reward
            notes.append(f"Reddit: {rd_bp*100:.0f}% bullish — moderate interest ({activity}), no extreme crowding")
    elif rd_total == 0:
        # No mentions at all — stock under the radar (positive for asymmetry)
        bullish_signals.append(True)
        notes.append("Reddit: no mentions found — not on retail radar, early positioning potential")
    else:
        notes.append("Reddit: credentials not configured")

    n_bullish = sum(1 for b in bullish_signals if b is True)
    n_bearish = sum(1 for b in bullish_signals if b is False)

    if n_bullish >= 2:
        tier, pts = "HIGH", 10
    elif n_bullish == 1:
        tier, pts = "MOD", 5
    elif n_bullish == 0 and n_bearish == 0:
        tier, pts = "MOD", 5   # all unavailable — neutral, don't penalise
    else:
        tier, pts = "LOW", 0

    return tier, pts, " · ".join(notes)


def _score_technical_analysis(data: dict) -> tuple[str, int, str]:
    """Score the technical setup using MACD crossover, 52w high proximity, and EMA stack.

    Each of the three sub-signals votes bullish/neutral/bearish.
    HIGH requires 2-3 bullish votes. MOD requires 1. LOW requires 0.
    """
    bullish_signals = []
    notes = []

    # Sub-signal 1: MACD — histogram positive AND MACD above signal = bullish crossover
    macd      = data.get("macd")
    macd_sig  = data.get("macd_signal")
    macd_hist = data.get("macd_hist")

    if macd is not None and macd_sig is not None and macd_hist is not None:
        if macd_hist > 0 and macd > macd_sig:
            bullish_signals.append(True)
            notes.append(f"MACD bullish (hist +{macd_hist:.2f}, MACD {macd:.2f} > signal {macd_sig:.2f})")
        elif macd_hist < 0:
            bullish_signals.append(False)
            notes.append(f"MACD bearish (hist {macd_hist:.2f})")
        else:
            notes.append("MACD neutral")
    else:
        notes.append("MACD unavailable")

    # Sub-signal 2: 52-week high proximity
    pct_off = data.get("pct_from_52w_high")
    high_52w = data.get("high_52w")

    if pct_off is not None:
        if pct_off >= -0.08:
            # Within 8% of 52w high — breakout zone
            bullish_signals.append(True)
            notes.append(f"{pct_off*100:.1f}% from 52w high (${high_52w:.2f}) — breakout watch zone")
        elif pct_off >= -0.25:
            # 8-25% below — neutral; not deeply buried
            notes.append(f"{pct_off*100:.1f}% from 52w high — intermediate range")
        else:
            # >25% below 52w high — stock deeply buried, needs big catalyst
            bullish_signals.append(False)
            notes.append(f"{pct_off*100:.1f}% from 52w high — stock deeply off highs")
    else:
        notes.append("52w high data unavailable")

    # Sub-signal 3: EMA stack — price > EMA21 > EMA50 = bullish alignment
    price = data.get("current_price")
    ema21 = data.get("ema_21")
    ema50 = data.get("ema_50")

    if price and ema21 and ema50:
        if price > ema21 > ema50:
            bullish_signals.append(True)
            notes.append(f"EMA stack bullish: price ${price:.2f} > EMA21 ${ema21:.2f} > EMA50 ${ema50:.2f}")
        elif price < ema21 or ema21 < ema50:
            bullish_signals.append(False)
            notes.append(f"EMA stack bearish: price ${price:.2f} / EMA21 ${ema21:.2f} / EMA50 ${ema50:.2f}")
        else:
            notes.append(f"EMA mixed: price ${price:.2f} / EMA21 ${ema21:.2f} / EMA50 ${ema50:.2f}")
    else:
        notes.append("EMA data unavailable")

    n_bullish = sum(1 for b in bullish_signals if b)
    n_bearish = sum(1 for b in bullish_signals if not b)

    if n_bullish >= 2:
        tier = "HIGH"
        pts  = 10
    elif n_bullish == 1 and n_bearish == 0:
        tier = "MOD"
        pts  = 5
    elif n_bullish == 1 and n_bearish == 1:
        tier = "MOD"
        pts  = 5
    elif n_bearish >= 2:
        tier = "LOW"
        pts  = 0
    else:
        tier = "MOD"
        pts  = 5

    return tier, pts, " · ".join(notes)


def _score_narrative(narrative_theme: str, narrative_strength: str) -> tuple[str, int, str]:
    """Narrative theme scored from user-supplied dropdown + strength tier."""
    if not narrative_theme or narrative_theme == "None":
        return "LOW", 0, "No narrative theme identified — speculative plays without a story rarely outperform"

    strength_pts = {"HIGH": 10, "MOD": 5, "LOW": 0}.get(narrative_strength.upper() if narrative_strength else "", 5)

    # Hot themes get a boost to HIGH more easily
    hot_themes = {"AI / Machine Learning", "Defence & Aerospace", "GLP-1 / Weight Loss",
                  "Nuclear / SMR", "Crypto-Adjacent", "Biotech Catalyst"}

    if strength_pts == 10:
        tier = "HIGH"
    elif strength_pts == 5:
        tier = "MOD"
    else:
        tier = "LOW"

    # Override tier upward if it's a hot theme with at least MOD strength
    if narrative_theme in hot_themes and strength_pts >= 5:
        tier = "HIGH"
        strength_pts = 10

    note = f"Theme: {narrative_theme} — {narrative_strength or 'MOD'} narrative connection to company"
    return tier, strength_pts, note


def _score_catalyst(catalyst_desc: str, catalyst_timing: str) -> tuple[str, int, str]:
    """Catalyst quality scored from user-supplied description and timing tier."""
    if not catalyst_desc or len(catalyst_desc.strip()) < 5:
        return "LOW", 0, "No specific catalyst identified — without a trigger, timing a move is speculative at best"

    timing_pts = {"near": 10, "medium": 5, "vague": 0}.get((catalyst_timing or "vague").lower(), 0)

    if timing_pts == 10:
        tier = "HIGH"
        note = f"Near-term catalyst: {catalyst_desc[:120]} — specific event with clear timing"
    elif timing_pts == 5:
        tier = "MOD"
        note = f"Medium-term catalyst: {catalyst_desc[:120]} — expected but timing uncertain"
    else:
        tier = "LOW"
        note = f"Vague catalyst: {catalyst_desc[:120]} — no clear timing anchor"

    return tier, timing_pts, note


# ── Main scorecard builder ─────────────────────────────────────────────────────

def build_speculative_scorecard(
    data: dict,
    narrative_theme: str   = "",
    narrative_strength: str = "MOD",
    catalyst_desc: str     = "",
    catalyst_timing: str   = "vague",    # "near" | "medium" | "vague"
) -> dict:
    """Score all 10 dimensions and return the full scorecard metrics dict."""

    scores = {}

    t, p, n = _score_momentum(data)
    scores["momentum"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_volume(data)
    scores["volume"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_short_interest(data)
    scores["short_interest"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_analyst_revisions(data)
    scores["analyst_revisions"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_float(data)
    scores["float_size"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_insider(data)
    scores["insider_buying"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_downside_floor(data)
    scores["downside_floor"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_options(data)
    scores["options_activity"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_technical_analysis(data)
    scores["technical_setup"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_social_trend(data)
    scores["social_trend"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_narrative(narrative_theme, narrative_strength)
    scores["narrative"] = {"tier": t, "pts": p, "note": n}

    t, p, n = _score_catalyst(catalyst_desc, catalyst_timing)
    scores["catalyst"] = {"tier": t, "pts": p, "note": n}

    total = sum(s["pts"] for s in scores.values())
    verdict = _verdict(total)

    # Stamp all auto-scored signals with the data date so readers know signal freshness
    _as_of = datetime.date.today().isoformat()
    _manual_keys = {"narrative", "catalyst"}
    for key, s in scores.items():
        if key not in _manual_keys and s.get("note"):
            s["note"] = s["note"] + f"  [data: {_as_of}]"

    return {
        "scores":            scores,
        "total_score":       total,
        "verdict":           verdict,
        "narrative_theme":   narrative_theme,
        "narrative_strength": narrative_strength,
        "catalyst_desc":     catalyst_desc,
        "catalyst_timing":   catalyst_timing,
    }


# ── Scenario / re-rating model ─────────────────────────────────────────────────

def build_scenario_model(data: dict, scorecard: dict, hold_months: int = 6) -> dict:
    """Build a bear/base/bull scenario model based on EV/Revenue re-rating.

    Uses forward revenue estimates and scenarios for EV/Revenue multiple expansion
    to project a target price in the hold_months timeframe.
    Returns a dict with all scenario inputs and outputs.
    """
    price     = data.get("current_price") or 0
    mktcap    = data.get("market_cap")    or 0
    net_debt  = data.get("net_debt")      or 0
    shares    = data.get("shares_out")    or 0
    trail_rev = data.get("trailing_rev")  or 0
    fwd_rev   = data.get("fwd_rev_1yr")   or trail_rev
    cur_ev    = data.get("current_ev")    or mktcap + net_debt

    # Current EV/fwd revenue multiple
    cur_ev_rev_mult = (cur_ev / fwd_rev) if fwd_rev else None

    # Scenario multiplier assumptions
    # We calibrate around the current multiple and the narrative theme
    total_score = scorecard.get("total_score", 50)

    if cur_ev_rev_mult and cur_ev_rev_mult > 0:
        bear_mult = cur_ev_rev_mult * 0.65   # narrative fails → de-rating
        base_mult = cur_ev_rev_mult * 1.00   # stays flat, only rev growth drives return
        # Bull multiple expansion: function of how high-conviction the setup is
        if total_score >= 75:
            bull_factor = 2.20
        elif total_score >= 60:
            bull_factor = 1.70
        elif total_score >= 45:
            bull_factor = 1.40
        else:
            bull_factor = 1.20
        bull_mult = cur_ev_rev_mult * bull_factor
    else:
        # No EV/Rev data — use simple price-based scenarios
        cur_ev_rev_mult = None
        bear_mult = base_mult = bull_mult = None

    def _price_from_mult(mult):
        if not mult or not fwd_rev or not shares:
            return None
        target_ev     = mult * fwd_rev
        target_equity = target_ev - net_debt
        if target_equity <= 0:
            return None
        return target_equity / shares

    bear_price = _price_from_mult(bear_mult)
    base_price = _price_from_mult(base_mult)
    bull_price = _price_from_mult(bull_mult)

    # Fallback: simple % moves if EV/Rev unavailable
    if bear_price is None and price:
        bear_price = price * 0.65
        base_price = price * 1.20
        bull_price = price * 1.75
        bear_mult = base_mult = bull_mult = None

    def _ret(tp):
        if tp is None or not price:
            return None
        return (tp / price) - 1

    bear_ret = _ret(bear_price)
    base_ret = _ret(base_price)
    bull_ret = _ret(bull_price)

    # Target price to achieve 1.5x
    target_1_5x_price = (price * 1.5) if price else None

    # What multiple is required for 1.5x?
    req_mult_for_1_5x = None
    if price and shares and fwd_rev:
        target_equity_1_5x = target_1_5x_price * shares
        target_ev_1_5x     = target_equity_1_5x + net_debt
        req_mult_for_1_5x  = target_ev_1_5x / fwd_rev if fwd_rev else None

    return {
        "hold_months":          hold_months,
        "current_price":        price,
        "fwd_rev_b":            fwd_rev / 1e9 if fwd_rev else None,
        "trail_rev_b":          trail_rev / 1e9 if trail_rev else None,
        "current_ev_b":         cur_ev / 1e9 if cur_ev else None,
        "current_ev_rev_mult":  cur_ev_rev_mult,
        # Scenario multiples
        "bear_mult":            bear_mult,
        "base_mult":            base_mult,
        "bull_mult":            bull_mult,
        # Scenario prices
        "bear_price":           bear_price,
        "base_price":           base_price,
        "bull_price":           bull_price,
        # Scenario returns
        "bear_ret":             bear_ret,
        "base_ret":             base_ret,
        "bull_ret":             bull_ret,
        # 1.5x analysis
        "target_1_5x_price":    target_1_5x_price,
        "req_mult_for_1_5x":    req_mult_for_1_5x,
        "bull_reaches_1_5x":    bull_price is not None and price > 0 and bull_price >= price * 1.5,
    }


# ── Excel workbook builder ─────────────────────────────────────────────────────

def build_speculative_excel(
    ticker: str,
    data: dict,
    scorecard: dict,
    scenario: dict,
) -> bytes:
    """Build a simple 3-tab Excel workbook for the speculative analysis."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import (
            Font, PatternFill, Alignment, Border, Side, numbers
        )
        from openpyxl.utils import get_column_letter
    except ImportError:
        return b""

    wb = Workbook()

    # ── Colour palette ────────────────────────────────────────────────────────
    BG_DARK    = "FF0F0F0F"
    BG_PANEL   = "FF1A1A2E"
    BG_HEADER  = "FF16213E"
    ORANGE     = "FFFF6B35"
    AMBER      = "FFFFA62B"
    GREEN      = "FF34D399"
    RED        = "FFEF4444"
    GOLD       = "FFE6C168"
    WHITE      = "FFFFFFFF"
    GREY       = "FF9CA3AF"

    def _fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def _font(bold=False, color=WHITE, size=11):
        return Font(bold=bold, color=color, name="Calibri", size=size)

    def _border():
        s = Side(style="thin", color="FF2D2D2D")
        return Border(left=s, right=s, top=s, bottom=s)

    def _hdr(ws, row, col, val, bold=True, bg=BG_HEADER, fg=WHITE, size=11):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = _font(bold=bold, color=fg, size=size)
        c.fill      = _fill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = _border()
        return c

    def _cell(ws, row, col, val, bold=False, bg=BG_PANEL, fg=WHITE, num_fmt=None, align="left"):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = _font(bold=bold, color=fg, size=10)
        c.fill      = _fill(bg)
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border    = _border()
        if num_fmt:
            c.number_format = num_fmt
        return c

    tier_color = {"HIGH": GREEN, "MOD": AMBER, "LOW": RED}

    # ═══════════════════════════════════════════════════════════════════════════
    # Sheet 1 — Speculative Scorecard
    # ═══════════════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "Speculative Scorecard"
    ws1.sheet_view.showGridLines = False
    ws1.column_dimensions["A"].width = 26
    ws1.column_dimensions["B"].width = 12
    ws1.column_dimensions["C"].width = 10
    ws1.column_dimensions["D"].width = 65

    # Title
    ws1.row_dimensions[1].height = 40
    c = ws1.cell(row=1, column=1,
                 value=f"MAYBE UNSAFE FINANCIAL FREEDOM — {ticker}")
    c.font      = Font(bold=True, color=ORANGE, size=16, name="Calibri")
    c.fill      = _fill(BG_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws1.merge_cells("A1:D1")

    # Sub-title
    ws1.row_dimensions[2].height = 20
    c = ws1.cell(row=2, column=1, value=f"Speculative Signal Scorecard · Generated {datetime.date.today()}")
    c.font      = Font(color=GREY, size=9, name="Calibri")
    c.fill      = _fill(BG_DARK)
    ws1.merge_cells("A2:D2")

    # Column headers
    ws1.row_dimensions[3].height = 22
    _hdr(ws1, 3, 1, "Signal", bg=BG_HEADER, fg=GOLD)
    _hdr(ws1, 3, 2, "Tier",   bg=BG_HEADER, fg=GOLD)
    _hdr(ws1, 3, 3, "Points", bg=BG_HEADER, fg=GOLD)
    _hdr(ws1, 3, 4, "Rationale", bg=BG_HEADER, fg=GOLD)

    LABELS = {
        "momentum":          "Price Momentum",
        "volume":            "Volume Signal",
        "short_interest":    "Short Interest / Squeeze",
        "analyst_revisions": "Analyst Revision Momentum",
        "float_size":        "Float Size / Market Cap",
        "insider_buying":    "Insider Buying (90d)",
        "downside_floor":    "Downside Floor",
        "options_activity":  "Options Activity (P/C)",
        "technical_setup":   "Technical Setup",
        "social_trend":      "Social / Trend Momentum",
        "narrative":         "Narrative Theme",
        "catalyst":          "Catalyst Quality",
    }

    r = 4
    scores = scorecard.get("scores", {})
    for key, label in LABELS.items():
        s = scores.get(key, {})
        tier  = s.get("tier", "MOD")
        pts   = s.get("pts",  5)
        note  = s.get("note", "")
        tc    = tier_color.get(tier, AMBER)
        ws1.row_dimensions[r].height = 32
        _cell(ws1, r, 1, label, bold=True,  fg=WHITE)
        c2 = _cell(ws1, r, 2, tier, bold=True, fg=tc, align="center")
        _cell(ws1, r, 3, pts,  bold=True,  fg=WHITE, num_fmt="0", align="center")
        c4 = _cell(ws1, r, 4, note, fg=GREY)
        c4.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        r += 1

    # Totals row
    ws1.row_dimensions[r].height = 28
    total  = scorecard.get("total_score", 0)
    verdict= scorecard.get("verdict", "")
    _hdr(ws1, r, 1, "TOTAL SCORE", bg="FF1F1F1F", fg=GOLD)
    _hdr(ws1, r, 2, verdict,       bg="FF1F1F1F", fg=ORANGE)
    _hdr(ws1, r, 3, total,         bg="FF1F1F1F", fg=WHITE)
    _hdr(ws1, r, 4, "Score out of 120 · Verdict: ≥90 Moonshot · ≥65 Strong Spec · ≥54 Spec Play · ≥36 High Risk · <36 Pass",
         bg="FF1F1F1F", fg=GREY)

    # ═══════════════════════════════════════════════════════════════════════════
    # Sheet 2 — Scenario Model
    # ═══════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Scenario Model")
    ws2.sheet_view.showGridLines = False
    for col, w in [("A", 28), ("B", 18), ("C", 18), ("D", 18)]:
        ws2.column_dimensions[col].width = w

    ws2.row_dimensions[1].height = 40
    c = ws2.cell(row=1, column=1, value=f"RE-RATING SCENARIO MODEL — {ticker}")
    c.font      = Font(bold=True, color=ORANGE, size=16, name="Calibri")
    c.fill      = _fill(BG_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws2.merge_cells("A1:D1")

    sc = scenario
    hm = sc.get("hold_months", 6)

    def _kv(ws, row, label, val, val_fmt=None, val_color=WHITE):
        ws.row_dimensions[row].height = 22
        _cell(ws, row, 1, label, bold=True, fg=GOLD)
        c = _cell(ws, row, 2, val, fg=val_color, align="right")
        if val_fmt:
            c.number_format = val_fmt
        ws.merge_cells(f"B{row}:D{row}")

    r2 = 2
    _hdr(ws2, r2, 1, "MODEL INPUTS", bg=BG_HEADER, fg=GOLD)
    ws2.merge_cells(f"A{r2}:D{r2}")
    r2 += 1
    _kv(ws2, r2, "Current Price",         sc.get("current_price"),     '"$"#,##0.00'); r2 += 1
    _kv(ws2, r2, "Current EV ($B)",       sc.get("current_ev_b"),      '#,##0.00'); r2 += 1
    _kv(ws2, r2, "Trailing Revenue ($B)", sc.get("trail_rev_b"),       '#,##0.00'); r2 += 1
    _kv(ws2, r2, "Fwd Revenue Est. ($B)", sc.get("fwd_rev_b"),         '#,##0.00'); r2 += 1
    _kv(ws2, r2, "Current EV/Fwd Rev",    sc.get("current_ev_rev_mult"), '#,##0.0"x"'); r2 += 1
    _kv(ws2, r2, f"Hold Horizon (months)",hm,                          '0 "months"'); r2 += 1

    r2 += 1
    # Scenario table header
    ws2.row_dimensions[r2].height = 22
    _hdr(ws2, r2, 1, "SCENARIO",         bg=BG_HEADER, fg=GOLD)
    _hdr(ws2, r2, 2, "EV/Rev Multiple",  bg=BG_HEADER, fg=GOLD)
    _hdr(ws2, r2, 3, "Target Price",     bg=BG_HEADER, fg=GOLD)
    _hdr(ws2, r2, 4, "Return",           bg=BG_HEADER, fg=GOLD)
    r2 += 1

    scenarios_data = [
        ("BEAR — Narrative Fails",    sc.get("bear_mult"), sc.get("bear_price"), sc.get("bear_ret"), RED),
        ("BASE — Revenue Grows Only", sc.get("base_mult"), sc.get("base_price"), sc.get("base_ret"), AMBER),
        ("BULL — Narrative Plays Out",sc.get("bull_mult"), sc.get("bull_price"), sc.get("bull_ret"), GREEN),
    ]
    for label, mult, tprice, tret, color in scenarios_data:
        ws2.row_dimensions[r2].height = 24
        _cell(ws2, r2, 1, label, bold=True, fg=color)
        _cell(ws2, r2, 2, mult,   fg=WHITE, align="right", num_fmt='#,##0.0"x"')
        _cell(ws2, r2, 3, tprice, fg=WHITE, align="right", num_fmt='"$"#,##0.00')
        c = _cell(ws2, r2, 4, tret,   fg=color, align="right", num_fmt='0.0%')
        r2 += 1

    r2 += 1
    ws2.row_dimensions[r2].height = 24
    _cell(ws2, r2, 1, "1.5× Target Price", bold=True, fg=GOLD)
    _cell(ws2, r2, 2, sc.get("req_mult_for_1_5x"), fg=AMBER, align="right", num_fmt='#,##0.0"x"')
    _cell(ws2, r2, 3, sc.get("target_1_5x_price"),  fg=AMBER, align="right", num_fmt='"$"#,##0.00')
    c = _cell(ws2, r2, 4, "Required for 1.5× return", fg=GREY)
    r2 += 2

    note_txt = ("NOTE: Scenario prices derived from EV/Forward Revenue re-rating. "
                "Bear assumes narrative fails and multiple de-rates. "
                "Base assumes flat multiple with revenue growth only. "
                "Bull assumes narrative takes hold and multiple expands. "
                "All scenarios are illustrative. Adjust multiples as your thesis evolves.")
    ws2.row_dimensions[r2].height = 50
    c = ws2.cell(row=r2, column=1, value=note_txt)
    c.font      = Font(color=GREY, size=9, name="Calibri", italic=True)
    c.fill      = _fill(BG_DARK)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws2.merge_cells(f"A{r2}:D{r2}")

    # ═══════════════════════════════════════════════════════════════════════════
    # Sheet 3 — Market Data
    # ═══════════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Market Data")
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 20

    ws3.row_dimensions[1].height = 40
    c = ws3.cell(row=1, column=1, value=f"RAW MARKET SIGNALS — {ticker}")
    c.font      = Font(bold=True, color=ORANGE, size=16, name="Calibri")
    c.fill      = _fill(BG_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws3.merge_cells("A1:B1")

    r3 = 2
    raw_fields = [
        ("Company Name",          data.get("company_name"),          None),
        ("Sector",                data.get("sector"),                 None),
        ("Current Price",         data.get("current_price"),          '"$"#,##0.00'),
        ("Market Cap ($B)",       (data.get("market_cap") or 0)/1e9, '#,##0.00'),
        ("Float Shares (M)",      (data.get("float_shares") or 0)/1e6, '#,##0.0'),
        ("RSI (14)",              data.get("rsi_14"),                 '#,##0.0'),
        ("Price vs 50d MA",       data.get("price_vs_50ma"),          '0.0%'),
        ("Price vs 200d MA",      data.get("price_vs_200ma"),         '0.0%'),
        ("10d Avg Volume (M)",    (data.get("vol_10d_avg") or 0)/1e6, '#,##0.00'),
        ("50d Avg Volume (M)",    (data.get("vol_50d_avg") or 0)/1e6, '#,##0.00'),
        ("Short Float %",         data.get("short_float_pct"),        '0.0"%"'),
        ("Days to Cover",         data.get("days_to_cover"),          '#,##0.0'),
        ("Upgrades (90d)",        data.get("upgrades_90d"),           '0'),
        ("Downgrades (90d)",      data.get("downgrades_90d"),         '0'),
        ("Est. Revision %",       data.get("est_rev_pct"),            '0.0%'),
        ("Insider Buys (90d)",    data.get("insider_buys"),           '0'),
        ("Insider Buy Value ($K)",(data.get("insider_buy_usd") or 0)/1e3, '#,##0'),
        ("Insider Sells (90d)",   data.get("insider_sells"),          '0'),
        ("Put/Call Ratio",        data.get("put_call_ratio"),         '#,##0.00'),
        ("Call Open Interest",    data.get("call_oi"),                '#,##0'),
        ("Put Open Interest",     data.get("put_oi"),                 '#,##0'),
        ("Cash ($M)",             (data.get("cash") or 0)/1e6,        '#,##0'),
        ("Total Debt ($M)",       (data.get("total_debt") or 0)/1e6,  '#,##0'),
        ("Net Debt ($M)",         (data.get("net_debt") or 0)/1e6,    '#,##0'),
        ("Trailing Revenue ($M)", (data.get("trailing_rev") or 0)/1e6,'#,##0'),
        ("Fwd Revenue Est. ($M)", (data.get("fwd_rev_1yr") or 0)/1e6, '#,##0'),
        ("EBITDA TTM ($M)",       (data.get("ebitda_ttm") or 0)/1e6,  '#,##0'),
        ("Data Source",           "FMP API + yfinance" + (" [MOCK]" if data.get("mock") else ""), None),
    ]

    for label, val, fmt in raw_fields:
        ws3.row_dimensions[r3].height = 20
        _cell(ws3, r3, 1, label, bold=True, fg=GOLD)
        c = _cell(ws3, r3, 2, val, fg=WHITE, align="right")
        if fmt and val is not None:
            c.number_format = fmt
        r3 += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
