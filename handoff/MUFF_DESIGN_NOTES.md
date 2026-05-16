# MUFF — Maybe Unsafe Financial Freedom: Design Notes & Session Log

_Last updated: 2026-05-17_

This document captures all design decisions, rationale, API gotchas, and non-obvious context from the two sessions in which the MUFF speculative tool was built. It is intended to be the definitive reference for anyone resuming work on this feature.

---

## 1. What MUFF Is

MUFF is a second tool on the Investment Utopia Render site, sitting alongside the existing "Safe Financial Freedom" DCF/value tool. It targets **speculative plays with 1.5× or greater return potential within a 3–12 month hold horizon**, using momentum and narrative signals rather than long-term DCF fundamentals.

The user population is the same as the DCF tool: professional-level investors who will immediately notice a wrong number or lazy logic. All language and scoring rationale should be written at Bloomberg/CFA reader level.

---

## 2. Navigation & Branding Split

- Existing tool renamed: **"Safe Financial Freedom"** — eyebrow says "Long-Term Value Investing"
- New tool: **"Maybe Unsafe Financial Freedom (MUFF)"** — intentionally self-aware about the risk profile
- Nav bar has two links: "Safe" anchors to `#generate`, "Unsafe" (orange) anchors to `#speculative`
- The two tools are sections on the same `index.html` page — no separate routes needed

---

## 3. Scorecard Architecture

### 3.1 Scoring Scale

```
HIGH = 10 pts   MOD = 5 pts   LOW = 0 pts
Max score = 120 (12 dimensions × 10)
```

### 3.2 Verdict Thresholds (75% / 60% / 45% / 30% of 120)

| Score | Verdict |
|---|---|
| ≥ 90 | Moonshot Conviction |
| ≥ 72 | Strong Speculative |
| ≥ 54 | Speculative Play |
| ≥ 36 | High Risk |
| < 36 | Pass |

### 3.3 The 12 Dimensions

**Auto-scored (10 signals from live APIs):**

| # | Dimension | Data Source |
|---|---|---|
| 1 | Price Momentum | FMP: RSI-14 + price vs 50d/200d MA |
| 2 | Volume Signal | FMP: 10d avg vs 50d avg volume |
| 3 | Short Interest | FMP: short float % + days to cover |
| 4 | Analyst Revision Momentum | FMP: upgrades/downgrades (90d) + estimate revision % |
| 5 | Float Size / Market Cap | FMP: market cap in $B (proxy for explosive move potential) |
| 6 | Insider Buying (90d) | FMP: open-market P-Purchase transactions only |
| 7 | Downside Floor | FMP: net cash / D-EBITDA leverage |
| 8 | Options Activity | yfinance: put/call ratio on nearest expiry |
| 9 | Technical Setup | FMP: MACD crossover + 52w high proximity + EMA21/50 stack |
| 10 | Social / Trend Momentum | Google Trends + yfinance news + Reddit (PRAW) |

**Manual (2 signals from user inputs):**

| # | Dimension | Input |
|---|---|---|
| 11 | Narrative Theme | Dropdown (11 themes) + strength tier |
| 12 | Catalyst Quality | Free-text description + timing select |

### 3.4 Why Technical Setup Was Added

Originally not in scope, added after discussion. The logic: a speculative name needs both a narrative AND a technical setup that confirms price action. MACD crossover + 52w high proximity + EMA stack each vote bullish/neutral/bearish; 2-of-3 bullish = HIGH. This avoids catching falling knives where the story is real but the chart is broken.

### 3.5 Why Social / Trend Momentum Was Split Into Sub-Signals

The key insight from testing: a ticker's own Google Trends volume can be quiet while the **sector narrative** is building (e.g. "HBM memory" trending before MU gets noticed). This is an early entry signal, not a red flag. So the scoring treats:

- **Hot sector + quiet ticker = MOD (opportunity), not LOW**
- **Hot sector + hot ticker = counts as 2 bullish votes**

---

## 4. Valuation Model: EV/Revenue Re-Rating (not DCF)

MUFF uses a scenario model based on **EV/Forward Revenue multiple expansion** rather than DCF. Reasons:

1. Speculative plays are rarely valued on earnings — they're valued on narrative and re-rating
2. DCF requires steady-state assumptions that don't apply to 3–12 month trades
3. The question being answered is: "what multiple does the market need to assign for me to make 1.5×?"

### 4.1 Scenario Structure

| Scenario | Multiple | Assumption |
|---|---|---|
| Bear | Current × 0.65 | Narrative fails, de-rating |
| Base | Current × 1.00 | Flat multiple, only revenue growth drives return |
| Bull | Current × 1.20–2.20 | Multiple expansion, narrative plays out |

Bull factor scales with total score:
- ≥ 75 pts → 2.20× (highest conviction)
- ≥ 60 pts → 1.70×
- ≥ 45 pts → 1.40×
- < 45 pts → 1.20×

### 4.2 1.5× Analysis

The model explicitly computes:
- Target price for exactly 1.5× return
- Required EV/Rev multiple to reach that price
- Whether the bull scenario reaches 1.5×

This is the key question for the target user: is 1.5× achievable without requiring an absurd multiple?

---

## 5. Google Trends — Key Facts & Gotchas

### 5.1 What the Numbers Mean

Google Trends returns a **relative 0–100 index**, not absolute search volume. 100 = the peak day in the queried window. The numbers are **not comparable across different queries run at different times** — only within the same query are they directly comparable.

We run all keywords (ticker + theme + custom) in a **single pytrends query** (max 5 keywords) so they share the same 0–100 baseline and can be directly compared.

### 5.2 Ratio Calculation

We compute a `7-day vs prior-period ratio` from the returned series:
- `recent = last 7 values`
- `prior = all values before that`
- `ratio = avg(recent) / avg(prior)`

A ratio of 2.0× means the last 7 days average twice the prior-period average — a genuine surge signal.

### 5.3 Rate Limiting

Google aggressively rate-limits pytrends. Fix: `time.sleep(2)` before every `build_payload()` call.

**Do NOT pass `retries=` or `backoff_factor=` to `TrendReq()`** — current pytrends is incompatible with newer urllib3 for this parameter (`method_whitelist` kwarg was removed). The retries parameter was removed entirely from our init.

### 5.4 Phrase-Match, Not Exact

Search terms are phrase-matched, not exact. "HBM memory" will also capture "HBM memory chips" and related phrases. Single-word terms are too broad (e.g. "AI" captures everything). Use 2–3 word compound terms for signal quality.

### 5.5 Custom Search Terms (User-Supplied)

Users can enter comma-separated custom terms in the form (e.g. `HBM memory, DRAM cycle, memory bandwidth`). These are merged into the same pytrends query alongside the ticker and auto-selected theme keyword. Slot allocation: `[ticker, theme_kw, custom_1, custom_2, custom_3]` (5-slot cap).

Scoring uses the **best-performing non-ticker keyword** so one hot narrative is enough — you don't need all of them trending simultaneously.

### 5.6 Theme Keywords Map

Each narrative theme has 1–2 pre-set search terms. The rationale: these are sector-level terms that should trend before individual tickers react (lead indicator). Full map in `THEME_KEYWORDS` dict in `speculative_engine.py`.

---

## 6. Social Sentiment — Architecture & Data Sources

### 6.1 StockTwits — Abandoned

Originally planned as the social sentiment source. As of May 2026, StockTwits has placed their public API (`api.stocktwits.com/api/2/streams/symbol/{ticker}.json`) behind Cloudflare bot protection. All requests return HTTP 403 via a Cloudflare challenge page. **Do not attempt to revive this without a paid StockTwits API key.**

### 6.2 yfinance News Sentiment (Active)

Uses `yf.Ticker(ticker).news` — no separate API key, no quota consumption.

**Schema change (important):** As of yfinance ~0.2.x, the news response changed format:
- **Old:** `item['title']`, `item['providerPublishTime']` (Unix timestamp)
- **New:** `item['content']['title']`, `item['content']['pubDate']` (ISO string), `item['content']['summary']`

Our `_fetch_yf_news_sentiment()` handles both schemas transparently.

Sentiment is scored via keyword matching on `title + summary`. Keyword sets in `_BULL_WORDS` and `_BEAR_WORDS` frozensets. A headline with both bull and bear words scores 0 (neutral). Bullish threshold for sub-signal: ≥65% bullish ratio.

### 6.3 Reddit PRAW (Pending Credentials)

Searches r/wallstreetbets, r/stocks, r/investing, r/StockMarket for ticker mentions in the last 7 days. Scores via:
1. Keyword sentiment on post title + body text
2. `upvote_ratio`: ≥0.75 = bullish crowd signal, <0.35 = bearish

Requires env vars on Render:
- `REDDIT_CLIENT_ID` — from reddit.com/prefs/apps (script app)
- `REDDIT_CLIENT_SECRET` — from same page
- `REDDIT_USER_AGENT` — e.g. `InvestmentResearch/1.0 by YourRedditUsername`

**Graceful degradation:** if credentials are missing, returns `None` for all fields without crashing. The scoring treats missing data as neutral (doesn't penalise). So Reddit being unconfigured costs zero points.

**Setup issue encountered:** Reddit's reCAPTCHA on the app creation page was blocking on Chrome and Edge. Workaround: try Firefox private window or mobile browser on mobile data.

### 6.4 Social Trend Scoring Logic (4 sub-signals)

```
Sub-signal 1: Ticker Google Trends ratio    ≥2.0× = bullish, ≥1.3× = neutral/elevated, <1.3× = bearish
Sub-signal 2: Best narrative keyword ratio  ≥2.0× = bullish, ≥1.3× = bullish (early entry), <1.3× = bearish
Sub-signal 3: yfinance news bullish %       ≥65% = bullish, 40-65% = neutral, <40% = bearish
Sub-signal 4: Reddit bullish %              ≥65% = bullish, 40-65% = neutral, <40% = bearish

HIGH = ≥3 bullish votes
MOD  = 1-2 bullish votes (or all unavailable)
LOW  = 0 bullish votes with ≥1 bearish vote
```

---

## 7. Narrative Theme Scoring

### 7.1 Hot Themes Get Automatic Upgrade

If a theme is in the "hot themes" set AND user selects MOD or HIGH strength, the tier is forced to HIGH (10 pts). Hot themes as of build date:

```
AI / Machine Learning, Defence & Aerospace, GLP-1 / Weight Loss,
Nuclear / SMR, Crypto-Adjacent, Biotech Catalyst
```

Rationale: the market is actively re-rating these sectors, so the narrative tailwind is structural, not just story.

### 7.2 No Theme = Automatic LOW

A speculative play without a narrative story rarely outperforms. The scoring penalises "None" with 0 pts and a note explaining why.

---

## 8. Catalyst Scoring

Timing tiers:
- **Near** (<3 months, specific event) = HIGH (10 pts)
- **Medium** (3–9 months, expected) = MOD (5 pts)
- **Vague** (no clear date) = LOW (0 pts)

A catalyst description under 5 characters = LOW regardless of timing. This prevents gaming the input with a one-word entry.

---

## 9. Data Sources Summary

| Signal | Source | API Key Required? |
|---|---|---|
| Price, MA, RSI, MACD, EMA | FMP `/technical_indicator/daily` | Yes (FMP) |
| Volume history | FMP `/historical-price-full` | Yes (FMP) |
| Short interest | FMP `/short-interest` | Yes (FMP) |
| Analyst revisions | FMP `/upgrades-downgrades` + `/analyst-estimates` | Yes (FMP) |
| Insider trading | FMP `/insider-trading` | Yes (FMP) |
| Balance sheet | FMP `/balance-sheet-statement` | Yes (FMP) |
| Income statement | FMP `/income-statement` | Yes (FMP) |
| Options P/C ratio | yfinance (free) | No |
| Google Trends | pytrends (free, rate-limited) | No |
| yfinance news sentiment | yfinance (free) | No |
| Reddit sentiment | PRAW (free, needs app credentials) | Reddit app only |

---

## 10. FMP API Endpoint Notes

All FMP calls use the `/stable` base URL (not the older versioned endpoints). The `symbol=` param format is used throughout (e.g. `profile?symbol=AAPL`).

Short interest: `short-interest?symbol={ticker}` — returns `shortPercentOfFloat` or `shortPercent` (both checked). Value comes back as a decimal (0.18 = 18%) so multiply by 100 before storing.

Insider trading: only `P-Purchase` transaction type counts as a buy signal. `S-Sale` is tracked separately. Option exercises and gifts are ignored by design.

MACD: FMP returns `macd`, `signal`, `histogram` keys. Bullish when `histogram > 0 AND macd > signal`.

EMA: separate calls for period=21 and period=50. Returns `ema` key in first list element.

---

## 11. Excel Workbook

Three tabs:
1. **Speculative Scorecard** — all 12 dimensions with tier/pts/rationale, dark orange theme
2. **Scenario Model** — EV/Rev re-rating bear/base/bull, 1.5× analysis
3. **Market Data** — raw signal values (all numeric fields) for user reference

Colour palette: `--orange: #FF6B35`, `--amber: #FFA62B`, dark backgrounds matching HTML report aesthetic.

---

## 12. HTML Report Template

File: `Speculative_Report_Template.html`

Dark orange/amber theme. Sections:
- Topbar + hero card (score/120, verdict, bull return, bull price)
- Narrative/catalyst block
- 12-row scorecard table (tier chip + rationale)
- Scenario model table (bear/base/bull + 1.5× analysis)
- Market data chips (key raw signals)
- Disclaimer

Score bar scales against 120: `pct = min(100, max(0, round(score / 120 * 100)))`

Verdict scale note in template: `≥90 Moonshot · ≥72 Strong Speculative · ≥54 Speculative Play · ≥36 High Risk · <36 Pass`

---

## 13. Server Routes

```
POST /generate-speculative   — runs full MUFF pipeline, returns JSON with score/verdict/scenario
GET  /download/speculative-model/<ticker>  — returns Excel workbook
```

The `/generate-speculative` body params:
```json
{
  "ticker": "RKLB",
  "hold_months": 6,
  "narrative_theme": "Space / Deep Tech",
  "narrative_strength": "HIGH",
  "catalyst_desc": "Neutron first launch scheduled Q4 2026",
  "catalyst_timing": "medium",
  "custom_terms": "reusable rocket, launch vehicle, Neutron rocket",
  "mock": false,
  "password": "..."
}
```

`custom_terms` is a comma-separated string, parsed server-side to a list before passing to `fetch_speculative_data()`.

---

## 14. Pending / Known Issues

| Item | Status | Notes |
|---|---|---|
| Reddit PRAW credentials | Pending | User couldn't get through Reddit reCAPTCHA. Try Firefox private window or mobile data. Once configured, add `REDDIT_CLIENT_ID`, `REDDIT_CLIENT_SECRET`, `REDDIT_USER_AGENT` to Render env vars. |
| Deploy to Render | Pending | All code committed locally, needs `git push` to trigger Render redeploy. |
| pytrends 429 rate limiting | Mitigated | 2-second sleep before each query. May still fail under heavy load — consider caching Trends results for 1 hour if this becomes a problem. |
| Google Trends weekend noise | Known | Trends values drop on weekends (fewer searches). 7-day window smooths this but doesn't eliminate it. Not currently addressed in scoring. |

---

## 15. File Map (MUFF-specific)

| File | Role |
|---|---|
| `speculative_engine.py` | All 12 scoring functions, fetch helpers, scenario model, Excel builder |
| `speculative_report_bridge.py` | Maps engine dict → HTML template variables |
| `Speculative_Report_Template.html` | Dark orange HTML report template |
| `server.py` | Added `/generate-speculative` and `/download/speculative-model/<ticker>` routes |
| `static/index.html` | Added `#speculative` section with full form + `runSpeculative()` JS |
| `requirements.txt` | Added `pytrends`, `praw` |
