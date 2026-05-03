# Investment Automation — Project Rules

## End User & Mindset
This tool is built for **professional-level investors** — people who run DCF models themselves, read 10-Ks, and will immediately notice a wrong number, a sloppy label, or an insight that doesn't hold up. The site is also being reviewed by senior finance professionals (e.g. structured credit directors at major banks) who will evaluate it as evidence of analytical rigour.

That sets the bar for everything built here:

- **Accuracy first.** All calculations, formulas, and reported numbers must be correct to institutional standard. No approximations presented as facts, no silent sign errors, no formulas that produce plausible-looking but wrong results. When in doubt, show the work.
- **High attention to detail.** Labels, units, formatting, source attributions, and edge-case handling matter. A number displayed as "N/A" with a clear reason is better than a number that's silently zero or miscomputed.
- **Insightful, not decorative.** Every piece of information shown — a scorecard tier, a news headline, a valuation overlay — should help the user make a better investment judgment. Features that add noise or require explanation rather than delivering immediate signal should not be built.
- **Never dumb it down.** Users understand WACC, EBITDA multiples, FCF yield, credit spreads, and leverage ratios. Write labels, rationales, and UI text at the level of someone who reads Bloomberg and CFA-level material daily.
- **Incomplete features must be clearly flagged.** If a section of the site is not yet functional (e.g. Heatmap), it must visibly say so — do not leave a broken or empty state that a professional reviewer would interpret as a mistake.

## Ancillary Features Philosophy
The site's non-report features (News, Dashboard, Heatmap, Daily Discoveries) exist to make the user's workflow more convenient and self-contained — not to replace the core DCF/scorecard output. Their design principle is:

- **Stay current**: surface updated market information (news headlines, price moves) so the user can quickly cross-check assumptions after a report is generated without leaving the site.
- **Link to sources**: wherever possible, provide direct access to primary reference data (news articles, filings) so the user can verify the model's inputs and check our work.
- **Don't add noise**: only show information relevant to tickers the user has already run reports for. The system auto-discovers covered tickers from `static/reports/` and scopes all feeds to that universe.
- **Respect API limits**: ancillary features run on a scheduled basis (not on every page load or every report generation) to avoid burning FMP free-tier quota. News is fetched via `daily_news.py` on a cron — not live. The FMP `/stable/stock_news` batch endpoint fetches all tickers in a single API call to minimise quota usage.

When building or extending these features, default to scheduled/cached data over live API calls, and always tie the ticker universe to the user's generated report set.

---

## Rule 1: HTML Report Must Exactly Reflect Model Outputs — No Exceptions

The HTML report is the primary deliverable. **Every scored, calculated, or tiered value it displays must be identical to what the Excel workbook produces.** This is non-negotiable and applies to all sections — not just the DCF valuation.

### Scorecard Tiers
All auto-scored criteria (Moat Profile, Management, Capital Returns, Execution Risk, Revenue CAGR, FCF Quality, ROIC, Leverage, Interest Cover, P/E, P/FCF) are computed once in `build_scorecard()` and stored in the `metrics` dict. `report_bridge.py` **must read these values directly** — never hardcode a fallback tier like `"MOD"` as a permanent default. If the engine value is missing (legacy cached report), `"MOD"` is acceptable as a last-resort fallback only.

The `metrics` dict keys for **all** tiers passed to report_bridge:
```
tier_moat, tier_mgmt, tier_cap_ret, tier_exec,
tier_rev_cagr, tier_fcf_ni, tier_roic, tier_leverage, tier_ebit_int, tier_pe, tier_pfcf
```
Section totals (`p1`, `p2`, `p3`) in `report_bridge.py` must use these live values so the HTML weighted scores match the Excel scorecard totals.

**report_bridge.py must NEVER re-derive a tier independently.** The scorecard engine applies sector-specific thresholds (`SECTOR_THRESHOLDS`), trend penalties (e.g. FCF/NI declined >15pp → down-tier), and a 4-tier scale (HIGH/MOD-HIGH/MOD-LOW/LOW). Any separate re-derivation in report_bridge will diverge. The `_tier_rev_cagr()`, `_tier_fcf_ni()`, `_tier_roic()`, `_tier_ebit_int()`, `_tier_d_ebitda()`, `_tier_pe()`, `_tier_pfcf()` functions exist only as legacy fallbacks for stale cached reports — do not use them for fresh renders.

`TIER_PTS` must include all 4 tiers: `{"HIGH": 10, "MOD-HIGH": 7, "MOD": 7, "MOD-LOW": 3, "LOW": 0}`

### Dual Scoring + Conservative Verdict
The scorecard is reported on two scales and **both must always be visible** in the HTML report when qualitative inputs are provided:

- **Quant Score (max 87.5)** — 11 auto-scored criteria from FMP data only (no user input). Always shown.
- **Full Score (max 100)** — Quant + Business Clarity (2.5 wt) + Long-Term Potential (10.0 wt). Shown when the user supplies BC and/or LTP via the Render web form.

**Verdict rule:** convert each score to % of its max, apply identical %-bands (`≥75% → High Conviction Buy`, `≥65% → Good Business at Fair Price`, `≥50% → Hold — Monitor`, else `Avoid`), and **take the more conservative (lower) verdict** between the two. Implemented in `_conservative_verdict()` in `report_bridge.py`.

**Qualitative does NOT flow into the DCF.** Business Clarity and Long-Term Potential are predictability/TAM judgments — they are not financial inputs and would distort cash-flow projections. They affect the scorecard verdict only.

**Excel pre-fill:** when the user supplies BC/LTP on the web form, `server.py` passes them into `build_scorecard()` so the Excel scorecard tier cells are pre-populated (matching the HTML). Dropdowns remain active so the user can override in Excel.

### Growth Tier Classification
Companies are auto-classified by 3-year average annual revenue growth (last 3 YoY periods from `is_data`):

| Tier | 3yr Avg Rev Growth | TGR Base | Bear TGR | Bull TGR | EM Base | EM Bear | EM Bull |
|---|---|---|---|---|---|---|---|
| Low | < 5% | 2.5% | 2.0% (×0.80) | 3.0% (×1.20) | 10x | 8x | 12x |
| Medium | 5%–12% | 3.0% | 2.25% (×0.75) | 3.75% (×1.25) | 15x | 11x | 19x |
| High | > 12% | 4.0% | 3.0% (×0.75) | 5.0% (×1.25) | 18x | 14x | 23x |

Tier is computed in `build_dcf()` and stored in `dcf_prices["growth_tier"]`. The Excel TGR cell and exit multiple cell use these values — not hardcoded constants.

### Primary Price Target Method
- **Low / Medium (<10% growth):** Gordon Growth is primary (stable cash-flow companies).
- **Medium (≥10% growth) / High:** Exit Multiple is primary (growth companies valued on EBITDA exit).
- Price target, method label, and 3-line rationale are computed in `report_bridge.py` and mapped to `PRICE_TARGET`, `PRICE_TARGET_METHOD`, `PRICE_TARGET_RATIONALE` template variables.

### Gordon Growth (GG) Bear / Base / Bull
- **Base case** = exact `gg_price` from the Python DCF engine (mirrors the Excel tab).
- **Bear** = TGR tier-bear AND WACC +0.5pp. **Bull** = TGR tier-bull AND WACC −0.5pp.
- WACC shift is ±0.5 percentage points (`_WACC_SHIFT = 0.005`) — stored as `dcf_prices["wacc_bear"]` / `dcf_prices["wacc_bull"]`.
- Pre-computed in `fmp_3statementv6.py` `build_dcf()` as `dcf_prices["gg_bear_price"]` / `dcf_prices["gg_bull_price"]`.
- `report_bridge.py` reads these keys directly. No approximation formulas.
- The HTML scenario table (`DCF_BEAR_WACC` / `DCF_BULL_WACC`) must display the scenario-specific WACC, not the base WACC.

### Exit Multiple (EM) Bear / Base / Bull
- **Base case** = exact `em_price` from the Python DCF engine.
- **Bear / Bull multiples** = tier-specific values above.
- Pre-computed in `fmp_3statementv6.py` as `dcf_prices["em_bear_price"]` / `dcf_prices["em_bull_price"]`.
- Report reads these directly.

### Sensitivity Grid (WACC × TGR matrix)
- The 6×5 grid in the HTML report is an approximation for visual reference only (spread-ratio formula).
- The primary scenario table rows (bear/base/bull) are always from the engine — never from the grid formula.

### Composite Fair Value
- Average of GG base and EM base — both from the engine.
- Label must state the exact WACC and exit multiple used.

---

## Rule 2: DCF Formula Correctness

All DCF formulas must conform to standard UFCF methodology:

```
UFCF = NOPAT + D&A − ΔWC − CapEx

NOPAT = EBIT × (1 − effective tax rate)
Effective tax rate = MAX(0, MIN(50%, incomeTaxExpense / incomeBeforeTax))
```

Key sign conventions in the Excel model:
- **Tax on EBIT row**: must always be ≤ 0 (it is a deduction). Formula: `= −EBIT × tax_rate`. Tax rate is clamped 0–50% to prevent sign flip from tax-benefit years.
- **D&A row**: stored as negative (cost). D&A add-back row flips sign back to positive.
- **CapEx row**: stored as negative (cash outflow).
- **NWC change row**: negative when NWC is growing (cash outflow). NWC% assumption = `+ΔNWC/Revenue` (positive when NWC consumes cash). Row formula: `= −Revenue × NWC%`.
- **ROIC denominator**: Equity + **Net Debt** (= STD + LTD − Cash). Never use LTD alone.

---

## Rule 3: Data Source Hierarchy

1. **FMP API** for income statement, balance sheet, cash flow (5 years).
2. **Analyst estimates** (FMP `/stable/analyst-estimates`) for years 1–3 revenue and EBITDA projections in the DCF.
3. **Gemini 1.5 Flash** for qualitative commentary only — never for financial figures.
4. No training-data assumptions for financial values. Always pull live from FMP.

---

## Rule 4: Consistent Python / Excel Computation

The Python DCF engine in `build_dcf()` (used for `dcf_prices`) must use **identical assumptions** to the Excel model:
- NWC%: `+ΔNWC/Revenue` (not the old negative form).
- Tax rate: `abs(tax) / abs(EBT)` clamped at 0–50%.
- Terminal year revenue grown by the scenario TGR (not always 3%).
- `_py_ufcf()`: `return nopat + da - rev * avg_capex_pct - rev * avg_nwc_pct`.

If the Excel formula logic changes, the Python mirror must be updated in the same commit.

---

## Rule 5: No Silent Failures on Valuation Numbers

- If a DCF price cannot be computed, show "N/A" — never show $0 or a stale cached value.
- If `wacc ≤ tgr`, the Gordon Growth formula is undefined — return `None`, display "N/A".
- Scenario prices that would imply negative equity value should return `None`.

## Rule 6: Negative-Earnings Regime — Disable Gordon Growth

Gordon Growth requires stable positive UFCF growing forever. When trailing FCF or trailing EBIT is **negative**, the perpetuity formula produces nonsense (negative terminal value → negative implied price per share). This is the canonical "DCF fails on this name" case (turnarounds, deeply cyclical bottoms, pre-profit growers).

Detection lives in `build_dcf()`:
```
_neg_earnings_regime = (trailing_FCF < 0) OR (trailing_EBIT < 0)
```

When triggered:
- `dcf_prices["gg_price"]`, `gg_bear_price`, `gg_bull_price`, and `gg_upside` are all set to `None`.
- `dcf_prices["neg_earnings_regime"] = True` and `dcf_prices["gg_disabled_reason"]` carries an explanation string.
- `report_bridge.py` overrides the tier-based primary method and forces **EV/EBITDA Exit Multiple as sole primary**, regardless of growth tier.
- The HTML scenario table shows `"N/A — GG disabled (negative FCF/EBIT)"` in the GG row — must NOT fall back to EM price.
- The price target rationale displays the trailing FCF and EBIT figures so the user understands why GG was bypassed.

## Rule 7: Narrative-Gap Banner

When `|price_target / current_price − 1| > 40%`, render a banner immediately below the hero card flagging the divergence. The model produces an honest fundamentals-only number and surfaces the gap — **never fudge inputs to match the market price**.

The banner content is **dynamic in two dimensions**, never company-specific:

1. **Direction** — premium (market > fundamentals) vs discount (market < fundamentals); each gets a different framing line and a different set of example drivers.
2. **Sector bucket** — `tech_growth` / `stable_compounder` / `cyclical` / `bank` (read from `scorecard_metrics["sector_bucket"]`). Each bucket has its own list of plausible premium and discount drivers. Falls back to generic language when sector is unknown.

The example drivers are intentionally generic ("rate-cycle benefit", "regulatory overhang", "takeout speculation") — never name specific companies, programs, or events (e.g. don't say "CHIPS Act"). The banner's job is to prompt user judgment, not diagnose the cause.

Template variable: `{{NARRATIVE_GAP_BANNER}}` — produces empty string when gap < 40%.

## Rule 8: Negative-Multiples Scoring

In `_t_val()` (Part 4 valuation scoring), if current P/E or P/FCF is ≤ 0, return tier `"LOW"` with a note: "Multiple meaningless when earnings/FCF are negative." A loss-making company does not get cheaper as losses widen; the math may compute a "−300% vs benchmark" reading but that signals distress, not value. Likewise if the historical 5yr average is ≤ 0 (loss-period distortion), return tier `None` with an N/A note rather than scoring against a meaningless baseline.

## Rule 9: EV/Sales Regime — Pre-Profit Secular-Growth Companies

Triggered when `neg_earnings_regime = True` **AND** `trailing_EBITDA < 0`. At this point both GG and EV/EBITDA Exit Multiple are unreliable (negative EBITDA makes the EM terminal value nonsense). EV/Sales with a mature-business multiple is used instead.

Detection in `build_dcf()`:
```
_evs_regime = _neg_earnings_regime AND (hist_ebitda[-1] < 0)
```

**`_secular_growth_subtype(ticker)`** classifies the company into:
- `secular_growth_deeptech` → 4.0x mature EV/Sales (space, quantum, robotics, biotech)
- `secular_growth_software` → 6.0x (SaaS/data platforms at scale)
- `secular_growth_resources` → 2.5x (clean energy, critical materials)
- `tech_growth` → 4.5x / `stable_compounder` → 3.5x / `cyclical` → 1.5x (fallbacks)

**Forward price target:**
```
Year-5 EV = Year-5 revenue (from DCF projections) × mature EV/Sales multiple
Year-5 equity value = Year-5 EV − net_debt − minority_interest
EVS price = Year-5 equity value / (1 + WACC)^5 / shares_outstanding (in USD)
```

**Reverse check** (what CAGR does current market price imply?):
```
current_EV = price × shares + net_debt + mi
required_rev_5yr = current_EV / mature_multiple
implied_CAGR = (required_rev_5yr / trailing_rev)^(1/5) − 1
```

`dcf_prices` keys: `evs_regime` (bool), `evs_price`, `evs_implied_cagr`, `evs_required_rev` ($B), `evs_mature_mult`, `evs_subtype`, `evs_yr5_rev_b` ($B), `evs_upside`.

In `report_bridge.py`:
- `_evs_regime` takes precedence over `_neg_earnings_regime` for primary method selection.
- Primary method label: `EV/Sales (Nx mature multiple)`.
- Rationale includes trailing FCF/EBIT/EBITDA, Year-5 revenue, WACC used, and reverse-check CAGR.
- Narrative-gap banner appends a reverse-check line when EV/Sales is active.
- Composite fair value uses `evs_price` alone (no composite with GG/EM).
- EV/EBITDA valuation verdict rows show "N/A — trailing EBITDA negative" rather than fabricated prices.

**EV/Sales price target in the Excel model:** Not currently written to the DCF sheet (EV/Sales is a Python-only overlay — it doesn't map to Excel rows that assume positive EBITDA).

---

## Architecture Reference

| File | Role |
|---|---|
| `fmp_3statementv6.py` | Excel workbook builder + Python DCF engine |
| `report_bridge.py` | Maps engine outputs → HTML template variables |
| `Report_Template.html` | HTML report template with `{{VARIABLE}}` placeholders |
| `server.py` | Flask backend — calls engine + bridge, persists outputs |
| `app.py` | Streamlit wrapper (legacy, wraps same engine) |
| `data_store.py` | Caches ticker data to avoid repeat FMP calls |
| `scenarios_db.py` | SQLite store for saved DCF scenarios |
| `outputs.csv` | Scorecard metrics per ticker — feeds heatmap dashboard |

### Key `dcf_prices` dict keys (returned by `build_dcf()`)
```python
{
  "gg_price":      float,   # Gordon Growth base (tier TGR)
  "gg_bear_price": float,   # Gordon Growth bear (tier TGR × bear factor)
  "gg_bull_price": float,   # Gordon Growth bull (tier TGR × bull factor)
  "em_price":      float,   # Exit Multiple base (tier base multiple)
  "em_bear_price": float,   # Exit Multiple bear (tier bear multiple)
  "em_bull_price": float,   # Exit Multiple bull (tier bull multiple)
  "em_base_mult":  float,   # e.g. 10.0 / 15.0 / 18.0 by tier
  "em_bear_mult":  float,   # e.g. 8.0 / 11.0 / 14.0 by tier
  "em_bull_mult":  float,   # e.g. 12.0 / 19.0 / 23.0 by tier
  "tgr_base":      float,   # e.g. 0.025 / 0.030 / 0.040 by tier
  "tgr_bear":      float,   # bear TGR for GG scenario
  "tgr_bull":      float,   # bull TGR for GG scenario
  "growth_tier":   str,     # "low" | "medium" | "high"
  "rev_3yr_avg":   float,   # 3yr avg annual revenue growth used for tier
  "gg_upside":         float,   # (gg_price / current_price) - 1
  "em_upside":         float,   # (em_price / current_price) - 1
  "trailing_ebitda_b": float,   # trailing EBITDA in $B
  "neg_earnings_regime": bool,  # trailing FCF < 0 OR trailing EBIT < 0
  "evs_regime":        bool,    # neg_earnings_regime AND trailing EBITDA < 0
  "evs_price":         float,   # EV/Sales fwd price target (USD)
  "evs_implied_cagr":  float,   # 5yr revenue CAGR implied by current market price
  "evs_required_rev":  float,   # required trailing revenue in $B at mature multiple
  "evs_mature_mult":   float,   # sector-calibrated mature EV/Sales multiple
  "evs_subtype":       str,     # secular_growth_deeptech | _software | _resources | ...
  "evs_yr5_rev_b":     float,   # Year-5 projected revenue in $B
  "evs_upside":        float,   # (evs_price / current_price) - 1
}
```
