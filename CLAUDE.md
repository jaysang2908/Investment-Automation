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

---

## Rule 10: Dashboard Score Must Always Match the Report Hero — No Exceptions

The `Auto_Score` column in `outputs.csv` is the single authoritative score that feeds the dashboard. It **must always equal the score displayed in the HTML report hero card** for that ticker. These are the same number and must never diverge.

### What score to write

| Situation | Hero score shown | Auto_Score to store |
|---|---|---|
| No qualitative inputs (BC/LTP) | `auto_score` (quant-only, 0-10) | `auto_score` |
| BC and/or LTP entered | `adj_score` (quant + qual pts, 0-10, capped by `floor_cap`) | `adj_score` |

The correct pattern in every write path:

```python
_display = adj_score if (biz_clarity or ltp) else auto_score
# store _display in Auto_Score column
```

### Write paths — all three must follow this rule

1. **`local_rerun.py` `_write_row()`** — pass `display_score=_display`; `_write_row()` stores it in `Auto_Score`.
2. **`server.py` `_update_outputs_csv()`** — same pattern via `display_score` parameter; called from `/generate` with `display_score=_display_score`.
3. **`server.py` `/api/qualitative/<ticker>`** — when qualitative inputs are updated via API, this endpoint regenerates the HTML report AND must also call `_update_outputs_csv(..., display_score=adj_score)` to keep the CSV in sync. It does this **in addition** to writing `qualitative_overrides.json`.

### After changing report_bridge.py or Report_Template.html

Run `python _rerender_reports.py` to re-render all 32+ HTML reports from cached data (no FMP calls). Then run `python _score_audit.py` to verify 0 mismatches between dashboard scores and report heroes before committing.

### Never re-derive scores in the dashboard

The dashboard reads `Auto_Score` directly from `outputs.csv` via `/api/scores`. It must not recalculate, reweight, or reinterpret scores. If the score on the dashboard differs from the report, fix the write path — do not patch the dashboard JS.

---

## Rule 11: WACC Floor at 8.5% (D-001)

`build_wacc()` floors `wacc_val` at 8.5% before returning. No equity investment should be discounted below the lowest reasonable equity return. 8.5% ≈ Damodaran composite Ke for a market-beta US name (Rf 4.3% + 1.0 × ERP 4.5% with a rounding buffer). FMP's 5-year regression betas systematically understate equity risk for stable compounders (PEP β = 0.41 → 6% Ke; AAPL β = 1.07 was hitting 2.1% WACC due to an upstream `rd_acctg ≈ 0` zero-out for cash-rich majors).

- Both Python (`fmp_3statementv6.py::build_wacc` ~L2077) and the Excel WACC tab (`wacc_formula` at ~L2016) wrap the formula in `MAX(0.085, …)` — Rule 1 requires both stay aligned.
- `wacc_refs` returns `wacc_val` (floored), `wacc_raw` (unfloored), `wacc_floored` (bool). These flow into `dcf_prices` and the report's WACC note: when floored, the note appends "Floored to 8.5% (raw was X.XX%) — FMP regression beta likely understates equity risk; override interactively at /dcf if needed".
- The DCF calculator at `/dcf` allows interactive WACC override per user request. This is the escape hatch for analysts who want to model a specific name with a different cost of capital.

## Rule 12: Output Guards — No Nonsense Numbers (F-I, F-B, F-H)

The hero card price target must never be a fabricated, negative, or stale value.

- **F-I — Negative price block:** In `build_dcf()` after computing `_ip_gg_usd` and `_ip_em_usd`, clamp values ≤ 0 to `None`. A negative equity-value-per-share is mathematically possible (positive EBITDA × multiple minus enormous net debt for banks / captive-finance autos) but unpublishable.
- **F-B — N/A instead of current price fallback:** In `report_bridge.py`, when neither GG nor EM produces a valid number, set `price_target = None` and `_primary_method = "N/A — Insufficient inputs"`. Never fall back to `current_price` (that masks a missing valuation as if it were a model output).
- **F-H — Honest method label:** "Composite avg" appears only when both methods are valid. Otherwise: `"Gordon Growth (primary — EM unavailable)"` / `"EV/EBITDA Exit (primary — GG unavailable)"` / `"N/A — Insufficient inputs"`.

When `price_target` is None, the rationale text below the hero explains the specific cause (foreign reporter, both methods unavailable, etc.) rather than leaving stale tier-branch wording.

## Rule 13: Engine-Entry Guards (F-P, F-S, F-N)

Inputs are validated before the engine commits to a valuation.

- **F-P — Foreign-reporter guard:** In `build_dcf()`, read `reportedCurrency` from `is_data[-1]` (the **latest** year — previous code read `[0]`, the oldest, which is a known sign-error). If reportedCurrency ≠ USD and the FMP FX endpoint fails (returns `_fx_to_usd = 1.0` fallback), set `_foreign_reporter_unsupported = True`, void all GG/EM prices to `None`, and surface `dcf_prices["foreign_reporter"]` + `dcf_prices["reported_currency"]` for the bridge to explain. TSM (TWD) was producing $22,854 fair value (+5,550%) until this guard landed.
- **F-S — Stale-data assertion:** At top of `report_bridge.py::build_report_data()`, check that `dcf_prices` has the core keys `{growth_tier, tgr_base, em_base_mult, neg_earnings_regime}`. If missing, render a "Data refresh required — run `python local_rerun.py <ticker>`" banner. Prevents stale JSONs from older engine schemas (TSLA case: data fetched 2026-04-26 with only 2 `dcf_prices` keys) from rendering as if they were complete.
- **F-N — MktCap ≈ Price × Shares triangulation (defensive):** In `build_report_data()`, if `|marketCap − price × sharesOutstanding| / marketCap > 0.10`, surface a soft amber banner. NOT a publish-block — legitimate stock splits can momentarily produce mismatched FMP snapshots (NFLX was a false alarm — it had done a real split and the data was internally consistent).

These three guards run at the engine boundary so a single bad input never produces a fully self-consistent but completely wrong report.

## Rule 14: Bank-Charter DCF Force-Disable (F-D Phase 1, F-M, F-Q)

Gordon Growth (FCF perpetuity) and EV/EBITDA Exit Multiple do not apply to deposit-funded balance sheet institutions. When `_is_bank_dcf = True`, the engine immediately voids all GG and EM outputs and bypasses EVS regime detection.

**Detection** (in `build_dcf()`)
```python
_BANK_DCF_EXCLUDE = {"V", "MA", "PYPL", "FIS", "FISV", "GPN", "WU", "DFS", "TRMK"}
_BANK_DCF_KW = {"bank", "banking", "financial services", "savings",
                "thrift", "mortgage", "credit union", "investment bank",
                "diversified financial"}
_is_bank_dcf = (
    any(kw in _prof_industry_dcf.lower() for kw in _BANK_DCF_KW)
    and ticker.upper() not in _BANK_DCF_EXCLUDE
)
```

Payment networks (V, MA, PYPL, etc.) are explicitly excluded from the bank classifier — they share FMP's `Financial - Credit Services` tag with deposit-taking banks but have completely different economics and DCF applicability. Add new payment-network tickers to `_BANK_DCF_EXCLUDE` as they are reviewed.

**When `_is_bank_dcf = True`:**
- `_gg_final`, `_em_final`, all scenario prices → `None`.
- `_neg_earnings_regime = False` and `_evs_regime = False` (F-M: bank FCF is accounting noise from loan/deposit flows — it must never trigger the negative-earnings regime or EVS overlay).
- `dcf_prices["bank_disabled"] = True` and `dcf_prices["bank_disabled_reason"]` carries a plain-English explanation for the report.

**In `report_bridge.py`:**
- Primary method label: `"N/A — Bank methodology pending (DDM / Justified P/B)"`.
- Rationale block explains that DDM (Dividend Discount Model) and Justified P/B (price-to-tangible-book vs ROE − g) are the correct methodologies and are pending Phase 2 implementation.
- Scorecard quality metrics remain valid and are displayed normally.

**Phase 2** (DDM / Justified P/B pricing) is deferred. Until implemented, bank reports show "N/A" for price target with a clear explanation.

## Rule 15: Scorecard Rescale Suppression for Valuation-Data Gaps (F-G)

The scorecard applies a `× (active_weight / scored_weight)` rescale when some criteria cannot be scored (e.g. EVS regime disables P/E and P/FCF). This correctly inflates the score back to an apples-to-apples 0–10 basis when criteria are *methodologically excluded* (banks, EVS tickers, cyclicals).

However, when **both P/E AND P/FCF tiers are None due to a data-fetch failure** (FMP ratios endpoint returned no data) — not a regime exclusion — the rescale must be suppressed. Inflating a score when the data simply didn't arrive is misleading.

**Detection** (in `build_scorecard()`):
```python
_fg_valuation_gap = (
    tier_pe is None and tier_pfcf is None
    and not is_bank and not evs_regime
)
_low_data_confidence = (
    _fg_valuation_gap
    or ((_scored_weight < 0.5 * _active_weight) if _active_weight else True)
)
if (_scored_weight > 0 and _scored_weight < _active_weight
        and not _low_data_confidence):
    _raw_sum = _raw_sum * (_active_weight / _scored_weight)
```

When `_fg_valuation_gap` fires: confidence level is set to `"LOW"` with note "P/E and P/FCF data unavailable — FMP ratios fetch failed; valuation criteria excluded, rescale suppressed".

**Always initialise `_fg_valuation_gap = False`** before the `if _scored:` block so the else-branch (no criteria scored at all) doesn't hit a NameError.

## Rule 16: Valuation Concordance Gate + Verdict Text Override (F-F, F-K)

Quality score and verdict are separate concerns. A high-quality business at an expensive price is not a "Buy" — but it shouldn't have its quality score penalised either.

**Concordance detection** (in `report_bridge.py`, after EVS/bank checks):
```python
_CONCORDANCE_THR = 0.25
_valuation_concordance = None
if (not _evs_regime and not _bank_disabled
        and _gg_up_con is not None and _em_up_con is not None):
    if _gg_up_con < -_CONCORDANCE_THR and _em_up_con < -_CONCORDANCE_THR:
        _valuation_concordance = "expensive"
    elif _gg_up_con > _CONCORDANCE_THR and _em_up_con > _CONCORDANCE_THR:
        _valuation_concordance = "cheap"
```

**Verdict text override** (F-K, immediately after `_conservative_verdict()`):
- If `_valuation_concordance == "expensive"` and current verdict rank > 2 → force to `"Hold — Premium Quality, Expensive"`.
- If `_valuation_concordance == "cheap"` and current verdict rank < 3 → force to `"Good Business at Fair Price"`.
- Quality score (`Auto_Score`) is **never modified** by this gate.

**Concordance banner** (F-F): a red/green HTML callout is injected below the hero card showing exact GG% and EM% upsides when concordance fires. This is the primary signal the reader sees — the verdict text change is the fallback guard, not the explanation.

**Concordance does NOT fire** for EVS-regime tickers (only one price method active) or bank-disabled tickers (no price methods at all).

## Rule 17: Cyclical Revenue Tier Smoothing (F-E)

For companies classified as `cyclical` by `_sector_bucket()`, using a simple 3-year average revenue growth for tier classification produces misleading results — commodity-cycle lows inflate or deflate the average based on which year you happen to land in.

**Smoothing rule** (in `build_dcf()`, after `_rev_3yr_avg_dcf` is computed):
- Collect all available YoY revenue growth rates from `hist_rev` (up to 5 years).
- If ≥ 3 valid periods exist, replace `_rev_3yr_avg_dcf` with the **median** of those rates.
- Median is more robust to a single catastrophic or boom year than the mean.
- Log: `"F-E: cyclical tier smoothing — using N-period YoY median"`.
- Falls back silently to 3yr average if fewer than 3 valid periods.

This only affects the growth tier (and therefore TGR/EM multiples) — it does not change any scorecard criteria.

## Rule 18: Widened EVS Regime Trigger (F-L)

The original EVS regime trigger (`neg_earnings_regime AND trailing_EBITDA < 0`) misses companies with technically positive EBITDA that is so thin it makes the Exit Multiple meaningless.

**Extended conditions** (added in `build_dcf()`, both exclude cyclicals to avoid false positives during commodity troughs):
```python
_ebitda_near_zero = (
    _trailing_ebitda_mm < 0.05 * max(_trailing_rev_mm_fl, 1)
    and not _is_cyclical_dcf
)
_fcf_deeply_neg = (
    _trailing_fcf_raw < -0.10 * max(_trailing_rev_mm_fl / 1e3, 1)
    and not _is_cyclical_dcf
)
_evs_regime = _neg_earnings_regime and (
    _trailing_ebitda_mm < 0
    or _ebitda_near_zero
    or _fcf_deeply_neg
)
```

- **EBITDA near-zero**: EBITDA < 5% of revenue (thin-margin turnarounds where the EM terminal value is a tiny number that moves enormously on small margin shifts).
- **FCF deeply negative**: FCF < −10% of revenue (cash-burn rate signals that a UFCF perpetuity is still meaningless even if accounting EBITDA is modestly positive).
- Cyclicals are excluded from both extensions because a cyclically depressed EBITDA/FCF is expected to recover — it's not the same structural situation as a secular pre-profit company.

## Rule 19: EV/Sales Subtype Classification (F-O)

`_secular_growth_subtype(ticker)` controls the mature EV/Sales multiple used when a company enters EVS regime. The mapping must reflect the actual business model economics, not generic tech labelling.

**Social media and ad-tech platforms** (SNAP, PINS, RDDT, BMBL, MTCH) are classified as `tech_growth` (4.5× mature EV/Sales) — NOT `secular_growth_deeptech` (4×). These platforms have consumer network effects, advertising-driven monetisation, and addressable markets calibrated to established consumer internet comps. `secular_growth_deeptech` is reserved for hardware/deep-science plays (space, quantum, robotics, novel biotech).

When adding a new pre-profit ticker: first determine the subtype before trusting a sector-label default. The multiplier difference between subtypes is 1–2× — a material impact on the derived price target.

## Rule 20: Valuation Field Audit (`_score_audit.py --full`)

Rule 10 audits `Auto_Score` (CSV) vs hero score (HTML). The extended audit (`python _score_audit.py --full`) additionally checks:
- `GG_Price` in `outputs.csv` vs `dcf_prices.gg_price` in `static/data/{ticker}_data.json`
- `EM_Price` in `outputs.csv` vs `dcf_prices.em_price` in the same JSON
- `price-value price-target` scraped from the HTML vs the above

Any discrepancy >$1 between CSV and JSON is flagged as a valuation field diff and listed at the end of the report. This catches state-management drift where a cached JSON was regenerated but the CSV wasn't synced (or vice versa).

Run `python _score_audit.py` (score check only) after every `_rerender_reports.py` pass. Run `python _score_audit.py --full` after any `local_rerun.py` pass.

## Rule 21: Historical EV/EBITDA Anchoring for Exit Multiple (F-C Phase 2)

Static tier defaults (10×/15×/18×) reflect median names in each growth bucket. Companies that consistently trade at a structural premium or discount have their own precedent — the historical anchor uses that signal.

**Logic** (in `build_dcf()`, runs after quality premium block):
```python
_HIST_EM_DISCOUNT = 0.80   # mean-reversion: market normalises somewhat over 5yr horizon
_HIST_EM_CAP      = 28.0   # prevents bubble-era averages from perpetuating
# fetch from FMP /stable/ratios?symbol={ticker}&limit=5
anchored_raw   = hist_5yr_ev_ebitda_avg × 0.80
anchored_final = max(tier_base, min(28.0, anchored_raw))   # floor + cap
# scale bear/bull proportionally; bull additionally capped at 32x
```

**Key design decisions:**
- **Floor = post-quality-premium base**: ensures the quality premium (+5×) is never erased by a depressed historical average (e.g. V, COST).
- **Cap = 28×**: NVDA (48.9×) and AMD (40.7×) five-year averages include 2020-2021 ZIRP peak. Using those raw would perpetuate bubble pricing in terminal value.
- **Cyclicals excluded**: F-E median-smoothing is a better anchor for commodity-cycle names.
- **Banks excluded**: EM is disabled for banks anyway.
- **EVS-regime tickers**: anchoring runs but is neutralised in `dcf_prices` output when `_evs_regime` fires.
- **Only applies if `abs(anchored_final − current_base) > 0.5×`** — avoids rounding noise.

**`dcf_prices` keys added:**
- `em_anchored` (bool): True when anchoring moved the multiple by >0.5×
- `em_hist_anchor_raw` (float): raw 5yr historical average before discount/cap
- `em_hist_anchor_capped` (bool): True when uncapped anchored value exceeded 28×

**Cap banner in `report_bridge.py`:**
When `em_hist_anchor_capped = True`, an amber callout renders below the hero:
> ⚠ Exit Multiple Capped at 28× (5yr historical avg: X.X×)
> The model applies a 28× ceiling... For extraordinary competitive positions, use the DCF Calculator to model it explicitly.

The banner is intentionally transparent: it shows the suppressed historical average so the user can decide whether the cap is appropriate or whether to override at `/dcf`.

**Portfolio impact (2026-05-25 patch):**

| Ticker | Old base | New base | Old EM% | New EM% | Capped? |
|--------|----------|----------|---------|---------|---------|
| KO | 10× | 15.4× | −48% | −19% | No |
| ABBV | 10× | 15.6× | −39% | −5% | No |
| AAPL | 10× | 18.4× | −56% | −20% | No |
| NKE | 10× | 19.0× | −13% | +67% | No |
| PEP | 10× | 13.5× | −32% | −8% | No |
| DIS | 10× | 16.1× | +5% | +69% | No |
| AMD | 18× | 28.0× | −40% | −6% | YES |
| NVDA | 18× | 28.0× | +28% | +99% | YES |
| ADBE | 15× | 23.3× | +77% | +175% | No |

Tickers unchanged by anchor (floor or within 0.5×): JNJ, WMT, HCA, FDX, TGT (floor), META, NFLX (floor), V (quality-premium floor), COST (quality-premium close), CSCO (minor +1.4×).

**When FMP API is available:** `local_rerun.py` automatically applies anchoring via the engine. Manual JSON patches above are temporary; re-run to confirm via engine when quota resets.

## Rule 22: No Speculative Valuation Claims (S-001)

**Before stating any current stock price, market cap, P/E, P/FCF, EPS, share count, or derived multiple in conversation, live data must be fetched from FMP.** Never use training-data prices, prior-session values, or "I recall/believe" phrasing for specific valuation numbers.

**Why this matters — the NFLX split case (2026-05-25):**  
NFLX executed a 10:1 forward split. The assistant recalled a pre-split price (~$880), halved it mentally to ~$308, and computed a P/E of 36× against split-adjusted EPS. The actual price was $88.60 and TTM P/E was 28× (Yahoo Finance) / 27.9× (FMP TTM calc). The error: mixing un-adjusted price memory with split-adjusted EPS from FMP — a silent 3× inflation of the multiple.

**Engine-level guard (S-001):**  
`build_dcf()` in `fmp_3statementv6.py` prints `[S-001 WARNING]` when `|shares_current − shares_prior_year| / shares_prior_year > 25%`. This catches cases where a split has occurred and FMP data may not yet be fully adjusted. Logged to stdout during report generation — does not block the DCF.

**How to verify quickly in conversation:**
```python
import requests, os
r = requests.get(
    f"https://financialmodelingprep.com/stable/profile"
    f"?symbol=TICKER&apikey={os.environ['FMP_API_KEY']}", timeout=8
).json()
print(r[0]["price"], r[0]["mktCap"], r[0]["sharesOutstanding"])
```

**Applies to:** every inline conversation claim about a specific ticker's current price or multiple. Does not apply to ranges or directional statements ("NFLX trades at a premium to the market" is fine; "NFLX P/E is 36×" requires live verification).

## Rule 23: Auto-Push Policy — Ship Every Improvement to Render by Default

**The user only ever views the live Render URL** (auto-deployed from `origin/main` of [jaysang2908/Investment-Automation](https://github.com/jaysang2908/Investment-Automation)). They never run the site locally. An edit that is coded but not pushed does not exist from their point of view — and historically this has wasted significant time going in circles ("nothing changed" → debugging cache → discovering the commit was never pushed).

**Default behaviour — push, don't ask:**
- Any change to a file under this repo that improves the model, engine, report, or site is committed **and pushed** to `origin/main` in the same turn it is made. This **overrides** the generic "commit/push only when asked" default — asking for the change *is* the authorization to ship it.
- This applies to work we initiate as well as work the user requests. If we identify and implement an improvement (a scoring fix, a bug fix, a refactor), we push it. We do **not** leave it staged or local "pending review".
- Do **not** ask "want me to push?" / "should I deploy?". Bundle the push with the change and state the commit SHA + "Render redeploys in ~2-4 min" in the summary.

**The only exceptions (do not push):**
1. The user explicitly says "don't push", "wait", "just stage it", "I'll push", or similar.
2. The change is genuinely incomplete or unverified (syntax error, failing smoke test, half-finished edit). Finish or revert it — never push known-broken code. If blocked (e.g. an FMP-quota'd re-run can't regenerate outputs), **say so explicitly in the summary** and record it in the project memory as pending, rather than silently leaving it unpushed.

**Hygiene when pushing:**
- Stage only the files relevant to the change. Never sweep in unrelated `M`/`??` files (other in-progress work, scratch scripts, unrelated data refreshes).
- If `git push` is rejected (remote ahead — a scheduled cloud run added tickers), rebase onto `origin/main`, resolve any `outputs.csv` conflict by keeping both the remote's new tickers and our updated rows, then push.
- After engine/scorecard changes, re-score affected tickers (`python _rescore_offline.py <tickers>` — zero FMP calls) and run `python _score_audit.py` (expect 0 mismatches) **before** pushing, so the deployed CSV/reports match the engine.

## Rule 24: Deferred Re-Runs Must Be Logged — No Silent Limbo (FMP-Blocked Work)

Some engine/scorecard changes only take visible effect after the affected tickers are re-run through FMP (the cached JSONs/reports/CSV predate the fix). When that re-run can't be completed in the same session — FMP daily quota exhausted, no API key in the run environment, etc. — the work is **deferred**. Historically these deferrals were noted vaguely ("re-run pending") and then forgotten, leaving fixes in limbo where the code shipped but the user never saw the effect.

**This is now structurally prevented. When a re-run is deferred, you MUST:**

1. **Add an entry to `pending_reruns.json`** (repo root) — `{id, tickers[], reason, fix_commit, deferred (date), blocker, action, status:"pending"}`. This is the single source of truth.
2. **Surface it in the turn summary** — state plainly which tickers are affected, what fix is waiting, and why it couldn't run. Never end a turn leaving deferred work implicit.
3. **Commit + push the ledger** with the change (Rule 23) so the dashboard banner updates.

**The ledger is self-surfacing and self-clearing:**
- `/api/pending-reruns` serves it; `static/dashboard.html` renders an amber banner listing the pending tickers whenever any exist — so the deferral is visible on the live URL the user actually looks at, not buried in a file.
- `server.py::_clear_pending_rerun()` runs after every successful live `/generate` and removes that ticker from the ledger automatically. So the moment the blocked re-run actually happens (user runs the ticker, or a cron does), the entry disappears on its own. No manual cleanup, no stale banner.

**Do NOT** use `pending_reruns.json` as a generic TODO list. It is specifically for "fix is committed, outputs are stale until an FMP re-run." Un-made fixes, feature ideas, and design debates belong in the project memory or a normal discussion — not here.

When you DO complete a deferred re-run yourself (e.g. quota reset, key available), remove the entry (or let the server clear it) and confirm the affected scores/prices in the summary.
