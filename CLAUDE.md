# Investment Automation — Project Rules

## End User
This tool is used by a **professional investor**. All calculations, formulas, and reported numbers must be accurate to institutional standard. No approximations presented as facts, no silent sign errors, no formulas that produce plausible-looking but incorrect results. When in doubt, show the work.

---

## Rule 1: HTML Report Must Exactly Reflect Model Outputs — No Exceptions

The HTML report is the primary deliverable. **Every scored, calculated, or tiered value it displays must be identical to what the Excel workbook produces.** This is non-negotiable and applies to all sections — not just the DCF valuation.

### Scorecard Tiers
All auto-scored criteria (Moat Profile, Management, Capital Returns, Execution Risk, Revenue CAGR, FCF Quality, ROIC, Leverage, Interest Cover, P/E, P/FCF) are computed once in `build_scorecard()` and stored in the `metrics` dict. `report_bridge.py` **must read these values directly** — never hardcode a fallback tier like `"MOD"` as a permanent default. If the engine value is missing (legacy cached report), `"MOD"` is acceptable as a last-resort fallback only.

The `metrics` dict keys for tiers passed to report_bridge:
```
tier_moat, tier_mgmt, tier_cap_ret, tier_exec
```
Section totals (`p1`, `p2`, `p3`) in `report_bridge.py` must use these live values so the HTML weighted scores match the Excel scorecard totals.

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
- **Bear / Bull TGR** = tier-specific values above. WACC is **never varied** — only TGR changes.
- Pre-computed in `fmp_3statementv6.py` `build_dcf()` as `dcf_prices["gg_bear_price"]` / `dcf_prices["gg_bull_price"]`.
- `report_bridge.py` reads these keys directly. No approximation formulas.

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
  "gg_upside":     float,   # (gg_price / current_price) - 1
  "em_upside":     float,   # (em_price / current_price) - 1
}
```
