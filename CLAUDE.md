# Investment Automation — Project Rules

## End User
This tool is used by a **professional investor**. All calculations, formulas, and reported numbers must be accurate to institutional standard. No approximations presented as facts, no silent sign errors, no formulas that produce plausible-looking but incorrect results. When in doubt, show the work.

---

## Rule 1: HTML Report Must Exactly Reflect DCF Excel Model Outputs

The HTML report is the primary deliverable. Every number it presents that originates from the DCF model **must be identical to what the Excel workbook produces**. This is non-negotiable.

### Gordon Growth (GG) Bear / Base / Bull
- **Base case** = exact `gg_price` from the Python DCF engine (mirrors the Excel tab).
- **Bear** = TGR 2.0%, WACC unchanged (from Excel WACC tab output).
- **Bull** = TGR 4.0%, WACC unchanged.
- WACC is **never varied** across GG scenarios — only TGR changes.
- Pre-computed in `fmp_3statementv6.py` `build_dcf()` as `dcf_prices["gg_bear_price"]` / `dcf_prices["gg_bull_price"]`.
- `report_bridge.py` reads these keys directly. No approximation formulas.

### Exit Multiple (EM) Bear / Base / Bull
- **Base case** = exact `em_price` from the Python DCF engine (Excel model default: 20x EV/EBITDA).
- **Bear** = 75% of base multiple (15x when base is 20x).
- **Bull** = 125% of base multiple (25x when base is 20x).
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
  "gg_price":      float,   # Gordon Growth base (TGR 3%)
  "gg_bear_price": float,   # Gordon Growth bear (TGR 2%)
  "gg_bull_price": float,   # Gordon Growth bull (TGR 4%)
  "em_price":      float,   # Exit Multiple base (20x)
  "em_bear_price": float,   # Exit Multiple bear (15x)
  "em_bull_price": float,   # Exit Multiple bull (25x)
  "em_base_mult":  float,   # 20.0
  "em_bear_mult":  float,   # 15.0
  "em_bull_mult":  float,   # 25.0
  "gg_upside":     float,   # (gg_price / current_price) - 1
  "em_upside":     float,   # (em_price / current_price) - 1
}
```
