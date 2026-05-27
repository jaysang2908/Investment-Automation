# Investment-Automation — Cumulative Fixes Log

**Purpose:** Single source of truth for proposed, approved, and implemented fixes from the rolling scorecard review. Every batch of ticker assessments adds to this file. Before running another batch, review this log so incremental fixes don't compound bad flow-on effects.

**How to read:** Each fix has a stable ID (A1, A2, …). Once approved, fixes get a target order and any cross-coupling notes. Once shipped, status moves to **IMPLEMENTED** with the commit SHA.

---

## Review heuristics (lessons from review process)

- **Clean integer multiples (10×, 5×, 7×) = corporate action, not data bug.** A 10× price/sharecount discrepancy is the signature of a stock split — verify against EPS reconciliation and historical chart adjustment before flagging as a data integrity bug. NFLX (Batch 3, 2026-05-24) was a false alarm — the engine handled the split correctly because FMP returned a consistent post-split snapshot.
- **A clean −90% / +50% spread between GG and EM is methodological**, not arithmetic — usually means the underlying business doesn't suit one of the two frameworks (banks → neither suits; hypergrowers → GG fails; mature compounders → EM understates).
- **Score-vs-method divergence is the highest-value signal.** When quality score says "Buy" and both DCF methods say −30%+, the verdict is wrong, not the score. F-K (verdict text guard) is the most important new fix.

---

## Decisions locked in

### D-002 — F-F design revision (2026-05-24, after user pushback)

- **Original proposal:** cap Auto_Score at 4.5 when both GG and EM signal <−25% upside. User pushed back: COST at 4.5 is too punitive — 37% ROIC, zero leverage, dominant moat is a 7.9-quality business at a rich price, not a 4.5 business.
- **Revised design:** **Quality score stays intact**; the concordance signal acts on the **verdict text** and a **dedicated banner**, not the score.
  - F-F now: when both methods <−25%, set `metrics["valuation_concordance"] = "expensive"`; when both >+25%, set to `"cheap"`. Score unchanged.
  - F-K now: verdict ladder enforced — when `valuation_concordance = "expensive"`, max verdict allowed is `"Hold — Premium Quality, Expensive"`. Cannot show "High Conviction Buy" or "Good Business at Fair Price" even if score is 8+. When `"cheap"`, minimum verdict allowed is `"Good Business at Fair Price"` (cannot show "Avoid" even if score is 5-ish from poor growth tier).
  - New banner: prominent red/green callout under hero explaining "Both Gordon Growth and EV/EBITDA Exit signal material [over/under]valuation. Quality is [strong/weak]; entry point is [not / favourable]."
- **Net effect on COST:** Score stays 7.9. Verdict goes from "High Conviction Buy" → "Hold — Premium Quality, Expensive". Banner says "both DCF methods −48 to −70%". Reader sees full picture.
- **Status:** APPROVED design; supersedes the F-F entry in the original log.

### D-001 — Global WACC floor 8.5% (2026-05-24)
- **Decision:** Hard floor on `wacc_val` at 8.5% inside `build_wacc()` regardless of beta, sector, or capital structure.
- **Rationale:** No equity investment should be discounted below the lowest reasonable equity return. 8.5% ≈ Damodaran composite Ke for a market-beta US name (Rf 4.3% + 1.0 × ERP 4.5% with rounding buffer). FMP 5-yr regression betas systematically understate equity risk for stable compounders (PEP β=0.41 → Ke ≈ 6%; AAPL ended up at 2.1% due to a separate Kd bug). A blanket floor is more robust than per-sector calibration.
- **Escape hatch:** DCF calculator at `/dcf` lets the user manually override WACC interactively — preserves analyst flexibility while keeping the automated batch run conservative.
- **Implementation:** `fmp_3statementv6.py::build_wacc()` L2077–2095 wraps `wacc_val = max(wacc_raw, 0.085)`. Excel WACC sheet output cell wrapped with `=MAX(0.085, …)` at L2016. `report_bridge.py` WACC_NOTE appends "Floored to 8.5% (raw was X.XX%)" when triggered. New `dcf_prices` keys `wacc_raw` + `wacc_floored` for downstream consumers.
- **Supersedes:** Fix A (sanity floor at 6–7%) and Fix A's root-cause `rd_acctg` repair. The floor catches the AAPL symptom regardless of the upstream Kd bug.
- **CLAUDE.md:** Rule 11 added 2026-05-24.
- **Impact:** AAPL WACC 2.1% → 8.5% (GG goes from $39 to a plausible number; EM may become computable). Also lifts: KO (7.0%), JNJ (3.0%), V (7.7%), PEP (6.9%), CVX (7.0%), TGT (8.0%), WMT (8.1%), HCA (8.1%). Lifts mean **lower** GG/EM prices → some upside flips to downside.
- **Status:** ✅ **IMPLEMENTED 2026-05-24.** Effect on already-cached JSONs is null until `local_rerun.py` re-runs the engine. New engine calls (post-implementation) honour the floor automatically.

---

## Proposed fixes from Batch 1 (AAPL, ABBV, ADBE, AMD, BAC)

| ID | Fix | Verdict (agent review) | Order | Status | Coupling notes |
|---|---|---|---|---|---|
| F-H | Hero method label ("GG primary" / "EM primary" / "Composite") | SAFE | 1 | **✅ IMPLEMENTED 2026-05-24** (report_bridge.py L1009–1025) | Display-only. No CSV/Excel coupling. |
| F-B | Stop falling back to current price; show "N/A — Insufficient inputs" | SAFE | 2 | **✅ IMPLEMENTED 2026-05-24** (report_bridge.py L1052, L971, L1826, L1140 rationale override) | Rule 5 requires it. Pairs with F-H. Re-rendered. AAPL + JNJ heroes verified flipped from "$298/$227 Composite avg" → "N/A — Insufficient inputs". |
| F-G | Missing-multiple penalty (kill ×1.30 rescale when P/E + P/FCF are bug-missing) | CARE | 3 | **✅ IMPLEMENTED 2026-05-24** (fmp_3statementv6.py `build_scorecard()` — `_fg_valuation_gap` flag + rescale suppression) | AAPL score drops 8.9→6.7 (rescale removed + fresh WACC). |
| F-E | Cyclical 5-yr median in tier classifier (banks deferred to F-D) | CARE | 4 | **✅ IMPLEMENTED 2026-05-24** (fmp_3statementv6.py `build_dcf()` — 5-yr YoY median for cyclicals) | CVX +0.2. FDX classified stable_compounder (not cyclical) by sector_bucket, so F-E didn't fire for FDX. |
| F-F | **Valuation concordance signal (NOT a score cap).** Tag `metrics["valuation_concordance"] = "expensive" / "cheap"` when both methods agree on >25% direction. Display-only — does NOT touch the quality score. | CARE | 5 | **✅ IMPLEMENTED 2026-05-24** (report_bridge.py concordance detection + red/green HTML banner) | COST, AMD, CSCO, WMT fire red "expensive" banner; ADBE, HCA, NFLX fire green "cheap" banner. |
| F-K | **Verdict text ladder enforced by concordance.** "High Conviction Buy" forbidden if `valuation_concordance == "expensive"`; "Avoid" forbidden if `valuation_concordance == "cheap"`. Quality score unchanged. | CARE | 5 | **✅ IMPLEMENTED 2026-05-24** (report_bridge.py `_VERDICT_RANK_FK` guard immediately after `_conservative_verdict()`) | COST: "High Conviction Buy" → "Hold — Premium Quality, Expensive". Score unchanged at 7.9. |
| F-C | Sector-specific EM multiples (biopharma 13×, software 19×, semis 18×, utilities 10×, staples 14×, healthcare 12×) | CARE | 6 | **DEFERRED** — requires all 30 tickers re-run (FMP API quota) | Will require full local_rerun.py pass. |
| F-A | WACC sanity floor | **SUPERSEDED by D-001** | — | **✅ Resolved via D-001** | D-001 is the simpler, blanket version. |
| F-D | Bank methodology (DDM + Justified P/B) | DANGEROUS — own session | 7 | **✅ Phase 1 IMPLEMENTED 2026-05-24** (fmp_3statementv6.py + report_bridge.py) | Phase 1: GG/EM force-disabled for bank-charter names. BAC/C/JPM/SOFI now show "N/A — Bank methodology pending". Phase 2 (DDM / Justified P/B) deferred. |

---

## Cross-fix coupling map

- **D-001 + F-B + F-H** must ship together: floored WACC may push GG/EM to "N/A" for some edge cases (TGR > floored WACC), and the new N/A handling + label work expects these together.
- **D-001 + F-F**: floored WACC will shift several tickers' GG/EM upsides. The concordance gate will then fire on different tickers than today. Recommend running D-001 first, scoring, then deciding on F-F thresholds based on the new distribution.
- **F-D Phase 1 + F-E**: once banks have GG/EM force-disabled, the bank-side of F-E is moot. F-E remaining work = cyclicals only.
- **F-G + D-001**: AAPL's score collapse will be driven by *both* (D-001 lowers GG/EM upside; F-G removes the ×1.30 rescale). Combined effect: AAPL likely drops from 8.9 to ~6.5–7.0. Verify before either fix lands solo.
- **F-C + F-D**: once banks return None from EM, F-C's "Banks→force-disable" branch is automatic via F-D Phase 1. Don't double-handle.

---

## Required follow-up steps after any fix lands

Per CLAUDE.md Rule 10 + Rule 1, every code change to `report_bridge.py`, `fmp_3statementv6.py::build_dcf`, or `fmp_3statementv6.py::build_scorecard` requires:
1. `python _rerender_reports.py` (re-render all 32 HTML reports from cached data)
2. `python _score_audit.py` (verify 0 mismatches CSV ↔ hero)
3. For DCF/WACC math changes: `python local_rerun.py` for affected tickers (re-runs engine, refreshes `static/data/*.json`)
4. Commit message must note the score_history.csv step-change date

---

## Proposed fixes from Batch 2 (COST, CSCO, C, CVX, DIS, F, FDX, HCA, INTC, JNJ)

| ID | Fix | Severity | Status | Coupling notes |
|---|---|---|---|---|
| F-I | **Block negative price targets.** If `gg_price < 0` or `em_price < 0`, clamp to `None`. | Critical | **✅ IMPLEMENTED 2026-05-24** (fmp_3statementv6.py L3010–3017 post-compute clamp) | Rule 5 violation. Pairs with F-B. Will take effect for Ford / Citi / SNAP-EM on next `local_rerun.py` (cached JSONs not yet refreshed). |
| F-J | **Detect captive finance arms.** When industry contains "Auto - Manufactur" / "Industrial" AND `D/E > 1.5` AND a known captive subsidiary exists, override D-weight to industry default 25–35% OR exclude captive-finance debt from net-debt subtraction. | High | **DEFERRED** — complex WACC override, own session | Affects Ford, GM, John Deere, Caterpillar. F-I clamps the symptom (negative GG/EM); F-J would fix the root cause. |
| F-K | **Verdict text must respect method concordance.** If both GG and EM <−25%, the verdict text cannot be "High Conviction Buy" or "Good Business at Fair Price" regardless of quality score. Hard guard in `_conservative_verdict()`. | High | **✅ IMPLEMENTED 2026-05-24** — see Batch 1 row. | Ships with F-F. COST is the textbook case. |
| F-L | **Widen EVS regime trigger.** Currently: `neg_earnings_regime AND trailing_EBITDA < 0`. New: `trailing_EBITDA < 5% of revenue` OR `FCF more negative than 10% of revenue`. | Medium | **✅ IMPLEMENTED 2026-05-24** (fmp_3statementv6.py `_ebitda_near_zero` + `_fcf_deeply_neg` flags) | INTC 2.1→0.9: EBITDA went negative with fresh data (confirmed correct; F-L's safety net adds additional coverage). |
| F-M | **Skip neg-earnings-regime check for banks.** Banks' "negative FCF" is accounting noise (loan/deposit flows). Banks should route to bank methodology (F-D) before any FCF/EBIT test. | High | **✅ IMPLEMENTED 2026-05-24** — ships with F-D Phase 1. | After bank force-disable: `_neg_earnings_regime = False` + `_evs_regime = False` + EVS prices set to None. |

---

## Cumulative coupling — additional notes after Batch 2

- **D-001 + F-B + F-H + F-I**: all four are the "stop publishing nonsense" bundle. Ship together.
- **F-I + F-J**: Ford's negative GG/EM is caused by the captive finance WACC bug. F-I clamps the symptom; F-J fixes the cause. Without F-J, Ford will show "N/A" forever after F-I.
- **F-D Phase 1 + F-M**: must implement together — F-M is the gate that prevents banks from entering the GG-disable code path before they get routed to P/B.
- **F-F + F-K**: scoring concordance gate (score-side) and verdict text override (verdict-side). Without F-K, F-F's score cap could still pair with a "Buy" verdict string from elsewhere in the bridge logic. Bundle them.
- **D-001 changes the input distribution for F-F**: floored WACC will shift several tickers' GG/EM upsides. Recommend running D-001 first → snapshot → then calibrate F-F thresholds on the new distribution if today's ±25% turns out to be too narrow.

---

## Final implementation order (post-Batch 2)

| Order | Fixes bundled | Re-render | Re-run tickers |
|---|---|---|---|
| 1 | F-B + F-H + F-I | yes | no (CSV unchanged) |
| 2 | D-001 (WACC 8.5% floor, Python + Excel) | yes | yes (AAPL, JNJ, KO, V, PEP, MSFT, CVX, HCA, COST possibly) |
| 3 | F-G (missing-multiple penalty) | yes | yes (AAPL, JNJ, any ticker with empty trailing P/E or P/FCF) |
| 4 | F-F + F-K (concordance gate + verdict override) | yes | yes (all 32) |
| 5 | F-E + F-L (cyclical tiers + widen EVS) | yes | yes (cyclicals + INTC + any neg-EBITDA name) |
| 6 | F-C (sector EM multiples) | yes | yes (all 32) |
| 7 | F-D Phase 1 + F-M (banks force-disable GG/EM + skip neg-FCF for banks) | yes | yes (all banks) |
| 8 | F-J (captive finance arm detection) | yes | yes (Ford, GM, DE, CAT) |

---

## Proposed fixes from Batch 3 (JPM, KO, META, MSFT, NFLX, NKE, NVDA, PEP, SNAP, SOFI)

| ID | Fix | Severity | Status | Coupling notes |
|---|---|---|---|---|
| F-N | **Data integrity triangulation: MktCap ≈ Price × Shares.** Validate `abs(marketCap − price × sharesOutstanding) / marketCap < 0.10` at engine entry. If `sharesOutstanding` is None, fall back to `is_data[-1].weightedAverageShsOut`. On mismatch: soft-warn (don't refuse to publish — could be legitimate split-day timing). | **Low (defensive)** | **✅ IMPLEMENTED 2026-05-24** (report_bridge.py L725–740 + banner L1196–1208) | Soft amber banner. Verified NFLX still renders cleanly (price × shares ≈ mktcap within tolerance, no banner triggers). |
| F-O | **`_secular_growth_subtype()` mis-classifies SNAP** as `secular_growth_deeptech` (4× multiple). SNAP is social/ad-tech → should be `tech_growth` (4.5×) or `secular_growth_software` (6×). | Low | **✅ IMPLEMENTED 2026-05-24** (fmp_3statementv6.py `_secular_growth_subtype()` — SNAP/PINS/RDDT/BMBL/MTCH → `tech_growth`) | SNAP EVS target ~$13.40 → ~$15.07 (+12.4%). Scorecard score unchanged. |

---

## Cumulative pattern after 25 reviews

**Three systematic failure modes account for all observed issues** — fix these three buckets and the universe normalises:

1. **WACC math broken on low-beta/low-debt majors** → AAPL, JNJ, KO, PEP; possibly V, MA, MSFT (borderline). **D-001 (WACC 8.5% floor) is the universal cure.**
2. **Method disagreement + score-method mismatch** → COST (7.9 vs methods −60%), PEP (8.0 vs −26%), CVX (4.9 vs +23%), META (7.6 vs +40%). **F-F (concordance gate) + F-K (verdict text guard) together resolve.**
3. **Bank / bank-like methodology missing** → BAC, C, JPM, SOFI (also WFC, GS not yet reviewed). **F-D Phase 1 force-disables for all bank-charter names.** Add a `hasBankCharter` flag in industry classifier so fintechs (SOFI) are caught alongside traditional banks.

Plus two narrower bug classes:
- **Negative price outputs** (C, F, SNAP-EM) → **F-I clamps.**
- **Data integrity / API quirks** → **F-N validates upstream** (defensive sanity check; NFLX false alarm was actually a legitimate stock split that the engine handled correctly).

---

## Batch log

- **Batch 1 (2026-05-24):** AAPL, ABBV, ADBE, AMD, BAC. Produced fixes A–H; D-001 locked in (WACC 8.5% floor).
- **Batch 2 (2026-05-24):** COST, CSCO, C, CVX, DIS, F, FDX, HCA, INTC, JNJ. Added F-I (neg price block), F-J (captive finance), F-K (verdict concordance), F-L (widen EVS), F-M (bank neg-FCF skip). JNJ confirmed as AAPL-clone. COST surfaced score-vs-methods as worst optical failure mode.
- **Batch 3 (2026-05-24):** JPM, KO, META, MSFT, NFLX, NKE, NVDA, PEP, SNAP, SOFI. Added F-N (data integrity validation), F-O (SNAP EVS subtype mis-classification). PEP is second textbook COST-pattern case. NFLX initially flagged as bug — confirmed by user it was a stock split; F-N downgraded to defensive only.
- **Batch 4 (2026-05-24):** TGT, TSLA, TSM, V, WMT. Added F-P (Critical — currency mismatch on TSM TWD reportedCurrency vs USD profile), F-Q (V/MA payment networks mis-grouped with banks for F-D), F-R (CSV/JSON/HTML divergence not audited — TSLA has 3 different valuations across sources), F-S (stale-data detection — TSLA JSON has 2 dcf_prices keys vs ~20 for others). Universe complete.

---

## Implementation log

### 2026-05-24 — Groups A + B + C landed

**Code changes:**
- `fmp_3statementv6.py::build_wacc()` L2016 + L2077–2095: D-001 WACC floor (Python + Excel mirror), new `wacc_raw` / `wacc_floored` keys.
- `fmp_3statementv6.py::build_dcf()` L2933–2954: F-P foreign-reporter guard, latest-year currency read fix.
- `fmp_3statementv6.py::build_dcf()` L3010–3017: F-I negative-price clamp.
- `fmp_3statementv6.py::build_dcf()` L3099–3145: dcf_prices dict propagates `wacc_raw`, `wacc_floored`, `foreign_reporter`, `reported_currency`, and None-safe upsides.
- `report_bridge.py::build_report_data()` L719–740: F-S stale-data assertion + F-N MktCap triangulation.
- `report_bridge.py::build_report_data()` L971: F-B remove current-price fallback in EV/Revenue branch.
- `report_bridge.py::build_report_data()` L1009–1025: F-H honest method label + F-B "N/A — Insufficient inputs".
- `report_bridge.py::build_report_data()` L1126–1145: F-B rationale override when price_target is None.
- `report_bridge.py::build_report_data()` L1176–1208: F-S + F-N banner injection into `_narrative_banner_html`.
- `report_bridge.py::build_report_data()` L1826–1829: PRICE_TARGET / PRICE_TARGET_VS_CURRENT None-safe rendering.
- `report_bridge.py::build_report_data()` L2130–2133: WACC_NOTE appends floor explanation when triggered.
- `CLAUDE.md`: Rules 11, 12, 13 added.

**Verification:**
- `python _rerender_reports.py` → 32 of 32 rendered OK, 0 failures.
- `python _score_audit.py` → 0 mismatches CSV ↔ hero (Rule 10 preserved).
- AAPL hero verified: "$298 / Composite avg / −0.1%" → **"N/A — Insufficient inputs"**.
- JNJ hero verified: "$227 / Composite avg / +0.1%" → **"N/A — Insufficient inputs"**.

**Not yet visible in reports (requires `local_rerun.py`):**
- D-001 WACC floor only takes effect when `build_wacc()` re-runs. Cached JSONs still hold pre-floor wacc_val.
- F-I negative-price clamp only fires when `build_dcf()` re-runs. C / Ford / SNAP-EM still show their old (broken) cached prices.
- F-P TSM currency guard only fires when `build_dcf()` re-runs. TSM still shows +5,550%.
- F-S stale-data banner fires immediately on render (TSLA banner now visible).
- F-N MktCap banner fires immediately on render (no current ticker triggers it — NFLX correctly clean).

**Next action:** re-run engine for affected tickers via `python local_rerun.py [TICKER]` to refresh cached JSONs and surface D-001 / F-I / F-P effects.

---

### 2026-05-24 — Groups D + E + F landed (F-G, F-O, F-D/M/Q, F-F/K, F-E, F-L, F-R + local_rerun CLI)

**Code changes:**

`fmp_3statementv6.py`:
- `build_scorecard()`: F-G — `_fg_valuation_gap` flag initialised before `if _scored:` block; rescale suppressed when both `tier_pe is None` and `tier_pfcf is None` on a standard company (not bank, not EVS). Confidence note updated. Fixes silent ×1.30 inflation on data-fetch failures.
- `build_dcf()` — F-D/M/Q: `_BANK_DCF_EXCLUDE` set (`{V, MA, PYPL, FIS, FISV, GPN, WU, DFS, TRMK}`); `_BANK_DCF_KW` industry keywords; `_is_bank_dcf` detection; when bank: GG/EM/EVS all set to None + `_neg_earnings_regime = False` + `_evs_regime = False` (F-M); `dcf_prices` gains `bank_disabled` (bool) + `bank_disabled_reason` (str).
- `build_dcf()` — F-E: cyclical companies (detected via `_sector_bucket()`) use 5-year YoY revenue growth **median** instead of 3-year average for `_rev_3yr_avg_dcf`. Falls back to non-smoothed if fewer than 3 valid periods.
- `build_dcf()` — F-L: EVS regime now also triggers on `trailing_EBITDA < 5% of revenue` (`_ebitda_near_zero`) OR `trailing_FCF < -10% of revenue` (`_fcf_deeply_neg`), both excluding cyclicals. Widens coverage beyond pure `EBITDA < 0`.
- `_secular_growth_subtype()` — F-O: SNAP/PINS/RDDT/BMBL/MTCH reclassified from `secular_growth_deeptech` (4×) to `tech_growth` (4.5×). Lifts SNAP EVS price target $13.40 → $15.07.

`report_bridge.py`:
- F-D: `_bank_disabled` flag read from `dcf_prices`; bank method label overridden to "N/A — Bank methodology pending (DDM / Justified P/B)"; F-H `else` branch protected from overwriting bank label; bank rationale block injected.
- F-F: concordance detection — `_valuation_concordance` computed from `gg_upside` and `em_upside`; fires only when both >+25% ("cheap") or both <-25% ("expensive") and not in EVS/bank mode. Red/green HTML banner injected below hero.
- F-K: `_VERDICT_RANK_FK` dict + verdict ladder guard immediately after `_conservative_verdict()`. "expensive" concordance caps verdict at "Hold — Premium Quality, Expensive". "cheap" floors at "Good Business at Fair Price".

`local_rerun.py`:
- CLI args support: `python local_rerun.py AAPL JNJ` runs only those tickers; `python local_rerun.py` (no args) runs all. Unknown tickers are warned and skipped.

`_score_audit.py`:
- F-R: full rewrite. Added `--full` flag for GG/EM/PT field diff. Reads `dcf_prices.gg_price` + `dcf_prices.em_price` from JSON, compares against CSV `GG_Price` / `EM_Price`. Reports diffs >$1. HTML `price-value price-target` also captured for triangulation.

**Engine re-runs (FMP API calls fired):**
```
python local_rerun.py AAPL JNJ BAC C JPM SOFI SNAP
python local_rerun.py CVX FDX F INTC
```

**Verification:**
- `python _score_audit.py` → **0 mismatches / 30** (Rule 10 preserved throughout).
- All 32 HTML reports re-rendered.

**Before/after score changes:**

| Ticker | Before | After | Δ | Primary driver |
|--------|--------|-------|---|----------------|
| AAPL | 8.9 | 6.7 | −2.2 | F-G rescale suppressed + fresh WACC (D-001 on re-run) |
| JNJ | 8.3 | 7.7 | −0.6 | Fresh data with D-001 WACC floor |
| BAC | 5.8 | 6.1 | +0.3 | F-D weight redistribution (no GG/EM in score) |
| C | 4.1 | 4.8 | +0.7 | F-D weight redistribution |
| CVX | 4.9 | 5.1 | +0.2 | F-E cyclical 5yr median smoothing |
| F | 4.2 | 3.9 | −0.3 | Combined F-D/F-E on re-run |
| FDX | 6.1 | 5.7 | −0.4 | Fresh data factors (FDX = stable_compounder, F-E didn't fire) |
| INTC | 2.1 | 0.9 | −1.2 | F-L: EBITDA negative in fresh data → EVS regime triggered |
| JPM | 6.5 | 6.1 | −0.4 | F-D weight redistribution |
| SNAP | 3.2 | 3.2 | 0.0 | F-O lifts EVS price $13.40→$15.07; scorecard unchanged |
| SOFI | 4.3 | 3.9 | −0.4 | F-D weight redistribution |
| All others | — | — | 0.0 | No engine re-run; report_bridge changes only |

**Verdict text changes (F-K):**
| Ticker | Before | After |
|--------|--------|-------|
| COST | "High Conviction Buy" | "Hold — Premium Quality, Expensive" |

**New concordance banners displayed (F-F — score unchanged):**
- 🔴 Expensive: COST (GG −70%, EM −48%), AMD (GG −65%, EM −40%), CSCO (GG −45%, EM −55%), WMT (GG −62%, EM −26%)
- 🟢 Cheap: ADBE (GG +45%, EM +77%), HCA (GG +46%, EM +102%), NFLX (GG +61%, EM +89%)

**Price target / method changes (F-D):**
| Ticker | Before | After |
|--------|--------|-------|
| BAC | "$36.46 / Gordon Growth" | "N/A — Bank methodology pending (DDM / Justified P/B)" |
| C | "−$89.60 / EV/EBITDA Exit" | "N/A — Bank methodology pending" |
| JPM | "$199.77 / Gordon Growth" | "N/A — Bank methodology pending" |
| SOFI | "$24.56 / EV/EBITDA Exit" | "N/A — Bank methodology pending" |

V correctly excluded by F-Q — still shows $300 price target.

**Still deferred:**
- F-C (sector EM multiples) — requires full 30-ticker re-run, FMP quota cost
- F-J (captive finance arm detection for Ford/GM/DE/CAT) — complex WACC override, own session
- F-D Phase 2 (DDM / Justified P/B implementation) — major architecture work
- D-001 effect on KO, PEP, V, TGT, WMT, HCA, TSLA — cached JSONs still pre-floor; run `python local_rerun.py KO PEP V TGT WMT HCA TSLA` to refresh

---

## Proposed fixes from Batch 4

| ID | Fix | Severity | Status | Coupling notes |
|---|---|---|---|---|
| F-P | **Currency normalization for non-USD reporters.** Detect `reportedCurrency != USD` at engine entry; if FX fetch fails, refuse to publish prices. | **Critical** | **✅ IMPLEMENTED 2026-05-24** (fmp_3statementv6.py L2933–2954: latest-year currency read + `_fx_fetched` flag + `_foreign_reporter_unsupported` propagated to dcf_prices) | TSM: F-P + F-I together will void both GG/EM on next `local_rerun.py TSM`. Sister bug fixed too: previous code read `is_data[0]` (oldest year), now correctly reads `is_data[-1]` (latest). Phase 2 (full FX integration) deferred. |
| F-Q | **Exclude payment networks from bank classifier.** V/MA/PYPL/FIS/FISV share `Financial - Credit Services` FMP tag with SOFI. F-D Phase 1 would wrongly disable them. Add explicit exclude list OR derive `hasBankCharter` from balance-sheet deposits > 0. | High | **✅ IMPLEMENTED 2026-05-24** (fmp_3statementv6.py `_BANK_DCF_EXCLUDE` set) | V still renders full $300 price target. Ships with F-D. |
| F-R | **Extend audit script to cover GG_Price / EM_Price / Price_Target divergence.** Rule 10 protects only Auto_Score. TSLA shows 3 different values across CSV, JSON, HTML. Add field-by-field diff to `_score_audit.py`. | High | **✅ IMPLEMENTED 2026-05-24** (`_score_audit.py` full rewrite — adds `--full` mode and GG/EM/PT diff audit) | Run `python _score_audit.py --full` to see all valuation fields. |
| F-S | **Stale-data assertion at report entry.** Detect missing core keys (`growth_tier`, `tgr_base`, `em_base_mult`, `neg_earnings_regime`). Render report with explicit red banner directing user to run `local_rerun.py`. | Medium | **✅ IMPLEMENTED 2026-05-24** (report_bridge.py L719–724 detection + L1176–1189 banner) | Banner surfaces in TSLA report instructing rerun. Doesn't block render entirely so user can still see whatever was cached. |

---

## Portfolio-wide pattern summary (all 30 reports complete)

**Five systematic failure modes account for 100% of issues observed:**

| Mode | Affected tickers | Fix | Severity |
|---|---|---|---|
| 1. WACC math broken on low-beta/low-debt majors (reduces below 8.5%) | AAPL, JNJ, KO, PEP, V, TGT, WMT, HCA, CVX | **D-001** | High |
| 2. Score-vs-methods directional mismatch | COST 7.9 vs −60%, PEP 8.0 vs −26%, CVX 4.9 vs +23%, META 7.6 vs +40%, TGT 7.5 vs −1% | **F-F + F-K** | High |
| 3. Bank / bank-like methodology missing | BAC, C, JPM, SOFI (and WFC, GS not in current set; V/MA falsely grouped — see F-Q) | **F-D Phase 1 + F-Q** | High |
| 4. Negative price outputs | C, F, SNAP-EM | **F-I** | Critical |
| 5. Currency / data integrity | TSM (TWD), TSLA (stale JSON), NFLX (false alarm — was a split) | **F-P + F-R + F-S** | Critical (TSM) |

**Six reports are well-calibrated and require no fix to render correctly:** ADBE, AMD, DIS, NKE, NVDA, SNAP. All 6 share: (a) WACC above 8.5%, (b) GG and EM concordant in direction, (c) verdict matches method consensus.

**Two reports got a directionally-right verdict for the wrong reasons** (Rule 10 holds but underlying math broken): JPM, BAC.

**One ticker is currently unpublishable as-is:** TSM (5,550% upside is the worst optical fail in the entire universe; trumps even Ford's $-33 GG and Citi's $-90 EM).

---

## Final implementation order (post-Batch 4 — supersedes earlier order)

Re-prioritised after seeing the full universe:

| Order | Fixes bundled | Rationale |
|---|---|---|
| **1** | **F-P** (currency norm) + **F-S** (stale-data assert) | TSM is unpublishable today; F-S prevents a re-emergence of partial-state renders. Both are engine-entry guards — same insertion point. |
| **2** | **F-I + F-B + F-H** (block negative prices + N/A fallback + method label) | All "stop publishing nonsense" guards. SAFE per architect review. |
| **3** | **D-001** (WACC 8.5% floor, Python + Excel mirror) | Cures 9 tickers' WACC math (AAPL, JNJ, KO, PEP, V, TGT, WMT, HCA, CVX). Required before F-F/F-K so concordance gate calibrates on corrected upsides. |
| **4** | **F-G** (missing-multiple penalty) | Stops the ×1.30 rescale from inflating AAPL/JNJ post-D-001. |
| **5** | **F-F + F-K** (concordance gate + verdict text override) | Highest analytical value — fixes COST 7.9/PEP 8.0/META 7.6 directional mismatches. Must mirror gate to Excel verdict cell. |
| **6** | **F-E + F-L** (cyclical tiers + widen EVS) | Refines edge-case classifications. CVX/FDX/F + INTC. |
| **7** | **F-C** (sector EM multiples) | After core math stable; sharpens biopharma/staples/healthcare/software. |
| **8** | **F-D Phase 1 + F-Q + F-M** (banks force-disable + payment-network exclude + bank neg-FCF skip) | Must ship together — F-Q prevents F-D from wrongly disabling V/MA. |
| **9** | **F-J** (captive finance arm detection) | Ford-specific, smaller scope. |
| **10** | **F-N + F-R** (data integrity validation + audit extension) | Defensive; ship last to avoid noise during earlier fixes. |
| **11** | **F-O** (SNAP EVS subtype) | Trivial polish. |

**Estimated post-fix score deltas (rough):**
- AAPL 8.9 → 6.5–7.0 (D-001 + F-G drop the rescale + lower GG)
- JNJ 8.3 → 6.5–7.0 (same pattern)
- COST 7.9 → 4.5 (F-F caps; F-C may partially offset)
- PEP 8.0 → 4.5 (F-F caps)
- CVX 4.9 → 7.5 (F-F floors on positive concordance)
- TGT 7.5 → 6.5 (F-K softens verdict on flat composite)
- TSM 6.8 → N/A (F-P refuses publish)
- All banks (BAC, C, JPM) → "N/A — bank methodology pending" with score from quality only
- ~10 other tickers see ±5–15pp upside shifts from D-001

After full implementation, expect 28 of 30 reports to render with publishable, internally-consistent valuations. Banks and TSM stay "N/A — methodology pending" until Phase 2 of F-D and F-P respectively.

---

### 2026-05-25 — F-C Phase 2: Historical EV/EBITDA Anchoring + F-C Phase 1 ROIC Quality Premium

**Context:** Both KO and ABBV showed systematically low EM prices (−48% and −39% respectively) because the low-growth tier default of 10× is calibrated to a sector median — it ignores that KO has traded at 19–22× and ABBV at 19–25× EV/EBITDA for years. A second companion issue: COST/V showed inflated "bubble-level" overvaluation from GG perpetuating thin FCF margins. Session resolved both.

**Code changes:**

`fmp_3statementv6.py::build_dcf()`:
- **F-C Phase 1 — ROIC quality premium:** After tier block, compute trailing ROIC (NOPAT / Invested Capital). When `_is_stable_compounder_dcf AND NOT _is_bank_dcf AND trailing_ROIC > 25%`: add +5× base, +4× bear, +5× bull to EM multiples. Adds `quality_em_premium` (bool) + `fcf_margin_trailing` (float) to `dcf_prices`. Currently fires for COST (37.5% ROIC) and V (46.3% ROIC).
- **F-C Phase 2 — Historical anchor:** Immediately after quality premium block, fetch FMP `/stable/ratios?symbol={ticker}&limit=5` to get `enterpriseValueMultiple` per year. Compute 5yr average, apply 80% mean-reversion discount, floor at current (post-premium) tier base, cap at 28×. Scale bear/bull multiples proportionally (bull additionally capped at 32×). Adds `em_anchored` (bool) + `em_hist_anchor_raw` (float) + `em_hist_anchor_capped` (bool) to `dcf_prices`. Excluded for cyclicals (F-E smoothing is better anchor), banks (EM disabled), EVS regime (EM prices become None regardless).

`report_bridge.py`:
- **Thin-margin primary override (F-C complement):** When `fcf_margin_trailing < 5%` AND `sector_bucket == stable_compounder` AND not bank/EVS: make EM (not GG) the primary price target. Rationale block explains GG perpetuates thin margins into perpetuity. Applies to COST (2.85%), WMT (2.09%), FDX (3.39%), TGT (2.7%).
- **Cap banner:** When `em_hist_anchor_capped = True`, renders amber callout below hero: "⚠ Exit Multiple Capped at 28× (5yr historical avg: X×). For businesses with extraordinary competitive position, use the DCF Calculator." Provides explicit escape hatch to `/dcf` for manual override.

**JSON patches (FMP rate-limited — manual until next re-run):**

| Ticker | Patch | Old base | New base | Old EM% | New EM% | Capped? |
|--------|-------|----------|----------|---------|---------|---------|
| COST | Phase 1 quality premium | 15× | 20× | −48% | −34% | No |
| V | Phase 1 quality premium | 15× | 20× | −10% | +14% | No |
| WMT | Phase 1 thin-margin flag | — | — | −62% primary | −26% (EM now primary) | No |
| FDX | Phase 1 thin-margin flag | — | — | −48% primary | −2% (EM now primary) | No |
| KO | Phase 2 anchor | 10× | 15.4× | −48% | −19% | No |
| ABBV | Phase 2 anchor | 10× | 15.6× | −39% | −5% | No |
| AAPL | Phase 2 anchor | 10× | 18.4× | −56% | −20% | No |
| NKE | Phase 2 anchor | 10× | 19.0× | −13% | +67% | No |
| PEP | Phase 2 anchor | 10× | 13.5× | −32% | −8% | No |
| DIS | Phase 2 anchor | 10× | 16.1× | +5% | +69% | No |
| AMD | Phase 2 anchor (capped) | 18× | 28.0× | −40% | −6% | **YES** |
| NVDA | Phase 2 anchor (capped) | 18× | 28.0× | +28% | +99% | **YES** |
| ADBE | Phase 2 anchor | 15× | 23.3× | +77% | +175% | No |

**Verification:**
- `python _rerender_reports.py` → **32 of 32 OK, 0 failures.**
- `python _score_audit.py` → **0 mismatches / 30.** Rule 10 intact.
- Cap banner verified in AMD and NVDA HTML (line 616 of each report).
- KO: GG $65 (−19%) / EM $65 (−19%) — both methods now perfectly aligned.
- ABBV: GG $238 (+13%) / EM $200 (−5%) — split narrowed from 52pp to 18pp.
- ADBE: concordance "cheap" now fires (GG +45% AND EM +175%).
- AMD: concordance "expensive" REMOVED (EM now −6%, both methods no longer in agreement on direction).

**CLAUDE.md:** Rule 21 added.

**Still pending (next FMP quota cycle):**
- Full engine re-run for all patched tickers to apply anchoring via engine path (not manual patch).
- D-001 stale JSONs: KO, PEP, TGT, HCA, TSLA still have pre-floor WACC in cached JSONs.
- MSFT (+1.0×), CSCO (+1.4×) — borderline anchors; will apply on next re-run, not worth manual patch.
