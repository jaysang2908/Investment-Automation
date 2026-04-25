"""
report_bridge.py
Maps raw FMP API data + engine outputs → HTML report template DATA dict,
then renders the template to an HTML string.

Auto-fills: all financial tables, ratios, DCF prices, WACC, credit, scorecard tiers.
Stubs:      thesis text, analyst consensus, segment descriptions — add manually after.
"""

import re
import os
import datetime

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "Report_Template.html")

# ── Formatters ────────────────────────────────────────────────────────────────

def _m(v):
    """Raw dollar → millions string, no decimals."""
    if v is None: return "N/A"
    m = v / 1e6
    if m < 0: return f"({abs(m):,.0f})"
    return f"{m:,.0f}"

def _b(v, dp=1):
    """Raw dollar → billions string like '$12.3B'."""
    if v is None: return "N/A"
    b = v / 1e9
    sign = "-" if b < 0 else ""
    return f"{sign}${abs(b):.{dp}f}B"

def _pct(v, dp=1):
    if v is None: return "N/A"
    return f"{v*100:.{dp}f}%"

def _x(v, dp=1):
    if v is None: return "N/A"
    return f"{v:.{dp}f}x"

def _js_arr(lst, dp=1):
    return "[" + ",".join(str(round(v, dp)) if v is not None else "0" for v in lst) + "]"

def _js_str_arr(lst):
    return "[" + ",".join(f"'{v}'" for v in lst) + "]"

def _delta(current, avg):
    """'+12%' or '−15%' delta vs average."""
    if current is None or avg is None or avg == 0: return "N/A"
    d = (current - avg) / avg
    sign = "+" if d >= 0 else "−"
    return f"{sign}{abs(d)*100:.0f}%"

def _vs(px, current_price):
    """'+34%' upside/downside vs current price."""
    if not px or not current_price: return "N/A"
    d = (px - current_price) / current_price
    sign = "+" if d >= 0 else "−"
    return f"{sign}{abs(d)*100:.0f}%"

# ── Financial helpers ─────────────────────────────────────────────────────────

def _roic(is_, bs_):
    ebit   = abs(is_.get("operatingIncome") or 0)
    tax_e  = abs(is_.get("incomeTaxExpense") or 0)
    pretax = abs(is_.get("incomeBeforeTax") or 1e-9)
    nopat  = ebit * (1 - min(tax_e / pretax, 0.50))
    equity = bs_.get("totalStockholdersEquity") or 0
    debt   = (bs_.get("shortTermDebt") or 0) + (bs_.get("longTermDebt") or 0)
    cash   = bs_.get("cashAndCashEquivalents") or 0
    ic     = equity + debt - cash
    return (nopat / ic) if ic > 1 else None

def _ebitda(is_, cf_):
    v = is_.get("ebitda") or 0
    if v: return v
    da = abs(is_.get("depreciationAndAmortization") or
             cf_.get("depreciationAndAmortization") or 0)
    return (is_.get("operatingIncome") or 0) + da

# ── Tier scoring (mirrors fmp_3statementv6 thresholds) ───────────────────────

def _t(v, thresholds):
    """thresholds: [(value, label), ...] in descending order."""
    if v is None: return "MOD"
    for threshold, label in thresholds:
        if v >= threshold: return label
    return "LOW"

def _tier_rev_cagr(v):
    return _t(v, [(0.12, "HIGH"), (0.05, "MOD")])

def _tier_roic(v):
    return _t(v, [(0.20, "HIGH"), (0.08, "MOD")])

def _tier_fcf_ni(v):
    return _t(v, [(0.80, "HIGH"), (0.40, "MOD")])

def _tier_d_ebitda(v):
    if v is None: return "MOD"
    if v < 1.5: return "HIGH"
    if v < 4.0: return "MOD"
    return "LOW"

def _tier_ebit_int(v):
    if v is None: return "HIGH"   # no debt → no interest → fortress
    if v > 8.0: return "HIGH"
    if v > 3.0: return "MOD"
    return "LOW"

def _tier_pe(current, avg):
    if current is None or avg is None or avg == 0: return "MOD"
    r = current / avg
    if r < 0.90: return "HIGH"
    if r < 1.10: return "MOD"
    return "LOW"

def _tier_pfcf(current, avg):
    if current is None or avg is None or avg == 0: return "MOD"
    r = current / avg
    if r < 0.90: return "HIGH"
    if r < 1.10: return "MOD"
    return "LOW"

TIER_PTS = {"HIGH": 10, "MOD": 7, "LOW": 0}

# ── Credit ────────────────────────────────────────────────────────────────────

def _credit_tier(rating):
    r = str(rating).upper()
    if r == "AAA":               return 1
    if r in ("AA+","AA","AA-"):  return 2
    if r in ("A+","A","A-"):     return 3
    if r in ("BBB+","BBB","BBB-"): return 4
    return 10

def _credit_note(sp, moody, tier, d_ebd, ebit_int, net_str):
    if tier <= 2:
        return (f"{sp}/{moody} ratings reflect exceptional credit quality — fortress balance sheet. "
                f"Leverage at {d_ebd} D/EBITDA; interest coverage {ebit_int} provides substantial cushion. "
                f"{net_str} enhances financial flexibility.")
    if tier <= 4:
        return (f"{sp}/{moody} ratings indicate strong credit quality. "
                f"Leverage of {d_ebd} D/EBITDA is conservative; interest coverage {ebit_int}. "
                f"{net_str}.")
    return f"Credit metrics: {d_ebd} D/EBITDA, {ebit_int} interest coverage. {net_str}."

# ── Verdict ───────────────────────────────────────────────────────────────────

def _verdict(score):
    if score is None: return "Analysis Pending"
    if score >= 75:   return "High Conviction Buy"
    if score >= 65:   return "Good Business at Fair Price"
    if score >= 50:   return "Hold — Monitor"
    return "Avoid"

# ── CSS class helpers (mirrors pipeline.py) ───────────────────────────────────

def _score_class(text):
    t = str(text).upper()
    if "HIGH" in t: return f'<span style="color: #10b981; font-weight: bold;">{text}</span>'
    if "MOD"  in t: return f'<span style="color: #f59e0b; font-weight: bold;">{text}</span>'
    if "LOW"  in t: return f'<span style="color: #dc2626; font-weight: bold;">{text}</span>'
    return text

def _sensitivity_class(current_px, calc_px):
    if calc_px < current_px:           return "low-range"
    if calc_px <= current_px * 1.10:   return "mid-range"
    return "near-market"

def _compute_css(d, current_price):
    c = {}
    score = float(str(d.get("FINAL_SCORE", 0)).replace(",", "") or 0)
    c["VERDICT_CLASS"]     = "verdict-green" if score >= 70 else ("verdict-amber" if score >= 50 else "verdict-red")
    c["SCORE_FINAL_CLASS"] = "score-final-green" if score >= 70 else ("score-final-amber" if score >= 50 else "score-final-red")

    def dcol(s):
        return "#00695c" if ("−" in str(s) or str(s).startswith("-")) else "#b45309"

    for k in ["TRAILING_PE_DELTA","FORWARD_PE_DELTA","TRAILING_PFCF_DELTA","FORWARD_PFCF_DELTA"]:
        c[k.replace("_DELTA","_DELTA_COLOR")] = dcol(d.get(k,""))

    c["TRAILING_PE_CELL_CLASS"]  = "warn" if "+" in str(d.get("TRAILING_PE_DELTA",""))  else ""
    c["FORWARD_PE_CELL_CLASS"]   = "" if "+" in str(d.get("FORWARD_PE_DELTA","")) else "highlight"
    c["FORWARD_PFCF_CELL_CLASS"] = "positive" if "−" in str(d.get("FORWARD_PFCF_DELTA","")) else ""
    c["REVENUE_VALUE_CLASS"] = "positive"
    c["FCF_VALUE_CLASS"]     = "positive"
    try:
        roic = float(str(d.get("ROIC_VALUE","0")).replace("%",""))
        c["ROIC_VALUE_CLASS"] = "positive" if roic >= 20 else ("" if roic >= 10 else "warn")
    except ValueError:
        c["ROIC_VALUE_CLASS"] = ""

    def rcls(r):
        r = str(r).upper()
        if r in ("NR","N/A",""): return "rating-muted"
        if any(r.startswith(p) for p in ("AAA","AA","A1","A2","A3","Aaa","Aa","A")): return "rating-green"
        if any(x in r for x in ("BBB","BAA","Baa")): return "rating-amber"
        return "rating-red"

    c["SP_RATING_CLASS"]     = rcls(d.get("SP_RATING","NR"))
    c["MOODYS_RATING_CLASS"] = rcls(d.get("MOODYS_RATING","NR"))
    c["FITCH_RATING_CLASS"]  = rcls(d.get("FITCH_RATING","NR"))

    def pcls(t):
        t = str(t).upper()
        if "HIGH" in t: return "score-high"
        if "MOD"  in t: return "score-mid"
        return "score-low"

    for k in [k for k in d if k.endswith("_SCORE_TEXT")]:
        c[k.replace("_TEXT","_CLASS")] = pcls(d[k])

    c["TRAIL_PE_COLOR"]   = dcol(d.get("TRAILING_PE_DELTA",""))
    c["FWD_PE_COLOR"]     = dcol(d.get("FORWARD_PE_DELTA",""))
    c["TRAIL_PFCF_COLOR"] = dcol(d.get("TRAILING_PFCF_DELTA",""))
    c["FWD_PFCF_COLOR"]   = dcol(d.get("FORWARD_PFCF_DELTA",""))

    for scenario in ["BEAR","BASE","BULL"]:
        vs = str(d.get(f"DCF_{scenario}_VS","+0%"))
        c[f"DCF_{scenario}_PX_CLASS"] = "price-up" if "+" in vs else "price-dn"

    BUY  = {"BUY","OVERWEIGHT","STRONG BUY","OUTPERFORM","MARKET OUTPERFORM"}
    HOLD = {"HOLD","NEUTRAL","MARKET PERFORM","EQUAL WEIGHT"}

    def arcls(t):
        t = str(t).upper()
        if any(b in t for b in BUY):  return "rating-buy"
        if any(h in t for h in HOLD): return "rating-hold"
        return "rating-sell"

    def ptcls(vs):
        s = str(vs)
        if "+" in s: return "pt-above"
        if "-" in s or "−" in s: return "pt-below"
        return "pt-flat"

    for i in range(1, 8):
        if f"A{i}_RATING_TEXT" in d:
            c[f"A{i}_RATING_CLASS"] = arcls(d[f"A{i}_RATING_TEXT"])
        if f"A{i}_PT_VS" in d:
            c[f"A{i}_PT_CLASS"] = ptcls(d[f"A{i}_PT_VS"])

    return c


# ══════════════════════════════════════════════════════════════════════════════
# MAIN: build DATA dict
# ══════════════════════════════════════════════════════════════════════════════

def build_report_data(ticker, profile, is_data, bs_data, cf_data, years,
                      wacc_val, dcf_prices, scorecard_metrics,
                      manual_rating=None, current_price=None, market_cap=None,
                      biz_clarity=None, ltp=None, adj_score=None,
                      analyst_ests=None):

    is0, bs0, cf0 = is_data[-1], bs_data[-1], cf_data[-1]
    today = datetime.date.today().strftime("%B %Y")
    dcf_prices = dcf_prices or {}

    # ── Company info ──────────────────────────────────────────────────────────
    company_name = profile.get("companyName") or ticker
    exchange     = profile.get("exchangeShortName") or ""
    industry     = profile.get("industry") or profile.get("sector") or "N/A"
    ceo          = profile.get("ceo") or "N/A"
    description  = profile.get("description") or ""
    ceo_info     = f"CEO: {ceo}."

    current_price = current_price or float(profile.get("price") or 0) or 0.0
    market_cap    = market_cap    or float(profile.get("mktCap") or 0) or 0.0

    shares = float(profile.get("sharesOutstanding") or 0)
    shares_str = (f"{shares/1e9:.2f}B shares" if shares > 1e9 else
                  f"{shares/1e6:.1f}M shares" if shares > 1e6 else str(int(shares)))

    ticker_line = f"{exchange}: {ticker}" if exchange else ticker

    # ── Balance sheet ─────────────────────────────────────────────────────────
    cash0  = bs0.get("cashAndCashEquivalents") or 0
    std0   = bs0.get("shortTermDebt") or 0
    ltd0   = bs0.get("longTermDebt") or 0
    debt0  = std0 + ltd0
    eq0    = bs0.get("totalStockholdersEquity") or 0
    net_cash_val = cash0 - debt0
    ev = (market_cap + debt0 - cash0) if market_cap else None

    if net_cash_val >= 0:
        net_cash_str = f"Net cash {_b(net_cash_val)}"
    else:
        net_cash_str = f"Net debt {_b(abs(net_cash_val))}"

    # ── Income statement ──────────────────────────────────────────────────────
    rev0    = is0.get("revenue") or 0
    gp0     = is0.get("grossProfit") or 0
    ebd0    = _ebitda(is0, cf0)
    ebit0   = is0.get("operatingIncome") or 0
    ni0     = is0.get("netIncome") or 0
    int0    = abs(is0.get("interestExpense") or 0)
    intang0 = bs0.get("goodwillAndIntangibleAssets") or bs0.get("intangibleAssets") or 0

    gm0      = gp0 / rev0    if rev0 else None
    ebitdam0 = ebd0 / rev0   if rev0 else None
    d_ebd    = debt0 / ebd0  if ebd0 > 0 else None
    ebit_int = abs(ebit0) / int0 if int0 > 0 else None

    # ── Cash flow ─────────────────────────────────────────────────────────────
    ocf0   = cf0.get("operatingCashFlow") or 0
    capex0 = abs(cf0.get("capitalExpenditure") or 0)
    fcf0   = cf0.get("freeCashFlow") or (ocf0 - capex0)

    # ── Scorecard metrics ─────────────────────────────────────────────────────
    roic_v        = scorecard_metrics.get("roic")
    rev_cagr_v    = scorecard_metrics.get("rev_cagr")
    fcf_ni_v      = scorecard_metrics.get("fcf_ni")
    d_ebitda_v    = scorecard_metrics.get("d_ebitda")
    equity_assets_v = scorecard_metrics.get("equity_assets")
    is_bank_v     = scorecard_metrics.get("is_bank", False)
    auto_score    = scorecard_metrics.get("auto_score")
    trailing_pe   = scorecard_metrics.get("pe_current")
    pe_5yr        = scorecard_metrics.get("pe_5yr_avg")
    trailing_pfc  = scorecard_metrics.get("pfcf_current")
    pfcf_5yr      = scorecard_metrics.get("pfcf_5yr_avg")

    # ── YoY growth ────────────────────────────────────────────────────────────
    def _yoy(series, key):
        if len(series) < 2: return None
        cur = series[-1].get(key) or 0
        prv = series[-2].get(key) or 0
        return (cur / prv - 1) if prv else None

    rev_yoy = _yoy(is_data, "revenue")
    fcf_prev = (cf_data[-2].get("freeCashFlow") or
                ((cf_data[-2].get("operatingCashFlow") or 0) - abs(cf_data[-2].get("capitalExpenditure") or 0))
                ) if len(cf_data) >= 2 else None
    fcf_yoy  = (fcf0 / fcf_prev - 1) if fcf_prev else None

    # ── Valuation multiples ────────────────────────────────────────────────────
    ev_ebitda = (ev / ebd0) if ev and ebd0 > 0 else None
    pe_delta  = _delta(trailing_pe, pe_5yr)
    pfcf_delta = _delta(trailing_pfc, pfcf_5yr)

    # ── DCF prices ────────────────────────────────────────────────────────────
    gg_px  = dcf_prices.get("gg_price")
    em_px  = dcf_prices.get("em_price")
    base_px = gg_px or em_px

    wacc_b = wacc_val or 0.09
    tgr_base = 0.030; tgr_bear = 0.020; tgr_bull = 0.035
    w_base = wacc_b;  w_bear = wacc_b + 0.01; w_bull = wacc_b - 0.01

    def _dcf_px(base, w_new, tgr_new):
        spread_base = w_base - tgr_base
        spread_new  = w_new  - tgr_new
        if spread_base > 0 and spread_new > 0 and base:
            return round(base * spread_base / spread_new, 0)
        return None

    bear_px = _dcf_px(base_px, w_bear, tgr_bear)
    bull_px = _dcf_px(base_px, w_bull, tgr_bull)

    # ── WACC components (approximate from available data) ──────────────────────
    RF_APPROX  = 0.043   # 10yr Treasury approx
    ERP_APPROX = 0.045   # Damodaran avg
    beta_v = float(profile.get("beta") or 1.0) or 1.0
    ke_approx = RF_APPROX + beta_v * ERP_APPROX
    # Kd approximate: wacc = E/V * Ke + D/V * Kd*(1-t); solve for Kd
    pti0 = is0.get("incomeBeforeTax") or is0.get("pretaxIncome") or 0
    te0  = abs(is0.get("incomeTaxExpense") or 0)
    eff_tax0 = te0 / pti0 if pti0 > 0 else 0.21
    cap_total = market_cap + debt0 if market_cap else max(debt0, 1)
    ew_v = market_cap / cap_total if market_cap else 1.0
    dw_v = debt0 / cap_total if debt0 else 0.0
    kd_pre = (wacc_b - ew_v * ke_approx) / (dw_v * (1 - eff_tax0)) if dw_v > 0.001 else RF_APPROX * 0.75

    # ── EV Bridge (approximate from base_px × shares + net debt) ──────────────
    shares_v = float(profile.get("sharesOutstanding") or 0)
    eq_val_approx  = (base_px or 0) * shares_v if base_px and shares_v else None
    ev_approx      = (eq_val_approx + net_cash_val * -1) if eq_val_approx is not None else None
    # TV typically ~60-70% of EV in most DCF; rough 65%
    pv_tv_approx   = eq_val_approx * 0.65 if eq_val_approx else None
    pv_fcfs_approx = (ev_approx - pv_tv_approx) if ev_approx and pv_tv_approx else None

    pt_vals = [p for p in [gg_px, em_px] if p]
    price_target = round(sum(pt_vals) / len(pt_vals), 0) if pt_vals else current_price

    # ── Reverse DCF: implied perpetuity FCF growth rate at current price ──────
    _mkt_cap_v = market_cap or (current_price * shares if current_price and shares else 0)
    if current_price and _mkt_cap_v > 0 and wacc_b > 0:
        if fcf0 > 0:
            _impl_g = wacc_b - fcf0 / _mkt_cap_v
            _impl_g = max(-0.15, min(_impl_g, wacc_b - 0.005))
            _gap    = _impl_g - (rev_cagr_v or 0)
            _fy_str = f"{fcf0/_mkt_cap_v*100:.2f}%"
            _ig_str = f"{_impl_g*100:.1f}%"
            _wk_str = f"{wacc_b*100:.1f}%"
            _cr_ref = f" vs. {_pct(rev_cagr_v)} trailing revenue CAGR" if rev_cagr_v else ""
            if _gap > 0.03:
                _vlbl, _vcol = "Ambitious Pricing", "var(--amber)"
                _vtxt = (f"Market embeds {_ig_str} implied FCF growth{_cr_ref}. "
                         f"Premium justified only if growth accelerates or margins expand materially.")
            elif _gap > -0.02:
                _vlbl, _vcol = "Fairly Priced", "var(--accent)"
                _vtxt = (f"Implied {_ig_str} FCF growth broadly in line{_cr_ref}. "
                         f"Market pricing continuity — upside requires a re-rating catalyst.")
            else:
                _vlbl, _vcol = "Conservative Pricing", "var(--up)"
                _vtxt = (f"Market prices in only {_ig_str} FCF growth{_cr_ref} — "
                         f"deceleration embedded. Upside if historical growth rate sustains.")
            rdcf_text = (
                '<div class="rdcf-stats">'
                f'<div class="rdcf-stat"><div class="rdcf-num">{_ig_str}</div><div class="rdcf-lbl">Implied FCF Growth</div><div class="rdcf-sub">Perpetuity rate at current price</div></div>'
                f'<div class="rdcf-stat"><div class="rdcf-num">{_fy_str}</div><div class="rdcf-lbl">FCF Yield</div><div class="rdcf-sub">Trailing FCF ÷ market cap</div></div>'
                f'<div class="rdcf-stat"><div class="rdcf-num">{_wk_str}</div><div class="rdcf-lbl">WACC</div><div class="rdcf-sub">Implied growth = WACC − yield</div></div>'
                '</div>'
                f'<div class="rdcf-verdict"><strong style="color:{_vcol}">{_vlbl}</strong> — {_vtxt}</div>'
            )
        else:
            _wk_str = f"{wacc_b*100:.1f}%"
            rdcf_text = (
                '<div class="rdcf-stats">'
                f'<div class="rdcf-stat"><div class="rdcf-num" style="color:var(--down)">N/A</div><div class="rdcf-lbl">Implied FCF Growth</div><div class="rdcf-sub">Requires positive FCF</div></div>'
                f'<div class="rdcf-stat"><div class="rdcf-num" style="color:var(--down)">{_b(fcf0)}</div><div class="rdcf-lbl">Trailing FCF</div><div class="rdcf-sub">Negative — model not applicable</div></div>'
                f'<div class="rdcf-stat"><div class="rdcf-num">{_wk_str}</div><div class="rdcf-lbl">WACC</div><div class="rdcf-sub">Discount rate</div></div>'
                '</div>'
                f'<div class="rdcf-verdict"><strong style="color:var(--amber)">FCF Inflection Play</strong> — '
                f'Trailing FCF is {_b(fcf0)}; reverse-DCF requires positive FCF. '
                f'Market pricing a future FCF inflection; trailing revenue CAGR: {_pct(rev_cagr_v)}.</div>'
            )
    else:
        rdcf_text = (
            '<div class="rdcf-stats">'
            '<div class="rdcf-stat"><div class="rdcf-num">—</div><div class="rdcf-lbl">Implied FCF Growth</div><div class="rdcf-sub">Insufficient data</div></div>'
            '<div class="rdcf-stat"><div class="rdcf-num">—</div><div class="rdcf-lbl">FCF Yield</div><div class="rdcf-sub">—</div></div>'
            '<div class="rdcf-stat"><div class="rdcf-num">—</div><div class="rdcf-lbl">WACC</div><div class="rdcf-sub">—</div></div>'
            '</div>'
            '<div class="rdcf-verdict">Insufficient data for reverse-DCF computation.</div>'
        )

    # ── Credit ────────────────────────────────────────────────────────────────
    from fmp_3statementv6 import MOODY_TO_SP
    sp_rating = manual_rating or "NR"
    moody_reverse = {v: k for k, v in MOODY_TO_SP.items()}
    moody_rating = moody_reverse.get(sp_rating, "NR") if sp_rating != "NR" else "NR"
    fitch_rating = "NR"
    cr_tier = _credit_tier(sp_rating)

    d_ebd_str   = _x(d_ebd)
    ebit_int_str = _x(ebit_int)
    credit_commentary = _credit_note(sp_rating, moody_rating, cr_tier,
                                      d_ebd_str, ebit_int_str, net_cash_str)

    # ── Scorecard tiers ───────────────────────────────────────────────────────
    t_rev    = _tier_rev_cagr(rev_cagr_v)
    t_roic   = _tier_roic(roic_v)
    t_fcf_ni = _tier_fcf_ni(fcf_ni_v)
    t_eint   = _tier_ebit_int(ebit_int)
    # Bank-aware leverage tier (report_bridge uses 3-tier: HIGH/MOD/LOW)
    if is_bank_v:
        ea = equity_assets_v
        t_debd = ("HIGH" if ea and ea > 0.10 else
                  "MOD"  if ea and ea > 0.06 else "LOW")
    else:
        t_debd = _tier_d_ebitda(d_ebd)
    t_pe     = _tier_pe(trailing_pe, pe_5yr)
    t_pfcf   = _tier_pfcf(trailing_pfc, pfcf_5yr)

    # Normalise manual qualitative tiers
    VALID_TIERS = {"HIGH", "MOD-HIGH", "MOD-LOW", "LOW"}
    t_bc  = (biz_clarity.upper() if biz_clarity and biz_clarity.upper() in VALID_TIERS else None)
    t_ltp = (ltp.upper() if ltp and ltp.upper() in VALID_TIERS else None)

    P = TIER_PTS
    p1 = round((P[t_bc or "MOD"]*2.5 + P["MOD"]*10.0 + P[t_ltp or "MOD"]*10.0 + P["MOD"]*7.5) / 10, 1)
    p2 = round((P[t_rev]*10.0 + P[t_fcf_ni]*10.0 + P["MOD"]*5.0 + P[t_roic]*7.5) / 10, 1)
    p3 = round((P[t_debd]*5.0 + P[t_eint]*7.5 + P["MOD"]*2.5) / 10, 1)
    p4 = round((P[t_pe]*10.0 + P[t_pfcf]*10.0) / 10, 1)
    # Use adj_score (Excel engine total + manual inputs) when available for accuracy
    final_score = adj_score or auto_score or round(p1 + p2 + p3 + p4, 1)

    # ── 5-year financial table ─────────────────────────────────────────────────
    fin = {}
    for i, (yr, is_, bs_, cf_) in enumerate(zip(years, is_data, bs_data, cf_data), 1):
        rev   = is_.get("revenue") or 0
        gp    = is_.get("grossProfit") or 0
        ebd   = _ebitda(is_, cf_)
        ebit  = is_.get("operatingIncome") or 0
        ni    = is_.get("netIncome") or 0
        cash  = bs_.get("cashAndCashEquivalents") or 0
        std   = bs_.get("shortTermDebt") or 0
        ltd   = bs_.get("longTermDebt") or 0
        debt  = std + ltd
        nc    = cash - debt
        eq    = bs_.get("totalStockholdersEquity") or 0
        intg  = bs_.get("goodwillAndIntangibleAssets") or bs_.get("intangibleAssets") or 0
        ocf   = cf_.get("operatingCashFlow") or 0
        cpx   = abs(cf_.get("capitalExpenditure") or 0)
        fcf   = cf_.get("freeCashFlow") or (ocf - cpx)
        ie    = abs(is_.get("interestExpense") or 0)

        fin[f"FY{i}"]         = f"FY{yr}"
        fin[f"REV_FY{i}"]     = _m(rev)
        fin[f"GP_FY{i}"]      = _m(gp)
        fin[f"GM_FY{i}"]      = _pct(gp/rev if rev else None)
        fin[f"EBITDA_FY{i}"]  = _m(ebd)
        fin[f"EBITDAM_FY{i}"] = _pct(ebd/rev if rev else None)
        fin[f"EBIT_FY{i}"]    = _m(ebit)
        fin[f"NI_FY{i}"]      = _m(ni)
        fin[f"CASH_FY{i}"]    = _m(cash)
        fin[f"DEBT_FY{i}"]    = _m(debt)
        _nc = cash - debt
        fin[f"NETCASH_FY{i}"] = _m(_nc) if _nc >= 0 else f"({_m(abs(_nc))})"
        fin[f"EQUITY_FY{i}"]  = _m(eq)
        fin[f"INTANG_FY{i}"]  = _m(intg)
        _de = debt / ebd if ebd > 0 else None
        fin[f"DEBITDA_FY{i}"] = _x(_de)
        _ei = abs(ebit) / ie if ie > 0 else None
        fin[f"EBITINT_FY{i}"] = _x(_ei)
        fin[f"OCF_FY{i}"]     = _m(ocf)
        fin[f"CAPEX_FY{i}"]   = f"({_m(cpx)})"
        fin[f"FCF_FY{i}"]     = _m(fcf)

        # New: net margin, D&A, FCF bridge
        fin[f"NM_FY{i}"]      = _pct(ni/rev if rev else None)
        da = max(0, ebd - ebit)   # D&A ≈ EBITDA − EBIT
        fin[f"DA_FY{i}"]      = _m(da)
        # Approximate effective tax rate from NI / pre-tax income
        pti = is_.get("incomeBeforeTax") or is_.get("pretaxIncome") or 0
        te  = abs(is_.get("incomeTaxExpense") or 0)
        eff_tax = te / pti if pti > 0 else 0.21  # fallback 21%
        nopat = ebit * (1 - eff_tax)
        ufcf  = nopat + da - cpx          # simplified (no ΔNWC from API)
        fin[f"NOPAT_FY{i}"]   = _m(nopat)
        fin[f"UFCF_FY{i}"]    = _m(ufcf)
        fin[f"UFCFM_FY{i}"]   = _pct(ufcf/rev if rev else None)
        fin[f"LFCFM_FY{i}"]   = _pct(fcf/rev if rev else None)
        fin[f"FCF_CONV_FY{i}"] = _x(fcf/ni if ni else None)

    # ── Chart arrays ──────────────────────────────────────────────────────────
    yr_labels  = [f"FY{y}" for y in years]
    rev_b_lst  = [(is_.get("revenue") or 0)/1e9 for is_ in is_data]
    ebd_b_lst  = [_ebitda(is_, cf_)/1e9 for is_, cf_ in zip(is_data, cf_data)]
    ni_b_lst   = [(is_.get("netIncome") or 0)/1e9 for is_ in is_data]
    ocf_b_lst  = [(cf_.get("operatingCashFlow") or 0)/1e9 for cf_ in cf_data]
    fcf_b_lst  = [(cf_.get("freeCashFlow") or 0)/1e9 for cf_ in cf_data]
    roic_lst   = [_roic(is_, bs_) for is_, bs_ in zip(is_data, bs_data)]
    roic_pct   = [round((r or 0)*100, 1) for r in roic_lst]

    # Shareholder returns: real data from CF statement
    buyback_b_lst = [abs(cf_.get("commonStockRepurchased") or
                         cf_.get("repurchaseOfCommonStock") or 0) / 1e9
                     for cf_ in cf_data]
    fcf_ps_lst    = [round(
        (cf_.get("freeCashFlow") or
         (cf_.get("operatingCashFlow") or 0) - abs(cf_.get("capitalExpenditure") or 0))
        / shares, 2) if shares > 0 else 0
        for cf_ in cf_data]

    # ── Assemble DATA dict ─────────────────────────────────────────────────────
    D = {
        # Header
        "COMPANY_NAME":     company_name,
        "TICKER_LINE":      ticker_line,
        "DESCRIPTION_LINE": industry,
        "VERDICT_TEXT":     _verdict(final_score),
        "CURRENT_PRICE":    current_price,
        "PRICE_TARGET":     price_target,
        "REPORT_DATE":      today,
        "FINAL_SCORE":      round(final_score, 1) if final_score else 0,
        "FINAL_SCORE_CALC": str(round(final_score, 1)) if final_score else "0",
        "N_ACTUALS":        str(len(is_data)),

        # Overview
        "COMPANY_SUMMARY_TEXT": (
            (description[:700] + "...") if len(description) > 700 else description
        ) or f"{company_name} operates in the {industry} sector. {ceo_info}",

        # Market data
        "MARKET_CAP":       _b(market_cap),
        "SHARES_DILUTED":   shares_str,
        "ENTERPRISE_VALUE": _b(ev),
        "NET_CASH_DEBT":    net_cash_str,

        # Valuation multiples
        "TRAILING_PE":          _x(trailing_pe),
        "TRAILING_PE_10YR":     (_x(pe_5yr) + " (5yr avg)") if pe_5yr else "N/A",
        "TRAILING_PE_DELTA":    pe_delta,
        "FORWARD_PE":           "N/A",
        "FORWARD_PE_EST":       "Add analyst EPS estimate",
        "FORWARD_PE_10YR":      _x(pe_5yr),
        "FORWARD_PE_DELTA":     pe_delta,
        "TRAILING_PFCF":        _x(trailing_pfc),
        "TRAILING_PFCF_10YR":   (_x(pfcf_5yr) + " (5yr avg)") if pfcf_5yr else "N/A",
        "TRAILING_PFCF_DELTA":  pfcf_delta,
        "FORWARD_PFCF":         "N/A",
        "FORWARD_PFCF_EST":     "Add analyst FCF estimate",
        "FORWARD_PFCF_10YR":    _x(pfcf_5yr),
        "FORWARD_PFCF_DELTA":   pfcf_delta,
        "EV_EBITDA_TRAILING":   _x(ev_ebitda),
        "EV_EBITDA_FWD_NOTE":   "Add fwd estimate",

        # Key metrics
        "REVENUE_YEAR_LABEL":   f"FY{years[-1]}",
        "REVENUE_VALUE":        _b(rev0),
        "REVENUE_SUB":          (f"+{rev_yoy*100:.0f}% YoY" if rev_yoy and rev_yoy >= 0
                                  else f"{rev_yoy*100:.0f}% YoY" if rev_yoy else "N/A"),
        "EBITDA_MARGIN_LABEL":  "EBITDA Margin",
        "EBITDA_MARGIN":        _pct(ebitdam0),
        "EBITDA_MARGIN_SUB":    f"FY{years[-1]}",
        "FCF_VALUE":            _b(fcf0),
        "FCF_SUB":              (f"+{fcf_yoy*100:.0f}% YoY" if fcf_yoy and fcf_yoy >= 0
                                  else f"{fcf_yoy*100:.0f}% YoY" if fcf_yoy else "N/A"),
        "ROIC_VALUE":           _pct(roic_v),
        "ROIC_SUB":             f"FY{years[-1]}",
        "CAP_RETURNS_LABEL":    "Capital Returns",
        "CAP_RETURNS_VALUE":    "See Excel",
        "CAP_RETURNS_SUB":      "Dividends + Buybacks — see CF tab",

        # Revenue mix (stubs — fill from Excel Segments tab)
        "REV_MIX_SECTION_LABEL": "Revenue Mix",
        "SEG1_EMOJI_NAME": "📊 Segment 1", "SEG1_REV_PCT": "—",
        "SEG1_DESC": "Open the Excel model → Segments tab for breakdown.",
        "SEG2_EMOJI_NAME": "📊 Segment 2", "SEG2_REV_PCT": "—",
        "SEG2_DESC": "Open the Excel model → Segments tab for breakdown.",
        "SEG3_EMOJI_NAME": "📊 Segment 3", "SEG3_REV_PCT": "—",
        "SEG3_DESC": "Open the Excel model → Segments tab for breakdown.",

        # Credit
        "SP_RATING":       sp_rating,    "SP_OUTLOOK":       "See 10-K",
        "SP_TIER_LABEL":   f"Tier: {cr_tier}",
        "MOODYS_RATING":   moody_rating, "MOODYS_OUTLOOK":   "See 10-K",
        "MOODYS_TIER_LABEL": f"Tier: {cr_tier}",
        "FITCH_RATING":    fitch_rating, "FITCH_OUTLOOK":    "NR",
        "FITCH_TIER_LABEL": "Tier: NR",
        "CREDIT_NOTE_TEXT": credit_commentary,

        # Thesis (smart stubs with real numbers embedded)
        "THESIS_MOAT_TEXT": (
            f"<strong>Auto-generated draft — add qualitative moat analysis.</strong> "
            f"{company_name} ({ticker}) operates in {industry} with FY{years[-1]} ROIC of {_pct(roic_v)}, "
            f"gross margin of {_pct(gm0)}, and revenue CAGR of {_pct(rev_cagr_v)} over three years. "
            f"Review the 10-K competitive dynamics section and update this field."
        ),
        "THESIS_CATALYSTS_TEXT": (
            "Add 2–3 specific catalysts. Review company guidance, product pipeline, M&A activity, "
            "and market expansion opportunities in the 10-K Management Discussion section."
        ),
        "THESIS_VALUATION_TEXT": (
            f"At ${current_price:.2f}, {ticker} trades at {_x(trailing_pe)} trailing P/E "
            f"vs. {_x(pe_5yr)} 5-year average ({pe_delta}). "
            f"DCF (Gordon Growth model): base ${base_px:.0f} ({_vs(base_px, current_price)}). "
            f"<strong>Add qualitative valuation commentary after reviewing the Excel DCF model.</strong>"
            if base_px and current_price else
            "Review the Excel DCF model and add valuation commentary here."
        ),

        # Financials
        "CURRENCY_NAME":   "USD",
        "CURRENCY_SYMBOL": "$",
        "FIN_TABLE_NOTE":  f"FY{years[-1]} data from {company_name} annual report via FMP API; figures in USD millions.",

        # Growth context
        "REV_CAGR_TRAIL": f"Trailing 3yr: {_pct(rev_cagr_v)}",
        "REV_CAGR_FWD":   "Add consensus forward estimate",
        "REV_CAGR_NOTE":  "Source: FMP API.",
        "FCF_CAGR_TRAIL": "Add manually",
        "FCF_CAGR_FWD":   "Add consensus estimate",
        "FCF_CAGR_NOTE":  "",
        "PE_TRAIL_STAT":  _x(trailing_pe),
        "PE_FWD_STAT":    "N/A",
        "PE_AVG_NOTE":    f"5-yr avg: {_x(pe_5yr)}",
        "PFCF_TRAIL_STAT": _x(trailing_pfc),
        "PFCF_FWD_STAT":  "N/A",
        "PFCF_AVG_NOTE":  f"5-yr avg: {_x(pfcf_5yr)}",

        # Chart data
        "FIN_LABELS_JS":   _js_str_arr(yr_labels),
        "REV_JS":          _js_arr(rev_b_lst, 1),
        "EBITDA_JS":       _js_arr(ebd_b_lst, 1),
        "NI_JS":           _js_arr(ni_b_lst, 1),
        "OCF_JS":          _js_arr(ocf_b_lst, 1),
        "FCF_JS":          _js_arr(fcf_b_lst, 1),
        "ROIC_LABELS_JS":  _js_str_arr(yr_labels),
        "ROIC_DATA_JS":    _js_arr(roic_pct, 1),
        "WACC_VALUE":      f"{(wacc_val or 0)*100:.2f}",
        "PRICE_LABELS_JS": _js_str_arr(yr_labels),
        "PRICE_DATA_JS":   "[0]",
        "RET_LABELS_JS":   _js_str_arr(yr_labels),
        "BUYBACK_JS":      _js_arr(buyback_b_lst, 2),
        "FCF_PER_SHARE_JS": _js_arr(fcf_ps_lst, 2),

        # Scorecard — totals use adj_score from Excel engine for consistency
        "P1_WEIGHTED":          str(p1),
        "P1_BC_SCORE_TEXT":     t_bc or "MOD", "P1_BC_WTD": str(round(P[t_bc or "MOD"]*2.5/10, 2)),
        "P1_BC_COMMENTARY":     (f"Manual input: {t_bc}." if t_bc else "Review business segment clarity from 10-K."),
        "P1_MOAT_SCORE_TEXT":   "MOD", "P1_MOAT_WTD": "7.0",
        "P1_MOAT_COMMENTARY":   f"Auto-proxy: {_pct(gm0)} gross margin, {_pct(rev_cagr_v)} rev CAGR. Confirm with 10-K moat analysis.",
        "P1_LTP_SCORE_TEXT":    t_ltp or "MOD", "P1_LTP_WTD": str(round(P[t_ltp or "MOD"]*10.0/10, 1)),
        "P1_LTP_COMMENTARY":    (f"Manual input: {t_ltp}." if t_ltp else "Review long-term positioning and TAM manually."),
        "P1_MGT_SCORE_TEXT":    "MOD", "P1_MGT_WTD": "5.25",
        "P1_MGT_COMMENTARY":    ceo_info + " Review track record and capital allocation.",

        "P2_WEIGHTED":          str(p2),
        "P2_RC_SCORE_TEXT":     t_rev,  "P2_RC_WTD": str(round(P[t_rev]*10.0/10, 1)),
        "P2_RC_COMMENTARY":     f"3yr revenue CAGR: {_pct(rev_cagr_v)}.",
        "P2_CQ_SCORE_TEXT":     t_fcf_ni, "P2_CQ_WTD": str(round(P[t_fcf_ni]*10.0/10, 1)),
        "P2_CQ_COMMENTARY":     f"FCF/NI: {_pct(fcf_ni_v)}.",
        "P2_CR_SCORE_TEXT":     "MOD",  "P2_CR_WTD": "3.5",
        "P2_CR_COMMENTARY":     "Review buyback + dividend history in Excel CF tab.",
        "P2_ROIC_SCORE_TEXT":   t_roic, "P2_ROIC_WTD": str(round(P[t_roic]*7.5/10, 1)),
        "P2_ROIC_COMMENTARY":   f"Latest ROIC: {_pct(roic_v)}.",

        "P3_WEIGHTED":          str(p3),
        "P3_CREDRISK_SCORE_TEXT": t_debd, "P3_CREDRISK_WTD": str(round(P[t_debd]*5.0/10, 1)),
        "P3_CREDRISK_COMMENTARY": (
            f"Capital Adequacy (Equity/Assets): {_pct(equity_assets_v)} — CET1 proxy."
            if is_bank_v else f"D/EBITDA: {d_ebd_str}."
        ),
        "P3_IC_SCORE_TEXT":     t_eint, "P3_IC_WTD": str(round(P[t_eint]*7.5/10, 1)),
        "P3_IC_COMMENTARY":     f"EBIT/Interest: {ebit_int_str}.",
        "P3_ER_SCORE_TEXT":     "MOD",  "P3_ER_WTD": "1.75",
        "P3_ER_COMMENTARY":     "Review execution risks from 10-K Risk Factors section.",

        "P4_WEIGHTED":          str(p4),
        "P4_PE_SCORE_TEXT":     t_pe,   "P4_PE_WTD": str(round(P[t_pe]*10.0/10, 1)),
        "P4_PE_COMMENTARY":     f"P/E {_x(trailing_pe)} vs 5yr avg {_x(pe_5yr)} ({pe_delta}).",
        "P4_PFCF_SCORE_TEXT":   t_pfcf, "P4_PFCF_WTD": str(round(P[t_pfcf]*10.0/10, 1)),
        "P4_PFCF_COMMENTARY":   f"P/FCF {_x(trailing_pfc)} vs 5yr avg {_x(pfcf_5yr)} ({pfcf_delta}).",

        "TOTAL_WEIGHTED_SCORE": str(round(final_score, 1)) if final_score else "0",

        # Scenarios
        "BEAR_PRICE_RANGE": f"${bear_px:.0f}" if bear_px else "N/A",
        "BASE_PRICE_RANGE": f"${base_px:.0f}" if base_px else "N/A",
        "BULL_PRICE_RANGE": f"${bull_px:.0f}" if bull_px else "N/A",
        "BEAR_REVENUE_GROWTH": "Add manually",
        "BASE_REVENUE_GROWTH": f"{_pct(rev_cagr_v)} CAGR (trailing)",
        "BULL_REVENUE_GROWTH": "Add manually",
        "BEAR_MARGIN":   "Add manually",
        "BASE_MARGIN":   f"{_pct(ebitdam0)} EBITDA",
        "BULL_MARGIN":   "Add manually",
        "FCF_YEAR_LABEL": f"FY{int(years[-1])+3}E" if years else "Fwd",
        "BEAR_FCF": "Add", "BASE_FCF": "Add", "BULL_FCF": "Add",
        "BEAR_MULTIPLE": "Add", "BASE_MULTIPLE": "Add", "BULL_MULTIPLE": "Add",
        "BEAR_CATALYST": "Execution miss, margin compression, macro headwinds",
        "BASE_CATALYST": "Consensus estimates met, stable macro",
        "BULL_CATALYST": "Revenue acceleration, margin expansion, multiple re-rating",
        "BEAR_UPSIDE": _vs(bear_px, current_price),
        "BASE_UPSIDE": _vs(base_px, current_price),
        "BULL_UPSIDE": _vs(bull_px, current_price),

        # Risks (sector-generic — add specifics from 10-K)
        "RISK_1_TITLE": "Competitive Risk",
        "RISK_1_TEXT":  "Add specific competitive risks from 10-K Risk Factors section.",
        "RISK_2_TITLE": "Execution Risk",
        "RISK_2_TEXT":  "Add execution risks — margin, capex intensity, integration.",
        "RISK_3_TITLE": "Regulatory / Macro Risk",
        "RISK_3_TEXT":  "Add regulatory and macroeconomic sensitivity risks.",
        "RISK_4_TITLE": "Capital Allocation Risk",
        "RISK_4_TEXT":  (
            f"Capital Adequacy (Equity/Assets): {_pct(equity_assets_v)}. Monitor regulatory capital ratios."
            if is_bank_v else f"D/EBITDA: {d_ebd_str}. Review capital allocation discipline."
        ),
        "RISK_5_TITLE": "Valuation Risk",
        "RISK_5_TEXT":  f"At {_x(trailing_pe)} trailing P/E, material multiple compression possible in risk-off environments.",

        # Valuation verdict
        "VALUATION_VERDICT_TITLE": "Valuation Analysis — Auto-Generated Draft (Review Excel Model)",
        "VALUATION_VERDICT_TEXT": (
            f"At ${current_price:.2f}, {company_name} ({ticker}) trades at {_x(trailing_pe)} trailing P/E "
            f"vs. {_x(pe_5yr)} 5-year average ({pe_delta}), and {_x(trailing_pfc)} P/FCF "
            f"vs. {_x(pfcf_5yr)} 5-year average ({pfcf_delta}). "
            f"DCF base case (Gordon Growth, WACC {wacc_b*100:.1f}%): <strong>${base_px:.0f} ({_vs(base_px, current_price)})</strong>. "
            f"Bear: ${bear_px:.0f} ({_vs(bear_px, current_price)}) | Bull: ${bull_px:.0f} ({_vs(bull_px, current_price)}). "
            f"<strong>Note: Add qualitative valuation commentary and forward estimates after reviewing the Excel model.</strong>"
        ) if base_px and current_price else "Review DCF in Excel model and add valuation commentary.",

        "DCF_BEAR_WACC": f"{w_bear*100:.1f}%", "DCF_BEAR_TGR": f"{tgr_bear*100:.1f}%",
        "DCF_BEAR_CAGR": "Add", "DCF_BEAR_PX": f"${bear_px:.0f}" if bear_px else "N/A",
        "DCF_BEAR_VS":   _vs(bear_px, current_price),
        "DCF_BASE_WACC": f"{w_base*100:.1f}%", "DCF_BASE_TGR": f"{tgr_base*100:.1f}%",
        "DCF_BASE_CAGR": _pct(rev_cagr_v) + " (trailing)",
        "DCF_BASE_PX":   f"${base_px:.0f}" if base_px else "N/A",
        "DCF_BASE_VS":   _vs(base_px, current_price),
        "DCF_BULL_WACC": f"{w_bull*100:.1f}%", "DCF_BULL_TGR": f"{tgr_bull*100:.1f}%",
        "DCF_BULL_CAGR": "Add", "DCF_BULL_PX": f"${bull_px:.0f}" if bull_px else "N/A",
        "DCF_BULL_VS":   _vs(bull_px, current_price),

        # Sensitivity headers
        "TGR_1": "2.0%","TGR_2": "2.5%","TGR_3": "3.0%","TGR_4": "3.5%","TGR_5": "4.0%",
        "WACC_1": f"{(wacc_b-0.015)*100:.1f}%", "WACC_2": f"{(wacc_b-0.01)*100:.1f}%",
        "WACC_3": f"{wacc_b*100:.1f}%",          "WACC_4": f"{(wacc_b+0.01)*100:.1f}%",
        "WACC_5": f"{(wacc_b+0.015)*100:.1f}%",  "WACC_6": f"{(wacc_b+0.02)*100:.1f}%",
        "SENSITIVITY_NOTE": "Red = Below Current Price | Yellow = 0–10% upside | Green = >10% upside",
        "REVERSE_DCF_TEXT": rdcf_text,

        # WACC Summary
        "WACC_RF":        f"{RF_APPROX*100:.1f}%",
        "WACC_BETA":      f"{beta_v:.2f}",
        "WACC_BETA_SOURCE": "FMP 5yr vs S&P500",
        "WACC_ERP":       f"{ERP_APPROX*100:.1f}%",
        "WACC_KE":        f"{ke_approx*100:.1f}%",
        "WACC_KD_PRETAX": f"{max(0,kd_pre)*100:.1f}%",
        "WACC_TAX_RATE":  f"{eff_tax0*100:.1f}%",
        "WACC_EW":        f"{ew_v*100:.0f}%",
        "WACC_DW":        f"{dw_v*100:.0f}%",
        "WACC_NOTE":      (f"WACC = {ew_v*100:.0f}% × {ke_approx*100:.1f}% (Ke) + {dw_v*100:.0f}% × {max(0,kd_pre)*100:.1f}% × (1 − {eff_tax0*100:.0f}% tax) ≈ {wacc_b*100:.1f}%. "
                           f"Beta: FMP 5yr monthly. ERP: Damodaran avg implied/historical. Rf: FRED DGS10. Review Excel WACC tab for full derivation."),

        # EV Bridge
        "DCF_PV_FCFS":   _b(pv_fcfs_approx) if pv_fcfs_approx else "—",
        "DCF_PV_TV":      _b(pv_tv_approx) if pv_tv_approx else "—",
        "DCF_TV_PCT":     "~65% (est.)" if pv_tv_approx else "—",
        "DCF_TV_METHOD":  "Gordon Growth",
        "DCF_EV":         _b(ev_approx) if ev_approx else "—",
        "DCF_NET_DEBT_B": (_b(abs(net_cash_val)) if net_cash_val < 0 else f"+{_b(net_cash_val)}"),
        "DCF_EQUITY_VAL": _b(eq_val_approx) if eq_val_approx else "—",
        "DCF_SHARES_B":   f"{shares_v/1e6:.1f}mm" if shares_v > 1e6 else str(int(shares_v)),

        # Multiples
        "FY_FWD1": f"FY{int(years[-1])+1}E" if years else "FY+1E",
        "FY_FWD2": f"FY{int(years[-1])+2}E" if years else "FY+2E",
        "PE_10YR_AVG":  (_x(pe_5yr) + " (5yr)") if pe_5yr else "N/A",
        "PE_5YR_AVG":   _x(pe_5yr),
        "PFCF_10YR_AVG": (_x(pfcf_5yr) + " (5yr)") if pfcf_5yr else "N/A",
        "PFCF_5YR_AVG": _x(pfcf_5yr),
        "EVEBITDA_10YR_AVG": _x(ev_ebitda) if ev_ebitda else "Add",
        "EVEBITDA_5YR_AVG":  _x(ev_ebitda) if ev_ebitda else "Add",
        **{k: "Add" for k in [
            "MULT_PE10_FY1_PX","MULT_PE10_FY1_UPS","MULT_PE10_FY2_PX","MULT_PE10_FY2_UPS",
            "MULT_PE5_FY1_PX","MULT_PE5_FY1_UPS","MULT_PE5_FY2_PX","MULT_PE5_FY2_UPS",
            "MULT_PFCF10_FY1_PX","MULT_PFCF10_FY1_UPS","MULT_PFCF10_FY2_PX","MULT_PFCF10_FY2_UPS",
            "MULT_PFCF5_FY1_PX","MULT_PFCF5_FY1_UPS","MULT_PFCF5_FY2_PX","MULT_PFCF5_FY2_UPS",
            "MULT_EV10_FY1_PX","MULT_EV10_FY1_UPS","MULT_EV10_FY2_PX","MULT_EV10_FY2_UPS",
            "MULT_EV5_FY1_PX","MULT_EV5_FY1_UPS","MULT_EV5_FY2_PX","MULT_EV5_FY2_UPS",
        ]},
        "MULTIPLES_METHOD_RATIONALE": "Method Rationale: Review multiples vs. historical averages after adding forward estimates.",
        "MULTIPLES_KEY_QUESTION":     f"Key Question: Is the {pe_delta} discount/premium to 5yr P/E justified by current growth?",
        "COMPOSITE_FAIR_VALUE":       f"${base_px:.0f}" if base_px else "N/A",
        "COMPOSITE_UPSIDE_NOTE":      f"DCF base: ${base_px:.0f} ({_vs(base_px, current_price)}). Add multiples-based range." if base_px else "See Excel model.",

        # Analysts (stubs)
        "TICKER_SHORT":   ticker,
        "BUY_PCT":        "N/A", "BUY_COUNT":  "Add manually",
        "HOLD_PCT":       "N/A", "HOLD_COUNT": "Add manually",
        "SELL_PCT":       "N/A", "SELL_COUNT": "Add manually",
        **{f"A{i}_{k}": "—" for i in range(1,8)
           for k in ["NAME","FIRM","PT","PT_VS","DATE"]},
        **{f"A{i}_RATING_TEXT": "N/A" for i in range(1,8)},
        "ANALYST_COUNT":      "N/A",
        "CONSENSUS_PT":       "Add manually",
        "CONSENSUS_PT_VS":    "Add manually",
        "PT_RANGE":           "Add manually",
        "ANALYST_TABLE_NOTE": "Add analyst consensus data from Bloomberg / FactSet / sell-side reports.",

        # Footnotes
        "FN1": f"Credit ratings: S&P {sp_rating} — manual input or Damodaran ICR model.",
        "FN2": f"Historical financials via FMP API. Fiscal year {years[-1]}.",
        "FN3": f"Revenue 3yr CAGR: {_pct(rev_cagr_v)} from FY{years[-4] if len(years)>=4 else years[0]}–{years[-1]}.",
        "FN4": "Share price history: add historical series manually.",
        "FN5": f"ROIC: {_pct(roic_v)} — NOPAT / Invested Capital from FMP data.",
        "FN6": "Valuation multiples: 5-year averages from FMP ratios API.",
        "FN7": "Risk factors: add from 10-K Risk Factors section.",
        "FN8": f"DCF: WACC {wacc_b*100:.1f}% (Damodaran-based), TGR 3.0%.",
        "DISCLAIMER_TEXT": "This report is auto-generated from public financial data. For informational purposes only. Not investment advice.",
    }

    D.update(fin)

    # EBIT/Interest note (template uses a single colspan=5 cell for this row)
    ebitint_vals = [fin.get(f"EBITINT_FY{i}", "N/A") for i in range(1, len(years) + 1)]
    non_na = [(years[i], v) for i, v in enumerate(ebitint_vals) if v != "N/A"]
    D["EBITINT_NOTE"] = " · ".join(f"FY{yr}: {v}" for yr, v in non_na) if non_na else "N/A — interest data not available"

    # ── Analyst estimates → forward multiples ─────────────────────────────────
    if analyst_ests:
        _ests = sorted(analyst_ests, key=lambda x: x.get("date", ""))
        _e1 = _ests[0] if _ests else {}
        _e2 = _ests[1] if len(_ests) >= 2 else {}

        _eps1  = _e1.get("estimatedEpsAvg") or 0
        _eps2  = _e2.get("estimatedEpsAvg") or 0
        _ebd1  = _e1.get("estimatedEbitdaAvg") or 0
        _ebd2  = _e2.get("estimatedEbitdaAvg") or 0

        # FCF/share proxy: forward EPS × historical FCF/NI conversion ratio
        _fn = max(0.3, min(fcf_ni_v or 0.7, 2.0))
        _fp1 = round(_eps1 * _fn, 2) if _eps1 > 0 else 0
        _fp2 = round(_eps2 * _fn, 2) if _eps2 > 0 else 0

        _pa  = pe_5yr          # 5yr P/E avg (used for both 5yr and 10yr labels)
        _pfa = pfcf_5yr        # 5yr P/FCF avg
        _eva = ev_ebitda       # trailing EV/EBITDA (used as proxy historical avg)
        _nd  = debt0 - cash0   # net debt (positive = debt > cash)
        _sh  = shares or 1

        def _mpx(mult, pershare):
            return round(mult * pershare, 0) if mult and pershare and pershare > 0 else None

        def _mpx_ev(mult, ebd):
            if mult and ebd and ebd > 0 and _sh > 0:
                eq_impl = mult * ebd - _nd
                return round(eq_impl / _sh, 0) if eq_impl > 0 else None
            return None

        def _fp(v): return f"${v:.0f}" if v else "N/A"
        def _fu(v): return _vs(v, current_price) if v else "N/A"

        mu = {}
        for _fy, _e, _fp_, _ebd in [("FY1", _eps1, _fp1, _ebd1), ("FY2", _eps2, _fp2, _ebd2)]:
            _pe_px   = _mpx(_pa,  _e);   mu[f"MULT_PE10_{_fy}_PX"] = _fp(_pe_px);   mu[f"MULT_PE10_{_fy}_UPS"] = _fu(_pe_px)
            mu[f"MULT_PE5_{_fy}_PX"]    = _fp(_pe_px);   mu[f"MULT_PE5_{_fy}_UPS"]    = _fu(_pe_px)
            _pf_px   = _mpx(_pfa, _fp_); mu[f"MULT_PFCF10_{_fy}_PX"] = _fp(_pf_px); mu[f"MULT_PFCF10_{_fy}_UPS"] = _fu(_pf_px)
            mu[f"MULT_PFCF5_{_fy}_PX"]  = _fp(_pf_px);  mu[f"MULT_PFCF5_{_fy}_UPS"]  = _fu(_pf_px)
            _ev_px   = _mpx_ev(_eva, _ebd); mu[f"MULT_EV10_{_fy}_PX"] = _fp(_ev_px); mu[f"MULT_EV10_{_fy}_UPS"] = _fu(_ev_px)
            mu[f"MULT_EV5_{_fy}_PX"]    = _fp(_ev_px);  mu[f"MULT_EV5_{_fy}_UPS"]    = _fu(_ev_px)

        if _eva:
            mu["EVEBITDA_5YR_AVG"]  = f"{_eva:.1f}x (trailing)"
            mu["EVEBITDA_10YR_AVG"] = f"{_eva:.1f}x (trailing)"

        # Composite: average all non-N/A implied prices + DCF base
        _all_implied = [float(v[1:]) for v in mu.values()
                        if isinstance(v, str) and v.startswith("$")]
        if _all_implied:
            _comp_mult = round(sum(_all_implied) / len(_all_implied), 0)
            if base_px and current_price:
                _comp_all = round((sum(_all_implied) + base_px) / (len(_all_implied) + 1), 0)
                mu["COMPOSITE_FAIR_VALUE"]  = f"${_comp_all:.0f}"
                mu["COMPOSITE_UPSIDE_NOTE"] = (
                    f"DCF ${base_px:.0f} ({_vs(base_px, current_price)}) + "
                    f"multiples avg ${_comp_mult:.0f} ({_vs(_comp_mult, current_price)}) "
                    f"→ composite ${_comp_all:.0f} ({_vs(_comp_all, current_price)})."
                )
            else:
                mu["COMPOSITE_FAIR_VALUE"]  = f"${_comp_mult:.0f}"
                mu["COMPOSITE_UPSIDE_NOTE"] = (
                    f"Avg of {len(_all_implied)} multiples-based targets: "
                    f"${_comp_mult:.0f} ({_vs(_comp_mult, current_price)})."
                )

        # Write — override "Add" stubs; preserve any real values already set
        for k, v in mu.items():
            if D.get(k) in ("Add", "N/A", None, ""):
                D[k] = v
        # Always update composite (it starts as DCF-only; now includes multiples)
        for k in ("COMPOSITE_FAIR_VALUE", "COMPOSITE_UPSIDE_NOTE"):
            if k in mu:
                D[k] = mu[k]

    return D


# ══════════════════════════════════════════════════════════════════════════════
# RENDER HTML
# ══════════════════════════════════════════════════════════════════════════════

def render_html_report(data):
    """Fill Report_Template.html with data dict, return HTML string."""

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    current_price = data.get("CURRENT_PRICE", 0) or 0

    # 1. Colorize SCORE_TEXT placeholders before main replacement
    for k in [k for k in data if "SCORE_TEXT" in k]:
        html = html.replace(f"{{{{{k}}}}}", _score_class(str(data[k])))

    # 2. Hide unused CEO placeholders
    html = html.replace("{{CEO_NAME}}", "").replace("{{CEO_TENURE}}", "")

    # 3. CSS classes
    for k, v in _compute_css(data, current_price).items():
        html = html.replace(f"{{{{{k}}}}}", str(v))

    # 4. Main data replacement
    for k, v in data.items():
        if isinstance(v, (str, int, float, bool)):
            html = html.replace(f"{{{{{k}}}}}", str(v))

    # 5. Sensitivity grid (6x5)
    _bpx_raw = str(data.get("DCF_BASE_PX", "0")).lstrip("$").replace(",", "")
    try:
        base_px = float(_bpx_raw) if _bpx_raw not in ("N/A", "", "—", "-") else 0.0
    except ValueError:
        base_px = 0.0
    base_wacc = float(str(data.get("DCF_BASE_WACC", "9.0%")).rstrip("%")) / 100
    base_tgr  = float(str(data.get("DCF_BASE_TGR",  "3.0%")).rstrip("%")) / 100
    base_spread = (base_wacc - base_tgr) if (base_wacc - base_tgr) > 0 else 0.07

    waccs = [float(str(data.get(f"WACC_{i}", "9.0%")).rstrip("%")) / 100 for i in range(1, 7)]
    tgrs  = [float(str(data.get(f"TGR_{j}", "3.0%")).rstrip("%")) / 100  for j in range(1, 6)]

    for i, w in enumerate(waccs):
        for j, t in enumerate(tgrs):
            spread = w - t
            if spread > 0 and base_px:
                implied = round(base_px * base_spread / spread, 0)
            else:
                implied = base_px or 0
            html = html.replace(f"{{{{S{i+1}{j+1}}}}}", _sensitivity_class(current_price, implied))
            html = html.replace(f"{{{{V{i+1}{j+1}}}}}", f"${implied:.0f}")

    return html
