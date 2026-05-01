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

from data_validation import validate_fmp_data, persist_anomalies

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "Report_Template.html")
DATA_DIR = os.path.join(os.path.dirname(__file__), "static", "data")

# ── Gemini AI key (injected from server.py at startup) ────────────────────────
GEMINI_KEY = os.environ.get("GEMINI_KEY", "")

def _gemini(prompt, timeout=20):
    """Call Gemini 1.5 Flash for qualitative commentary. Returns None on failure."""
    if not GEMINI_KEY:
        return None
    try:
        import requests as _r
        url = (f"https://generativelanguage.googleapis.com/v1beta/"
               f"models/gemini-1.5-flash:generateContent?key={GEMINI_KEY}")
        resp = _r.post(url, json={
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.25, "maxOutputTokens": 300},
        }, timeout=timeout)
        if resp.status_code == 200:
            return resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
    except Exception:
        pass
    return None

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

# ── Thesis builder ────────────────────────────────────────────────────────────

def _build_thesis(ticker, metrics, years, is_data, cf_data):
    """
    Build a 3-sentence investment thesis from computed metrics.
    Returns a dict with keys: moat, valuation, risk
    Each value is one sentence (~25-40 words).
    """
    roic = metrics.get("roic")
    rev_cagr = metrics.get("rev_cagr")
    fcf_ni = metrics.get("fcf_ni")
    pe_current = metrics.get("pe_current")
    pe_5yr = metrics.get("pe_5yr")
    dcf_base_px = metrics.get("dcf_base_px")
    current_price = metrics.get("current_price")
    d_ebitda = metrics.get("d_ebitda")
    ebit_interest = metrics.get("ebit_interest")
    gm_trend = metrics.get("gm_trend", 0)
    fcf_margin = metrics.get("fcf_margin")
    pfcf_current = metrics.get("pfcf_current")
    pfcf_5yr = metrics.get("pfcf_5yr")

    result = {}

    # ── Sentence 1: Business quality / moat ──────────────────────────────────
    if roic is not None and rev_cagr is not None:
        if roic > 0.20 and rev_cagr > 0.10:
            result["moat"] = (
                f"{ticker} is a high-quality compounder generating {roic:.0%} ROIC on "
                f"{rev_cagr:.0%} revenue CAGR, indicating durable competitive advantages "
                f"and strong reinvestment economics."
            )
        elif roic > 0.20 and rev_cagr <= 0.10:
            result["moat"] = (
                f"{ticker} generates exceptional returns on capital ({roic:.0%} ROIC) with "
                f"moderate growth — a capital-light franchise with pricing power in a maturing market."
            )
        elif roic >= 0.12 and rev_cagr > 0.08:
            result["moat"] = (
                f"{ticker} combines solid capital returns ({roic:.0%} ROIC) with {rev_cagr:.0%} "
                f"revenue growth, suggesting an expanding competitive position."
            )
        elif roic >= 0.12 and rev_cagr <= 0.08:
            result["moat"] = (
                f"{ticker} produces steady returns ({roic:.0%} ROIC) in a low-growth environment "
                f"— valuation discipline and capital allocation are key."
            )
        elif roic < 0.12 and rev_cagr > 0.10:
            result["moat"] = (
                f"{ticker} is a high-growth business ({rev_cagr:.0%} revenue CAGR) with returns "
                f"still below cost of capital ({roic:.0%} ROIC) — profitability inflection is the "
                f"key watchpoint."
            )
        else:
            result["moat"] = (
                f"{ticker} operates with below-cost-of-capital returns ({roic:.0%} ROIC) "
                f"— value creation depends on margin expansion or capital efficiency improvement."
            )
    elif roic is not None:
        # Have ROIC but no rev_cagr
        if roic > 0.20:
            result["moat"] = (
                f"{ticker} generates exceptional returns on capital ({roic:.0%} ROIC) "
                f"— a capital-efficient franchise with durable competitive advantages."
            )
        elif roic >= 0.12:
            result["moat"] = (
                f"{ticker} produces solid returns on capital ({roic:.0%} ROIC) "
                f"— valuation discipline and capital allocation are key."
            )
        else:
            result["moat"] = (
                f"{ticker} operates with below-cost-of-capital returns ({roic:.0%} ROIC) "
                f"— value creation depends on margin expansion or capital efficiency improvement."
            )
    elif fcf_margin is not None and rev_cagr is not None:
        result["moat"] = (
            f"{ticker} generates {fcf_margin:.0%} FCF margins on {rev_cagr:.0%} revenue CAGR "
            f"— cash generation profile suggests competitive durability worth monitoring."
        )
    else:
        result["moat"] = None  # fallback to existing text

    # ── Sentence 2: Valuation vs history ─────────────────────────────────────
    _has_pe = pe_current is not None and pe_5yr is not None and pe_5yr > 0
    _has_dcf = (dcf_base_px is not None and current_price is not None
                and current_price > 0 and dcf_base_px > 0)
    _dcf_upside = ((dcf_base_px / current_price) - 1) if _has_dcf else None

    if _has_pe:
        _ratio = pe_current / pe_5yr
        _discount = 1 - _ratio   # positive means discount
        _premium = _ratio - 1    # positive means premium

        if _ratio < 0.85 and _has_dcf and _dcf_upside > 0:
            result["valuation"] = (
                f"Trading at {pe_current:.1f}x P/E vs. {pe_5yr:.1f}x 5-year average "
                f"({_discount:.0%} discount), with DCF base case implying {_dcf_upside:.0%} upside "
                f"— potentially undemanding if earnings hold."
            )
        elif _ratio < 0.85:
            result["valuation"] = (
                f"Trading at {pe_current:.1f}x P/E vs. {pe_5yr:.1f}x 5-year average "
                f"— a {_discount:.0%} discount to history suggesting the market is pricing in "
                f"deterioration."
            )
        elif _ratio > 1.15 and _has_dcf and _dcf_upside < 0:
            result["valuation"] = (
                f"At {pe_current:.1f}x P/E vs. {pe_5yr:.1f}x 5-year average (+{_premium:.0%} "
                f"premium), DCF base case implies {_dcf_upside:.0%} downside — premium only "
                f"justified if growth reaccelerates."
            )
        elif _ratio > 1.15:
            result["valuation"] = (
                f"At {pe_current:.1f}x P/E vs. {pe_5yr:.1f}x history, the stock trades at a "
                f"{_premium:.0%} premium — growth execution must remain flawless."
            )
        else:
            result["valuation"] = (
                f"Valuation appears broadly in line with history at {pe_current:.1f}x P/E vs. "
                f"{pe_5yr:.1f}x 5-year average — returns will likely track earnings rather than "
                f"multiple expansion."
            )
    elif pfcf_current is not None and pfcf_5yr is not None and pfcf_5yr > 0:
        _pf_ratio = pfcf_current / pfcf_5yr
        if _pf_ratio < 0.85:
            result["valuation"] = (
                f"Trading at {pfcf_current:.1f}x P/FCF vs. {pfcf_5yr:.1f}x 5-year average "
                f"— a discount to history that may reflect underappreciated cash generation."
            )
        elif _pf_ratio > 1.15:
            result["valuation"] = (
                f"At {pfcf_current:.1f}x P/FCF vs. {pfcf_5yr:.1f}x 5-year average, "
                f"the stock commands a premium — sustained FCF growth must validate the multiple."
            )
        else:
            result["valuation"] = (
                f"P/FCF of {pfcf_current:.1f}x vs. {pfcf_5yr:.1f}x 5-year average appears "
                f"broadly fair — returns will likely track cash flow growth."
            )
    else:
        result["valuation"] = None  # fallback to existing text

    # ── Sentence 3: Key risk or catalyst ─────────────────────────────────────
    if (d_ebitda is not None and d_ebitda > 3.5
            and ebit_interest is not None and ebit_interest < 3):
        result["risk"] = (
            f"Leverage of {d_ebitda:.1f}x EBITDA with {ebit_interest:.1f}x interest coverage "
            f"represents the primary downside risk in a higher-rate environment."
        )
    elif fcf_ni is not None and fcf_ni < 0.5:
        result["risk"] = (
            f"FCF conversion of {fcf_ni:.0%} of net income warrants monitoring — sustained "
            f"below-average conversion may signal earnings quality or heavy reinvestment needs."
        )
    elif gm_trend < -0.03:
        result["risk"] = (
            f"Gross margin compression of {abs(gm_trend)*100:.1f}pp over the review period is "
            f"the key risk — pricing power or input cost dynamics will determine whether "
            f"compression is cyclical or structural."
        )
    elif rev_cagr is not None and rev_cagr > 0.15:
        result["risk"] = (
            f"At {rev_cagr:.0%} revenue CAGR, execution risk and competitive response are key "
            f"watchpoints — premium valuation leaves little room for growth deceleration."
        )
    elif (d_ebitda is not None and fcf_ni is not None and rev_cagr is not None):
        result["risk"] = (
            f"With {d_ebitda:.1f}x leverage, {fcf_ni:.0%} FCF conversion, and {rev_cagr:.0%} "
            f"revenue CAGR, the base case is stable compounding — key upside optionality lies "
            f"in capital returns acceleration."
        )
    elif rev_cagr is not None and fcf_ni is not None:
        result["risk"] = (
            f"With {fcf_ni:.0%} FCF conversion and {rev_cagr:.0%} revenue CAGR, the business "
            f"offers a stable earnings profile — monitor for capital allocation catalysts."
        )
    elif d_ebitda is not None and d_ebitda > 3.5:
        result["risk"] = (
            f"Leverage of {d_ebitda:.1f}x EBITDA is elevated — debt reduction or refinancing "
            f"risk is the primary watchpoint in the current rate environment."
        )
    else:
        result["risk"] = None  # fallback to existing text

    return result


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
                      analyst_ests=None, analyst_targets=None):

    is0, bs0, cf0 = is_data[-1], bs_data[-1], cf_data[-1]
    today = datetime.date.today().strftime("%B %Y")
    dcf_prices = dcf_prices or {}

    # ── Company info ──────────────────────────────────────────────────────────
    company_name = profile.get("companyName") or ticker
    exchange     = profile.get("exchangeShortName") or ""
    industry     = profile.get("industry") or profile.get("sector") or profile.get("industryType") or ""
    ceo          = profile.get("ceo") or "N/A"
    description  = profile.get("description") or ""
    ceo_info     = f"CEO: {ceo}."

    current_price = float(current_price or profile.get("price") or 0) or 0.0
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

    # ── Fallbacks: compute ROIC / rev_cagr directly if scorecard cache is stale ─
    if roic_v is None:
        roic_v = _roic(is0, bs0)
    if rev_cagr_v is None and len(is_data) >= 2:
        _r0 = is_data[0].get("revenue") or 0
        _rn = is_data[-1].get("revenue") or 0
        _n  = len(is_data) - 1
        rev_cagr_v = (_rn / _r0) ** (1 / _n) - 1 if (_r0 > 0 and _rn > 0 and _n > 0) else None
    if fcf_ni_v is None:
        _ni0 = is0.get("netIncome") or 0
        fcf_ni_v = (fcf0 / _ni0) if _ni0 and _ni0 > 0 else None

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

    # 3-year average annual revenue growth
    _rev_yoys_3yr = []
    for _j in range(max(1, len(is_data) - 3), len(is_data)):
        _pr = is_data[_j-1].get("revenue") or 0
        _cr = is_data[_j].get("revenue") or 0
        if _pr > 0 and _cr:
            _rev_yoys_3yr.append(_cr / _pr - 1)
    _rev_3yr_avg = sum(_rev_yoys_3yr) / len(_rev_yoys_3yr) if _rev_yoys_3yr else None

    # ── Valuation multiples (LTM) ─────────────────────────────────────────────
    ev_ebitda = (ev / ebd0) if ev and ebd0 > 0 else None
    ev_rev    = (ev / rev0) if ev and rev0 > 0 else None
    ev_gp     = (ev / gp0)  if ev and gp0  > 0 else None
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

    # ── Revenue-based valuation fallback (for negative/zero FCF companies) ────
    # When DCF yields no price (negative FCF), use growth-adjusted EV/Revenue
    # multiples to produce a realistic fair-value range.
    _val_method = "DCF (Gordon Growth)"
    _rev_val_note = ""
    _need_rev_val = (base_px is None) and (rev0 > 0) and (shares > 0)

    if _need_rev_val and ev is not None:
        # Current observed EV/Revenue (market-implied)
        _ev_rev_obs = ev / rev0

        # Growth-calibrated target multiples
        _rc = rev_cagr_v or 0.0
        if _rc > 0.30:
            _rm_bear, _rm_base, _rm_bull = 3.0, 6.0, 10.0
            _rm_tier = "High growth (>30% CAGR)"
        elif _rc > 0.15:
            _rm_bear, _rm_base, _rm_bull = 2.0, 4.0, 7.0
            _rm_tier = "Growth (15–30% CAGR)"
        elif _rc > 0.05:
            _rm_bear, _rm_base, _rm_bull = 1.0, 2.5, 4.0
            _rm_tier = "Moderate growth (5–15% CAGR)"
        else:
            _rm_bear, _rm_base, _rm_bull = 0.5, 1.5, 2.5
            _rm_tier = "Low/negative growth"

        # Apply multiples to next-year forward revenue estimate
        _fwd_rev1 = rev0 * (1 + max(0, _rc))
        net_debt0 = debt0 - cash0   # positive = net debt; negative = net cash

        def _rev_px(mult):
            _eq = mult * _fwd_rev1 - net_debt0
            px  = _eq / shares if shares > 0 else None
            return round(px, 0) if px and px > 0 else None

        bear_px = _rev_px(_rm_bear)
        base_px = _rev_px(_rm_base)
        bull_px = _rev_px(_rm_bull)

        _val_method = "EV/Revenue (revenue multiple)"
        _rev_val_note = (
            f"DCF not applicable (trailing FCF negative: {_b(fcf0)}). "
            f"Valuation based on EV/Revenue multiples calibrated to {_pct(_rc)} revenue CAGR "
            f"({_rm_tier}). "
            f"Applied to FY{int(years[-1])+1}E revenue estimate ({_b(_fwd_rev1)}): "
            f"bear {_rm_bear:.1f}x, base {_rm_base:.1f}x, bull {_rm_bull:.1f}x EV/Revenue. "
            f"Current market-implied EV/Revenue: {_ev_rev_obs:.1f}x."
        )
        price_target = base_px or current_price

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

    # ── Pre-compute scenario and risk metrics ─────────────────────────────────
    _gm_vals = [(is_.get("grossProfit") or 0) / max(1, is_.get("revenue") or 1) for is_ in is_data]
    _gm_trend = _gm_vals[-1] - _gm_vals[0] if len(_gm_vals) >= 2 else 0
    _capex_intensity = capex0 / rev0 if rev0 > 0 else None

    # FCF trailing CAGR
    _fcf_first_raw = (cf_data[0].get("freeCashFlow") or
                      ((cf_data[0].get("operatingCashFlow") or 0) -
                       abs(cf_data[0].get("capitalExpenditure") or 0)))
    _n_fcf = len(cf_data) - 1
    fcf_cagr_v = (
        (fcf0 / _fcf_first_raw) ** (1 / _n_fcf) - 1
        if _fcf_first_raw and _fcf_first_raw > 0 and fcf0 > 0 and _n_fcf > 0
        else None
    )

    # Scenario growth parameters (BEAR/BASE/BULL)
    _bear_rev_g  = max(-0.05, (rev_cagr_v or 0.03) * 0.4)
    _base_rev_g  = rev_cagr_v or 0.05
    _bull_rev_g  = min(0.30, (rev_cagr_v or 0.05) * 1.4)
    _bear_margin = max(0.01, (ebitdam0 or 0.15) - 0.03)
    _bull_margin = min(0.60, (ebitdam0 or 0.15) + 0.02)
    _n_fwd       = 3

    _bear_fcf_fwd = (fcf0 * (1 + _bear_rev_g) ** _n_fwd) if fcf0 > 0 else None
    _base_fcf_fwd = (fcf0 * (1 + _base_rev_g) ** _n_fwd) if fcf0 > 0 else None
    _bull_fcf_fwd = (fcf0 * (1 + _bull_rev_g) ** _n_fwd) if fcf0 > 0 else None

    _bear_mult_pe = round((trailing_pe or 15) * 0.80, 1)
    _base_mult_pe = round(pe_5yr or (trailing_pe or 18), 1)
    _bull_mult_pe = round((trailing_pe or 18) * 1.15, 1)

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

    # Normalise manual qualitative tiers (MOD-HIGH / MOD-LOW both → MOD)
    def _norm(v):
        v = (v or "").strip().upper()
        if v in ("MOD-HIGH", "MOD-LOW", "MOD"): return "MOD"
        if v == "HIGH": return "HIGH"
        if v == "LOW":  return "LOW"
        return None
    t_bc  = _norm(biz_clarity)
    t_ltp = _norm(ltp)

    P = TIER_PTS
    p1 = round((P[t_bc or "MOD"]*2.5 + P["MOD"]*10.0 + P[t_ltp or "MOD"]*10.0 + P["MOD"]*7.5) / 10, 1)
    p2 = round((P[t_rev]*10.0 + P[t_fcf_ni]*10.0 + P["MOD"]*5.0 + P[t_roic]*7.5) / 10, 1)
    p3 = round((P[t_debd]*5.0 + P[t_eint]*7.5 + P["MOD"]*2.5) / 10, 1)
    p4 = round((P[t_pe]*10.0 + P[t_pfcf]*10.0) / 10, 1)
    # Use adj_score (Excel engine total + manual inputs) when available for accuracy
    final_score = adj_score or auto_score or round(p1 + p2 + p3 + p4, 1)

    # ── 5-year financial table ─────────────────────────────────────────────────
    fin = {}
    _prev_rev = _prev_ebd = _prev_ebit = None  # for YoY growth tracking
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
        fin[f"EBITINT_FY{i}"] = _x(_ei) if _ei is not None else "—"
        fin[f"OCF_FY{i}"]     = _m(ocf)
        fin[f"CAPEX_FY{i}"]   = f"({_m(cpx)})"
        fin[f"FCF_FY{i}"]     = _m(fcf)

        # YoY growth rates (blank for first year)
        fin[f"REV_YOY_FY{i}"]    = _pct(rev/_prev_rev - 1) if (_prev_rev and rev and _prev_rev > 0) else "—"
        fin[f"EBITDA_YOY_FY{i}"] = _pct(ebd/_prev_ebd - 1) if (_prev_ebd and ebd and _prev_ebd > 0) else "—"
        fin[f"EBIT_YOY_FY{i}"]   = _pct(ebit/_prev_ebit - 1) if (_prev_ebit and ebit and _prev_ebit != 0) else "—"
        _prev_rev = rev; _prev_ebd = ebd; _prev_ebit = ebit

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
    fcf_b_lst  = [(cf_.get("freeCashFlow") or
                   ((cf_.get("operatingCashFlow") or 0) - abs(cf_.get("capitalExpenditure") or 0))
                  ) / 1e9 for cf_ in cf_data]
    roic_lst   = [_roic(is_, bs_) for is_, bs_ in zip(is_data, bs_data)]
    roic_pct   = [round((r or 0)*100, 1) for r in roic_lst]

    # Shareholder returns: real data from CF statement
    # FMP uses different field names across API versions — try all known variants
    def _buyback(cf_):
        for key in ("commonStockRepurchased", "repurchaseOfCommonStock",
                    "stockRepurchase", "purchasesOfCommonStock"):
            v = cf_.get(key)
            if v and v != 0:
                return abs(v) / 1e9
        return 0.0

    buyback_b_lst = [_buyback(cf_) for cf_ in cf_data]
    fcf_ps_lst    = [round(
        (cf_.get("freeCashFlow") or
         (cf_.get("operatingCashFlow") or 0) - abs(cf_.get("capitalExpenditure") or 0))
        / shares, 2) if shares > 0 else 0
        for cf_ in cf_data]

    # ── Pre-compute commentary strings ────────────────────────────────────────
    def _div_paid(cf_):
        for key in ("dividendsPaid", "commonDividendsPaid",
                    "dividendsAndOtherCashDistributions"):
            v = cf_.get(key)
            if v and v != 0:
                return abs(v) / 1e9
        return 0.0

    div_b_lst      = [_div_paid(cf_) for cf_ in cf_data]
    total_buybacks = sum(buyback_b_lst)
    total_divs     = sum(div_b_lst)
    total_returns  = total_buybacks + total_divs
    n_yr           = len(cf_data)

    # Capital returns key metric card
    if total_returns < 0.05:
        _cr_val = "Minimal"
        _cr_sub = f"Capital primarily reinvested at {_pct(roic_v)} ROIC; direct shareholder returns negligible over {n_yr}yr."
    elif total_buybacks > total_divs * 2:
        _cr_val = f"${total_returns:.1f}B"
        _cr_sub = f"Buyback-led: ${total_buybacks:.1f}B repurchases + ${total_divs:.1f}B dividends over {n_yr}yr."
    elif total_divs > total_buybacks * 2:
        _cr_val = f"${total_returns:.1f}B"
        _cr_sub = f"Dividend-led: ${total_divs:.1f}B dividends + ${total_buybacks:.1f}B buybacks over {n_yr}yr."
    else:
        _cr_val = f"${total_returns:.1f}B"
        _cr_sub = f"Balanced returns: ${total_buybacks:.1f}B buybacks + ${total_divs:.1f}B dividends over {n_yr}yr."

    # Moat commentary
    _roic_trend_dir = (
        "improving" if len(roic_lst) >= 2 and (roic_lst[-1] or 0) > (roic_lst[0] or 0)
        else "declining" if len(roic_lst) >= 2 and (roic_lst[-1] or 0) < (roic_lst[0] or 0)
        else "stable"
    )
    _moat_strength = (
        "Strong"   if (roic_v or 0) > 0.20 and (gm0 or 0) > 0.40 else
        "Moderate" if (roic_v or 0) > 0.12 or  (gm0 or 0) > 0.35 else
        "Narrow"
    )
    _moat_commentary = (
        f"{_moat_strength} competitive position: {_pct(gm0)} gross margin "
        f"({'improving' if _gm_trend > 0.01 else 'declining' if _gm_trend < -0.01 else 'stable'} "
        f"over {len(is_data)}yr), ROIC {_pct(roic_v)} ({_roic_trend_dir})."
        + (" ROIC comfortably above cost of capital — durable advantage supported by financials."
           if (roic_v or 0) > 0.15
           else " ROIC near cost of capital; competitive advantage not yet clearly evident in returns.")
    )

    # Business Clarity commentary
    if t_bc:
        _bc_commentary = (
            f"Rated {t_bc}: "
            + ("clear, asset-light model with straightforward revenue drivers."
               if t_bc == "HIGH" else
               "moderately complex; key drivers identifiable from financials."
               if t_bc in ("MOD", "MOD-HIGH", "MOD-LOW") else
               "complex or opaque model requiring deeper qualitative diligence.")
        )
    else:
        _bc_model_type = (
            "High-margin, asset-light model" if (gm0 or 0) > 0.50 and _capex_intensity < 0.05
            else "Capital-efficient with moderate complexity" if (gm0 or 0) > 0.35
            else "Capital-intensive or complex operating model"
        )
        _bc_commentary = (
            f"{_bc_model_type}: {_pct(gm0)} gross margin, {_pct(fcf_ni_v)} FCF/NI conversion, "
            f"CapEx {_pct(_capex_intensity)} of revenue. Not manually rated."
        )

    # Long-term positioning commentary
    if t_ltp:
        _ltp_commentary = (
            f"Rated {t_ltp}: "
            + ("strong secular growth with high-return reinvestment opportunities."
               if t_ltp == "HIGH" else
               "solid runway; addressable market supports continued expansion."
               if t_ltp in ("MOD", "MOD-HIGH", "MOD-LOW") else
               "limited structural upside; mature category or significant headwinds.")
            + f" {_n_fcf}yr revenue CAGR {_pct(rev_cagr_v)}, ROIC {_pct(roic_v)}."
        )
    else:
        _ltp_desc = (
            "Strong secular growth with high-return reinvestment"
            if (rev_cagr_v or 0) > 0.12 and (roic_v or 0) > 0.15 else
            "Moderate runway; well-positioned within addressable market"
            if (rev_cagr_v or 0) > 0.05 else
            "Mature business; growth reliant on market expansion or M&A"
        )
        _ltp_commentary = (
            f"{_ltp_desc}. {_n_fcf}yr revenue CAGR: {_pct(rev_cagr_v)}, "
            f"ROIC {_pct(roic_v)}, FCF {_b(fcf0)} in FY{years[-1]}. Not manually rated."
        )

    # Management / capital allocation commentary
    _mgt_return_str = (
        f"${total_buybacks:.1f}B buybacks"
        + (f" + ${total_divs:.1f}B dividends" if total_divs > 0.01 else "")
        + f" over {n_yr}yr."
        if total_buybacks > 0.01 else
        f"${total_divs:.1f}B dividends over {n_yr}yr. No material buyback activity."
        if total_divs > 0.01 else
        f"Minimal direct returns over {n_yr}yr; capital reinvested at {_pct(roic_v)} ROIC."
    )
    _mgt_commentary = (
        ceo_info + f" ROIC {_pct(roic_v)} ({_roic_trend_dir} over {len(is_data)}yr). " + _mgt_return_str
    )

    # Capital returns scorecard commentary (P2)
    _p2_cr_commentary = _cr_sub

    # Execution risk commentary (P3)
    _er_parts = []
    if (fcf_ni_v or 0) < 0.60:
        _er_parts.append(f"FCF/NI {_pct(fcf_ni_v)} — watch working capital and capex discipline")
    else:
        _er_parts.append(f"FCF/NI {_pct(fcf_ni_v)} — healthy conversion, limited near-term execution concern")
    if _capex_intensity and _capex_intensity > 0.06:
        _er_parts.append(
            f"CapEx-intensive ({_pct(_capex_intensity)} of revenue, {_b(capex0)}) — delivery risk on major projects"
        )
    if abs(_gm_trend) > 0.015:
        _er_parts.append(
            f"Gross margin {'contracting' if _gm_trend < 0 else 'expanding'} {_pct(abs(_gm_trend))} over {len(is_data)}yr"
        )
    _er_commentary = (
        "; ".join(_er_parts) + "."
        if _er_parts else
        f"CapEx {_pct(_capex_intensity)} of revenue; FCF/NI {_pct(fcf_ni_v)}. No material execution flags from financial data."
    )

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
        "FORWARD_PE_EST":       "Awaiting analyst estimates",
        "FORWARD_PE_10YR":      _x(pe_5yr),
        "FORWARD_PE_DELTA":     pe_delta,
        "TRAILING_PFCF":        _x(trailing_pfc),
        "TRAILING_PFCF_10YR":   (_x(pfcf_5yr) + " (5yr avg)") if pfcf_5yr else "N/A",
        "TRAILING_PFCF_DELTA":  pfcf_delta,
        "FORWARD_PFCF":         "N/A",
        "FORWARD_PFCF_EST":     "Awaiting analyst estimates",
        "FORWARD_PFCF_10YR":    _x(pfcf_5yr),
        "FORWARD_PFCF_DELTA":   pfcf_delta,
        "EV_EBITDA_TRAILING":   _x(ev_ebitda),
        "EV_EBITDA_FWD_NOTE":   f"Trailing: {_x(ev_ebitda)}; fwd pending analyst estimates",
        "EV_REV_LTM":           _x(ev_rev),
        "EV_GP_LTM":            _x(ev_gp),
        "VALUATION_METHOD":     _val_method,
        "REV_VAL_NOTE":         _rev_val_note,

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
        "CAP_RETURNS_VALUE":    _cr_val,
        "CAP_RETURNS_SUB":      _cr_sub,

        # Financial Highlights (segment detail not available via FMP free tier)
        "REV_MIX_SECTION_LABEL": f"FY{years[-1]} · Segment detail in 10-K",
        "SEG1_EMOJI_NAME": "📈 Revenue Growth",
        "SEG1_REV_PCT": (_pct(rev_yoy) if rev_yoy else _b(rev0)),
        "SEG1_DESC": (
            f"Total revenue {_b(rev0)} in FY{years[-1]}. "
            f"3yr avg annual growth {_pct(_rev_3yr_avg)}; "
            f"{_n_fcf}yr CAGR {_pct(rev_cagr_v)}. Segment mix in 10-K."
        ),
        "SEG2_EMOJI_NAME": "💰 EBITDA Margin",
        "SEG2_REV_PCT": _pct(ebitdam0),
        "SEG2_DESC": f"EBITDA {_b(ebd0)} in FY{years[-1]}; margin {'improving' if _gm_trend > 0 else 'declining' if _gm_trend < 0 else 'stable'} over {_n_fcf}yr review period.",
        "SEG3_EMOJI_NAME": "🏦 FCF Margin",
        "SEG3_REV_PCT": _pct(fcf0/rev0 if rev0 else None),
        "SEG3_DESC": f"FCF {_b(fcf0)} in FY{years[-1]}, {_pct(fcf_ni_v)} FCF/NI conversion. {_n_fcf}yr FCF CAGR: {_pct(fcf_cagr_v) if fcf_cagr_v else 'N/A'}.",

        # Credit
        "SP_RATING":       sp_rating,    "SP_OUTLOOK":       "N/A",
        "SP_TIER_LABEL":   f"Tier: {cr_tier}",
        "MOODYS_RATING":   moody_rating, "MOODYS_OUTLOOK":   "N/A",
        "MOODYS_TIER_LABEL": f"Tier: {cr_tier}",
        "FITCH_RATING":    fitch_rating, "FITCH_OUTLOOK":    "NR",
        "FITCH_TIER_LABEL": "Tier: NR",
        "CREDIT_NOTE_TEXT": credit_commentary,

        # Thesis (auto-generated from financial data)
        "THESIS_MOAT_TEXT": (
            f"{company_name} ({ticker}) operates in {industry}. "
            f"FY{years[-1]} ROIC: {_pct(roic_v)} ({'above' if (roic_v or 0) > 0.15 else 'below'} 15% threshold), "
            f"gross margin: {_pct(gm0)}"
            + (f" ({'improving' if _gm_trend > 0.01 else 'declining' if _gm_trend < -0.01 else 'stable'} over review period)." if len(is_data) >= 2 else ".")
            + f" Revenue CAGR: {_pct(rev_cagr_v)}; FCF conversion: {_pct(fcf_ni_v)}."
        ),
        "THESIS_CATALYSTS_TEXT": (
            f"Base case: {_pct(_base_rev_g)} revenue growth with {_pct(ebitdam0)} EBITDA margins sustained. "
            f"Bear: growth slows to {_pct(_bear_rev_g)} with margin compression to {_pct(_bear_margin)}. "
            f"Bull: {_pct(_bull_rev_g)} revenue acceleration with margin expansion to {_pct(_bull_margin)}. "
            + (f"DCF range: ${bear_px:.0f}–${bull_px:.0f} (bear–bull)." if bear_px and bull_px else
               f"DCF base: ${base_px:.0f} (bear/bull not computed)." if base_px else "DCF range unavailable.")
            if base_px else
            f"FY{years[-1]} fundamentals: revenue {_b(rev0)}, EBITDA margin {_pct(ebitdam0)}, FCF {_b(fcf0)}."
        ),
        "THESIS_VALUATION_TEXT": (
            f"At ${current_price:.2f}, {ticker} trades at {_x(trailing_pe)} trailing P/E "
            f"({pe_delta} vs {_x(pe_5yr)} 5yr avg) and {_x(trailing_pfc)} P/FCF "
            f"({pfcf_delta} vs {_x(pfcf_5yr)} 5yr avg). "
            f"Gordon Growth DCF (WACC {wacc_b*100:.1f}%): base ${base_px:.0f} ({_vs(base_px, current_price)})"
            + (f", bear ${bear_px:.0f} ({_vs(bear_px, current_price)}), bull ${bull_px:.0f} ({_vs(bull_px, current_price)})." if bear_px and bull_px else ".")
            if base_px and current_price else
            f"Trailing P/E {_x(trailing_pe)} vs 5yr avg {_x(pe_5yr)}. "
            f"P/FCF {_x(trailing_pfc)} vs 5yr avg {_x(pfcf_5yr)}. "
            f"EV/EBITDA: {_x(ev_ebitda)}. See DCF tab."
        ),

        # Thesis risk (placeholder — overridden by _build_thesis below)
        "THESIS_RISK_TEXT": "",

        # Financials
        "CURRENCY_NAME":   "USD",
        "CURRENCY_SYMBOL": "$",
        "FIN_TABLE_NOTE":  f"FY{years[-1]} data from {company_name} annual report via FMP API; figures in USD millions.",

        # Growth context
        "REV_CAGR_TRAIL": f"Trailing {_n_fcf}yr: {_pct(rev_cagr_v)}",
        "REV_CAGR_FWD":   "Pending analyst estimates",
        "REV_CAGR_NOTE":  f"Source: FMP API. FY{years[0]}–{years[-1]}.",
        "FCF_CAGR_TRAIL": f"Trailing {_n_fcf}yr: {_pct(fcf_cagr_v)}" if fcf_cagr_v else "—",
        "FCF_CAGR_FWD":   "Pending analyst estimates",
        "FCF_CAGR_NOTE":  f"FCF: {_b(fcf0)} in FY{years[-1]}.",
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
        "P1_BC_COMMENTARY":     _bc_commentary,
        "P1_MOAT_SCORE_TEXT":   "MOD", "P1_MOAT_WTD": "7.0",
        "P1_MOAT_COMMENTARY":   _moat_commentary,
        "P1_LTP_SCORE_TEXT":    t_ltp or "MOD", "P1_LTP_WTD": str(round(P[t_ltp or "MOD"]*10.0/10, 1)),
        "P1_LTP_COMMENTARY":    _ltp_commentary,
        "P1_MGT_SCORE_TEXT":    "MOD", "P1_MGT_WTD": "5.25",
        "P1_MGT_COMMENTARY":    _mgt_commentary,

        "P2_WEIGHTED":          str(p2),
        "P2_RC_SCORE_TEXT":     t_rev,  "P2_RC_WTD": str(round(P[t_rev]*10.0/10, 1)),
        "P2_RC_COMMENTARY":     f"3yr revenue CAGR: {_pct(rev_cagr_v)}.",
        "P2_CQ_SCORE_TEXT":     t_fcf_ni, "P2_CQ_WTD": str(round(P[t_fcf_ni]*10.0/10, 1)),
        "P2_CQ_COMMENTARY":     f"FCF/NI: {_pct(fcf_ni_v)}.",
        "P2_CR_SCORE_TEXT":     "MOD",  "P2_CR_WTD": "3.5",
        "P2_CR_COMMENTARY":     _p2_cr_commentary,
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
        "P3_ER_COMMENTARY":     _er_commentary,

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
        "BEAR_REVENUE_GROWTH": f"{_pct(_bear_rev_g)} (stress — 40% of {_pct(rev_cagr_v)} trailing CAGR)",
        "BASE_REVENUE_GROWTH": f"{_pct(_base_rev_g)} CAGR (matches trailing {_n_fcf}yr)",
        "BULL_REVENUE_GROWTH": f"{_pct(_bull_rev_g)} (~140% of trailing CAGR — acceleration case)",
        "BEAR_MARGIN":   f"{_pct(_bear_margin)} EBITDA (−3pp compression)",
        "BASE_MARGIN":   f"{_pct(ebitdam0)} EBITDA (current rate maintained)",
        "BULL_MARGIN":   f"{_pct(_bull_margin)} EBITDA (+2pp expansion)",
        "FCF_YEAR_LABEL": f"FY{int(years[-1])+_n_fwd}E" if years else "Fwd",
        "BEAR_FCF": _b(_bear_fcf_fwd) if _bear_fcf_fwd else f"FCF neg (last: {_b(fcf0)})",
        "BASE_FCF": _b(_base_fcf_fwd) if _base_fcf_fwd else f"FCF neg (last: {_b(fcf0)})",
        "BULL_FCF": _b(_bull_fcf_fwd) if _bull_fcf_fwd else f"FCF neg (last: {_b(fcf0)})",
        "BEAR_MULTIPLE": f"{_bear_mult_pe:.1f}x P/E (de-rating to −20%)",
        "BASE_MULTIPLE": f"{_base_mult_pe:.1f}x P/E (5yr historical avg)",
        "BULL_MULTIPLE": f"{_bull_mult_pe:.1f}x P/E (re-rating to +15%)",
        "BEAR_CATALYST": "Execution miss, margin compression, macro headwinds",
        "BASE_CATALYST": "Consensus estimates met, stable macro",
        "BULL_CATALYST": "Revenue acceleration, margin expansion, multiple re-rating",
        "BEAR_UPSIDE": _vs(bear_px, current_price),
        "BASE_UPSIDE": _vs(base_px, current_price),
        "BULL_UPSIDE": _vs(bull_px, current_price),

        # Risks (auto-generated from computed metrics)
        "RISK_1_TITLE": "Competitive / Pricing Risk",
        "RISK_1_TEXT": (
            f"Gross margin {_pct(gm0)} in FY{years[-1]}"
            + (f", {'down ' if _gm_trend < -0.005 else 'up '}{_pct(abs(_gm_trend))} vs FY{years[0]}"
               if abs(_gm_trend) > 0.005 else ", stable over review period")
            + f". Revenue CAGR {_pct(rev_cagr_v)} — "
            + ("strong growth trajectory may attract competitive entry or invite pricing pressure."
               if (rev_cagr_v or 0) > 0.12
               else "moderate growth; monitor market share and pricing power in annual filings.")
        ),
        "RISK_2_TITLE": "Execution / Operating Risk",
        "RISK_2_TEXT": (
            f"CapEx {_pct(_capex_intensity)} of revenue ({_b(capex0)}). "
            f"EBITDA margin {_pct(ebitdam0)}; FCF conversion {_pct(fcf_ni_v)}."
            + (" FCF/NI below 50% — watch working capital and capex discipline."
               if (fcf_ni_v or 1) < 0.5
               else " Healthy FCF conversion suggests strong operational execution.")
        ),
        "RISK_3_TITLE": "Macro / Rate Risk",
        "RISK_3_TEXT": (
            f"Beta {beta_v:.2f}"
            + (" — elevated market sensitivity; material drawdown risk in risk-off periods."
               if beta_v > 1.3
               else " — moderate market correlation; relatively resilient to broad selloffs.")
            + f" {net_cash_str}."
            + (" Net debt increases refinancing exposure if rates remain elevated."
               if net_cash_val < 0
               else " Net cash provides strong macro buffer and optionality.")
        ),
        "RISK_4_TITLE": "Capital Allocation Risk",
        "RISK_4_TEXT": (
            f"Capital Adequacy (Equity/Assets): {_pct(equity_assets_v)}. Monitor regulatory CET1 ratio trends."
            if is_bank_v else
            f"D/EBITDA: {d_ebd_str}; interest coverage: {ebit_int_str}."
            + (" High leverage constrains financial flexibility and amplifies downside in a downturn."
               if (d_ebd or 0) > 3.5
               else " Conservative leverage with ample capacity for returns or strategic deployment.")
        ),
        "RISK_5_TITLE": "Valuation / Multiple Risk",
        "RISK_5_TEXT": (
            f"Trailing P/E {_x(trailing_pe)} vs. 5yr avg {_x(pe_5yr)} ({pe_delta}); "
            f"P/FCF {_x(trailing_pfc)} vs. 5yr avg {_x(pfcf_5yr)} ({pfcf_delta})."
            + (" Trading above historical averages — vulnerable to de-rating if earnings miss."
               if (trailing_pe and pe_5yr and trailing_pe > pe_5yr * 1.05)
               else " Valuation broadly in line with history — limited multiple compression risk absent a shock.")
        ),

        # Valuation verdict
        "VALUATION_VERDICT_TITLE": "Valuation Analysis",
        "VALUATION_VERDICT_TEXT": (
            f"At ${current_price:.2f}, {company_name} ({ticker}) trades at "
            + (f"{_x(trailing_pe)} trailing P/E vs. {_x(pe_5yr)} 5-year average ({pe_delta}), "
               f"and {_x(trailing_pfc)} P/FCF vs. {_x(pfcf_5yr)} 5-year average ({pfcf_delta}). "
               if (trailing_pe or trailing_pfc) else "")
            + f"EV/Revenue: {_x(ev_rev)}; EV/EBITDA: {_x(ev_ebitda)}. "
            + (f"{_val_method} base case: <strong>${base_px:.0f} ({_vs(base_px, current_price)})</strong>. "
               + (f"Bear: ${bear_px:.0f} ({_vs(bear_px, current_price)}) | Bull: ${bull_px:.0f} ({_vs(bull_px, current_price)})." if bear_px and bull_px else "")
               + (f"<br><em style='font-size:12px;color:var(--muted)'>{_rev_val_note}</em>" if _rev_val_note else "")
               if base_px else "Fair value not computable.")
        ) if current_price else (
            f"Valuation: EV/Revenue {_x(ev_rev)}, EV/EBITDA {_x(ev_ebitda)}, "
            f"P/E {_x(trailing_pe)}, P/FCF {_x(trailing_pfc)}."
        ),

        "DCF_BEAR_WACC": f"{w_bear*100:.1f}%", "DCF_BEAR_TGR": f"{tgr_bear*100:.1f}%",
        "DCF_BEAR_CAGR": _pct(_bear_rev_g) + " (bear — 40% of trailing)",
        "DCF_BEAR_PX":   f"${bear_px:.0f}" if bear_px else "N/A",
        "DCF_BEAR_VS":   _vs(bear_px, current_price),
        "DCF_BASE_WACC": f"{w_base*100:.1f}%", "DCF_BASE_TGR": f"{tgr_base*100:.1f}%",
        "DCF_BASE_CAGR": _pct(_base_rev_g) + " (base — trailing CAGR)",
        "DCF_BASE_PX":   f"${base_px:.0f}" if base_px else "N/A",
        "DCF_BASE_VS":   _vs(base_px, current_price),
        "DCF_BULL_WACC": f"{w_bull*100:.1f}%", "DCF_BULL_TGR": f"{tgr_bull*100:.1f}%",
        "DCF_BULL_CAGR": _pct(_bull_rev_g) + " (bull — 140% of trailing)",
        "DCF_BULL_PX":   f"${bull_px:.0f}" if bull_px else "N/A",
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
        "WACC_NOTE":      (f"WACC = {ew_v*100:.0f}% equity × {ke_approx*100:.1f}% Ke + {dw_v*100:.0f}% debt × {max(0,kd_pre)*100:.1f}% Kd × (1 − {eff_tax0*100:.0f}% tax) = {wacc_b*100:.1f}%. "
                           f"Beta {beta_v:.2f} (FMP 5yr), ERP {ERP_APPROX*100:.1f}% (Damodaran), Rf from FRED DGS10."),

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
        "EVEBITDA_10YR_AVG": _x(ev_ebitda) if ev_ebitda else "—",
        "EVEBITDA_5YR_AVG":  _x(ev_ebitda) if ev_ebitda else "—",
        **{k: "—" for k in [
            "MULT_PE10_FY1_PX","MULT_PE10_FY1_UPS","MULT_PE10_FY2_PX","MULT_PE10_FY2_UPS",
            "MULT_PE5_FY1_PX","MULT_PE5_FY1_UPS","MULT_PE5_FY2_PX","MULT_PE5_FY2_UPS",
            "MULT_PFCF10_FY1_PX","MULT_PFCF10_FY1_UPS","MULT_PFCF10_FY2_PX","MULT_PFCF10_FY2_UPS",
            "MULT_PFCF5_FY1_PX","MULT_PFCF5_FY1_UPS","MULT_PFCF5_FY2_PX","MULT_PFCF5_FY2_UPS",
            "MULT_EV10_FY1_PX","MULT_EV10_FY1_UPS","MULT_EV10_FY2_PX","MULT_EV10_FY2_UPS",
            "MULT_EV5_FY1_PX","MULT_EV5_FY1_UPS","MULT_EV5_FY2_PX","MULT_EV5_FY2_UPS",
        ]},
        "MULTIPLES_METHOD_RATIONALE": (
            f"Method: P/E ({pe_delta} vs 5yr avg {_x(pe_5yr)}), "
            f"P/FCF ({pfcf_delta} vs 5yr avg {_x(pfcf_5yr)}), EV/EBITDA {_x(ev_ebitda)}. "
            f"Forward multiples pending analyst estimates (FMP free tier)."
        ),
        "MULTIPLES_KEY_QUESTION":     (
            f"Key Question: Does the {pe_delta} to 5yr P/E reflect a re-rating opportunity or "
            f"justified caution given {_pct(rev_cagr_v)} revenue CAGR and {_pct(roic_v)} ROIC?"
        ),
        "COMPOSITE_FAIR_VALUE":       f"${base_px:.0f}" if base_px else "N/A",
        "COMPOSITE_UPSIDE_NOTE":      (
            (
                f"{_val_method} base: ${base_px:.0f} ({_vs(base_px, current_price)}); "
                + (f"bear ${bear_px:.0f} ({_vs(bear_px, current_price)}), bull ${bull_px:.0f} ({_vs(bull_px, current_price)})." if bear_px and bull_px else "bear/bull scenarios not computed.")
                + (f" {_rev_val_note}" if _rev_val_note else "")
            ) if base_px else
            f"No DCF price computed — FCF negative. Multiples: EV/Revenue {_x(ev_rev)}, P/E {_x(trailing_pe)}, EV/EBITDA {_x(ev_ebitda)}."
        ),

        # Analysts (auto-fetched where available via analyst_ests param)
        "TICKER_SHORT":   ticker,
        "BUY_PCT":        "—", "BUY_COUNT":  "—",
        "HOLD_PCT":       "—", "HOLD_COUNT": "—",
        "SELL_PCT":       "—", "SELL_COUNT": "—",
        **{f"A{i}_{k}": "—" for i in range(1,8)
           for k in ["NAME","FIRM","PT","PT_VS","DATE"]},
        **{f"A{i}_RATING_TEXT": "—" for i in range(1,8)},
        "ANALYST_COUNT":      "—",
        "CONSENSUS_PT":       "—",
        "CONSENSUS_PT_VS":    "—",
        "PT_RANGE":           "—",
        "ANALYST_TABLE_NOTE": f"Forward analyst estimates not available on FMP free tier. DCF-implied fair value: ${base_px:.0f} ({_vs(base_px, current_price)})." if base_px else "Forward analyst estimates not available on FMP free tier.",

        # Footnotes
        "FN1": f"Credit ratings: S&P {sp_rating} — manual input or Damodaran ICR model.",
        "FN2": f"Historical financials via FMP API. Fiscal year {years[-1]}.",
        "FN3": f"Revenue 3yr CAGR: {_pct(rev_cagr_v)} from FY{years[-4] if len(years)>=4 else years[0]}–{years[-1]}.",
        "FN4": "Share price history: sourced from FMP API where available.",
        "FN5": f"ROIC: {_pct(roic_v)} — NOPAT / Invested Capital from FMP data.",
        "FN6": "Valuation multiples: 5-year averages from FMP ratios API.",
        "FN7": f"Key risks: {_er_commentary[:120]}{'...' if len(_er_commentary) > 120 else ''}",
        "FN8": f"DCF: WACC {wacc_b*100:.1f}% (Damodaran-based), TGR 3.0%.",
        "DISCLAIMER_TEXT": "This report is auto-generated from public financial data. For informational purposes only. Not investment advice.",
    }

    D.update(fin)

    # ── Auto-generated thesis sentences ──────────────────────────────────────
    _fcf_margin_v = fcf0 / rev0 if rev0 > 0 else None
    _thesis_metrics = {
        "roic": roic_v,
        "rev_cagr": rev_cagr_v,
        "fcf_ni": fcf_ni_v,
        "pe_current": trailing_pe,
        "pe_5yr": pe_5yr,
        "dcf_base_px": base_px,
        "current_price": current_price,
        "d_ebitda": d_ebd,
        "ebit_interest": ebit_int,
        "gm_trend": _gm_trend,
        "fcf_margin": _fcf_margin_v,
        "pfcf_current": trailing_pfc,
        "pfcf_5yr": pfcf_5yr,
    }
    _thesis = _build_thesis(ticker, _thesis_metrics, years, is_data, cf_data)
    if _thesis.get("moat"):
        D["THESIS_MOAT_TEXT"] = _thesis["moat"]
    if _thesis.get("valuation"):
        D["THESIS_VALUATION_TEXT"] = _thesis["valuation"]
    if _thesis.get("risk"):
        D["THESIS_CATALYSTS_TEXT"] = _thesis["risk"]

    # ── Gemini AI: qualitative moat, growth catalysts, risks, scorecard ─────────
    # These prompts draw on Gemini's world knowledge — not just financial metrics.
    # Context is provided so responses are company-specific and analytical.
    _ai_fin_ctx = (
        f"Financial context (FY{years[-1]}): Revenue {_b(rev0)} ({_pct(rev_yoy)} YoY, "
        f"{_n_fcf}yr CAGR {_pct(rev_cagr_v)}); EBITDA margin {_pct(ebitdam0)}; "
        f"ROIC {_pct(roic_v)}; FCF {_b(fcf0)} ({_pct(fcf_ni_v)} FCF/NI); "
        f"D/EBITDA {_x(d_ebd)}; gross margin {_pct(gm0)}."
    )

    if GEMINI_KEY:
        # ── Section 4: The Moat ───────────────────────────────────────────────
        _ai_moat = _gemini(
            f"You are a buy-side equity analyst. Write a concise economic moat assessment "
            f"for {company_name} ({ticker}), sector: {industry}.\n"
            f"{_ai_fin_ctx}\n"
            f"In 3 sentences (max 90 words), assess: (1) the PRIMARY source of competitive "
            f"advantage — be specific: brand recognition, ecosystem lock-in, switching costs, "
            f"network effects, patents/IP, regulatory moat, or structural cost advantage; "
            f"(2) how durable this moat is and what specific threat could erode it. "
            f"Draw on your knowledge of this company's products, strategy, and competitive "
            f"dynamics. Go beyond the financial metrics above. No disclaimers or caveats."
        )
        if _ai_moat:
            D["THESIS_MOAT_TEXT"] = _ai_moat
            D["P1_MOAT_COMMENTARY"] = _ai_moat  # also used in scorecard

        # ── Section 4: Growth Catalysts ───────────────────────────────────────
        _ai_catalysts = _gemini(
            f"You are a buy-side equity analyst. Write a concise growth catalyst assessment "
            f"for {company_name} ({ticker}), sector: {industry}.\n"
            f"{_ai_fin_ctx}\n"
            f"In 3 sentences (max 90 words), identify: (1) the most important near-term "
            f"catalyst (e.g. new product launch, AI feature integration, pricing power, "
            f"regulatory approval, geographic expansion, margin inflection); "
            f"(2) the key structural long-term growth driver. "
            f"Be specific to this company's known strategy and product roadmap. "
            f"Do NOT just restate the financial metrics above. No disclaimers."
        )
        if _ai_catalysts:
            D["THESIS_CATALYSTS_TEXT"] = _ai_catalysts
            D["P1_LTP_COMMENTARY"] = _ai_catalysts  # also used in scorecard LTP row

        # ── Risk Factors: qualitative overlay ─────────────────────────────────
        _risk_ctx = (
            f"Beta {beta_v:.2f}; D/EBITDA {_x(d_ebd)}; interest coverage {_x(ebit_int)}; "
            f"gross margin {'declining' if _gm_trend < -0.01 else 'expanding' if _gm_trend > 0.01 else 'stable'} over {len(is_data)}yr."
        )
        _ai_risks = _gemini(
            f"You are a buy-side equity analyst. Write a concise risk assessment "
            f"for {company_name} ({ticker}), sector: {industry}.\n"
            f"{_ai_fin_ctx} {_risk_ctx}\n"
            f"In 3 sentences (max 90 words), identify: (1) the primary company-specific "
            f"business risk (e.g. competitive threat, product concentration, regulatory "
            f"scrutiny, customer concentration, technology disruption); "
            f"(2) the key macro or market risk (rate sensitivity, geopolitical exposure, FX, "
            f"commodity costs, valuation multiple risk). "
            f"Be specific to this company. Go beyond what the financial metrics already show. "
            f"No disclaimers or generic statements."
        )
        if _ai_risks:
            # Override the first two auto-generated risk items with the AI narrative
            D["RISK_1_TEXT"] = _ai_risks
            D["RISK_1_TITLE"] = "Primary Business Risk"
            D["RISK_2_TITLE"] = "AI-Assessed Macro / Structural Risk"
            D["RISK_2_TEXT"] = (
                f"Valuation: P/E {_x(trailing_pe)} vs. {_x(pe_5yr)} 5yr avg ({pe_delta}); "
                f"P/FCF {_x(trailing_pfc)} vs. {_x(pfcf_5yr)} 5yr avg ({pfcf_delta}). "
                + ("Multiple premium warrants flawless execution." if trailing_pe and pe_5yr and trailing_pe > pe_5yr * 1.05 else
                   "Valuation broadly in line with history.")
            )

    # EBIT/Interest note (kept for backwards compat but template now uses individual columns)
    ebitint_vals = [fin.get(f"EBITINT_FY{i}", "N/A") for i in range(1, len(years) + 1)]
    non_na = [(years[i], v) for i, v in enumerate(ebitint_vals) if v != "N/A"]
    D["EBITINT_NOTE"] = " · ".join(f"FY{yr}: {v}" for yr, v in non_na) if non_na else "N/A — interest data not available"

    # ── Analyst price targets (yfinance) ──────────────────────────────────────
    _at = analyst_targets or {}
    if not _at:
        try:
            import yfinance as _yf
            _yfi = _yf.Ticker(ticker).info
            _pt_mean = _yfi.get("targetMeanPrice")
            _pt_low  = _yfi.get("targetLowPrice")
            _pt_high = _yfi.get("targetHighPrice")
            _pt_n    = _yfi.get("numberOfAnalystOpinions")
            _rec     = (_yfi.get("recommendationKey") or "").replace("_", " ").upper()
            if _pt_mean and current_price:
                D["CONSENSUS_PT"]    = f"${_pt_mean:.2f}"
                D["CONSENSUS_PT_VS"] = _vs(_pt_mean, current_price)
            if _pt_n:
                D["ANALYST_COUNT"] = str(int(_pt_n))
            if _pt_low and _pt_high:
                D["PT_RANGE"] = f"${_pt_low:.2f} – ${_pt_high:.2f}"
            if _pt_mean:
                D["ANALYST_TABLE_NOTE"] = (
                    f"Consensus target ${_pt_mean:.2f} ({_vs(_pt_mean, current_price)}) · "
                    f"Range ${_pt_low:.2f}–${_pt_high:.2f} · {_pt_n} analysts · "
                    f"Consensus: {_rec}. Source: Yahoo Finance."
                )
            # Fill individual analyst rows from upgrades_downgrades
            try:
                _upgrades = _yf.Ticker(ticker).upgrades_downgrades
                if _upgrades is not None and not _upgrades.empty:
                    _upgrades = _upgrades.reset_index()
                    # Sort by date descending, take top 7
                    _upgrades = _upgrades.sort_values("GradeDate", ascending=False).head(7)
                    for _row_i, (_, _ug) in enumerate(_upgrades.iterrows(), 1):
                        _firm = str(_ug.get("Firm") or "")
                        _action = str(_ug.get("Action") or "")
                        _to_grade = str(_ug.get("ToGrade") or "")
                        _from_grade = str(_ug.get("FromGrade") or "")
                        _date = str(_ug.get("GradeDate", ""))[:10]
                        _rating_text = _to_grade or _action
                        D[f"A{_row_i}_NAME"]        = _action.title() if _action else "—"
                        D[f"A{_row_i}_FIRM"]        = _firm or "—"
                        D[f"A{_row_i}_RATING_TEXT"] = _rating_text or "—"
                        D[f"A{_row_i}_PT"]          = "—"
                        D[f"A{_row_i}_PT_VS"]       = "—"
                        D[f"A{_row_i}_DATE"]        = _date or "—"
            except Exception:
                pass
        except Exception:
            pass

    # ── Analyst estimates → forward multiples ─────────────────────────────────
    if analyst_ests:
        _ests = sorted(analyst_ests, key=lambda x: x.get("date", ""))
        _e1 = _ests[0] if _ests else {}
        _e2 = _ests[1] if len(_ests) >= 2 else {}

        _eps1  = _e1.get("estimatedEpsAvg") or 0
        _eps2  = _e2.get("estimatedEpsAvg") or 0
        _ebd1  = _e1.get("estimatedEbitdaAvg") or 0
        _ebd2  = _e2.get("estimatedEbitdaAvg") or 0
        _rev1  = _e1.get("estimatedRevenueAvg") or 0
        _rev2  = _e2.get("estimatedRevenueAvg") or 0

        # FCF/share proxy: forward EPS × historical FCF/NI conversion ratio
        _fn = max(0.3, min(fcf_ni_v or 0.7, 2.0))
        _fp1 = round(_eps1 * _fn, 2) if _eps1 > 0 else 0
        _fp2 = round(_eps2 * _fn, 2) if _eps2 > 0 else 0

        # Forward P/E and P/FCF
        if _eps1 > 0 and current_price:
            _fwd_pe   = round(current_price / _eps1, 1)
            _fwd_pe_d = _delta(_fwd_pe, pe_5yr)
            D["FORWARD_PE"]        = _x(_fwd_pe)
            D["FORWARD_PE_EST"]    = f"Based on FY+1E EPS ${_eps1:.2f}"
            D["FORWARD_PE_DELTA"]  = _fwd_pe_d
            D["PE_FWD_STAT"]       = _x(_fwd_pe)
        if _fp1 > 0 and current_price:
            _fwd_pfc  = round(current_price / _fp1, 1)
            _fwd_pf_d = _delta(_fwd_pfc, pfcf_5yr)
            D["FORWARD_PFCF"]      = _x(_fwd_pfc)
            D["FORWARD_PFCF_EST"]  = f"Based on FY+1E FCF/sh ${_fp1:.2f} (EPS × {_fn:.2f} conv.)"
            D["FORWARD_PFCF_DELTA"] = _fwd_pf_d
            D["PFCF_FWD_STAT"]     = _x(_fwd_pfc)

        # Forward revenue CAGR from estimates
        if _rev1 > 0 and rev0 > 0:
            _rev_fwd_cagr = (_rev1 / rev0) - 1
            D["REV_CAGR_FWD"]  = f"FY+1E: {_pct(_rev_fwd_cagr)} (vs {_pct(rev_cagr_v)} trailing)"
        if _rev2 > 0 and rev0 > 0 and _rev2 != _rev1:
            _rev_fwd2_cagr = (_rev2 / rev0) ** 0.5 - 1
            D["REV_CAGR_FWD"]  = f"FY+1E: {_pct((_rev1/rev0)-1 if _rev1 else 0)}, FY+2E 2yr CAGR: {_pct(_rev_fwd2_cagr)}"
        if _eps1 > 0:
            D["FCF_CAGR_FWD"] = f"FY+1E EPS ${_eps1:.2f} → FCF/sh ${_fp1:.2f} (×{_fn:.2f} conv.)"

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

        # Write — override stub values; preserve any real values already set
        for k, v in mu.items():
            if D.get(k) in ("—", "Add", "N/A", None, ""):
                D[k] = v
        # Always update composite (it starts as DCF-only; now includes multiples)
        for k in ("COMPOSITE_FAIR_VALUE", "COMPOSITE_UPSIDE_NOTE"):
            if k in mu:
                D[k] = mu[k]

    # ── Consistency check: key displayed values vs source data ────────────────
    # All financial data flows from FMP (is_data / bs_data / cf_data).
    # The checks below verify the summary metrics match the financial table rows
    # so that the HTML report and Excel model (same FMP source) stay in sync.
    _checks = {
        "Revenue": (D.get("REVENUE_VALUE"), _b(rev0)),
        "FCF":     (D.get("FCF_VALUE"),     _b(fcf0)),
        "EBITDA_margin": (D.get("EBITDA_MARGIN"), _pct(ebitdam0)),
        "ROIC":    (D.get("ROIC_VALUE"),    _pct(roic_v)),
    }
    _mismatches = [(k, a, b) for k, (a, b) in _checks.items() if a and b and a != b]
    if _mismatches:
        import sys
        for k, got, exp in _mismatches:
            print(f"[report_bridge] WARNING: {ticker} {k} mismatch — D has {got!r}, expected {exp!r}",
                  file=sys.stderr)

    # ── Data anomaly detection ────────────────────────────────────────────────
    try:
        anomalies = validate_fmp_data(ticker, is_data, bs_data, cf_data)
    except Exception:
        anomalies = []

    D["DATA_ANOMALIES"] = anomalies
    D["HAS_ANOMALIES"] = len([a for a in anomalies if a["severity"] in ("ERROR", "WARNING")]) > 0

    # Persist anomalies to static/data/anomalies.json (best-effort)
    if anomalies:
        try:
            persist_anomalies(ticker, anomalies, DATA_DIR)
        except Exception:
            pass

    return D


# ══════════════════════════════════════════════════════════════════════════════
# RENDER HTML
# ══════════════════════════════════════════════════════════════════════════════

def _build_anomaly_banner(data):
    """Build HTML for data quality anomaly banner. Returns empty string if no issues."""
    anomalies = data.get("DATA_ANOMALIES", [])
    has_anomalies = data.get("HAS_ANOMALIES", False)
    if not has_anomalies or not anomalies:
        return ""

    items_html = []
    for a in anomalies:
        sev = a.get("severity", "INFO")
        if sev not in ("ERROR", "WARNING"):
            continue
        yr = a.get("year", "")
        msg = a.get("message", "")
        items_html.append(
            f'<div class="anomaly-item anomaly-{sev.lower()}">'
            f'[{sev}] {yr}: {msg}</div>'
        )

    if not items_html:
        return ""

    return (
        '<div class="data-quality-banner">\n'
        '  <strong>Data Quality Notes</strong>\n'
        '  ' + '\n  '.join(items_html) + '\n'
        '</div>\n'
    )


_ANOMALY_CSS = """
/* ===== Data Quality Banner ===== */
.data-quality-banner {
  max-width: 1180px; margin: 0 auto 16px;
  padding: 16px 24px;
  background: #FFFBF0;
  border-left: 4px solid var(--amber);
  border-radius: var(--radius-sm);
  font-size: 13px;
  line-height: 1.6;
}
.data-quality-banner strong {
  display: block;
  font-size: 14px;
  margin-bottom: 8px;
  color: var(--ink);
}
.anomaly-item {
  padding: 4px 0;
  font-family: var(--font-mono);
  font-size: 12px;
}
.anomaly-error {
  color: var(--down);
  font-weight: 600;
}
.anomaly-warning {
  color: var(--amber);
}
.anomaly-info {
  color: var(--muted);
}
"""


def render_html_report(data):
    """Fill Report_Template.html with data dict, return HTML string."""

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    current_price = data.get("CURRENT_PRICE", 0) or 0

    # 0. Inject anomaly CSS and banner
    # CSS: insert before closing </style>
    if data.get("HAS_ANOMALIES"):
        html = html.replace("</style>", _ANOMALY_CSS + "</style>", 1)

    # Banner: insert between hero and tab nav
    anomaly_banner = _build_anomaly_banner(data)
    if anomaly_banner:
        # Primary: insert before the TAB NAV comment
        tab_marker = '<!-- ═══ TAB NAV ═══ -->'
        tab_pos = html.find(tab_marker)
        if tab_pos != -1:
            html = html[:tab_pos] + anomaly_banner + '\n' + html[tab_pos:]
        else:
            # Fallback: insert before <main
            main_pos = html.find('<main')
            if main_pos != -1:
                html = html[:main_pos] + anomaly_banner + '\n' + html[main_pos:]

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
