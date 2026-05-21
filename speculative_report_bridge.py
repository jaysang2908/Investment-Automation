"""
speculative_report_bridge.py — Maps speculative engine outputs to HTML template variables.

Reads Speculative_Report_Template.html, fills {{VARIABLE}} placeholders, returns HTML.
"""

import os
import datetime

# ── Helpers ────────────────────────────────────────────────────────────────────

def _fmt_price(v):
    if v is None:
        return "N/A"
    return f"${v:,.2f}"

def _fmt_pct(v, decimals=1):
    if v is None:
        return "N/A"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v*100:.{decimals}f}%"

def _fmt_x(v):
    if v is None:
        return "N/A"
    return f"{v:.1f}x"

def _fmt_b(v):
    if v is None:
        return "N/A"
    return f"${v:.2f}B"

def _fmt_m(v):
    if v is None:
        return "N/A"
    return f"${v:.0f}M"

def _fmt_num(v, dp=1):
    if v is None:
        return "N/A"
    return f"{v:,.{dp}f}"

def _tier_badge(tier: str) -> str:
    cls = {"HIGH": "tier-high", "MOD": "tier-mod", "LOW": "tier-low"}.get(tier, "tier-mod")
    return f'<span class="tier-badge {cls}">{tier}</span>'

def _ret_class(ret):
    if ret is None:
        return "neutral"
    return "pos" if ret >= 0 else "neg"

def _verdict_class(verdict: str) -> str:
    v = verdict.upper()
    if "MOONSHOT" in v:
        return "verdict-moonshot"
    if "STRONG" in v:
        return "verdict-strong"
    if "SPECULATIVE PLAY" in v or "SPECULATIVE" in v:
        return "verdict-spec"
    if "HIGH RISK" in v:
        return "verdict-risk"
    return "verdict-pass"

def _score_bar(score: int) -> str:
    pct = min(100, max(0, round(score / 120 * 100)))
    if pct >= 75:
        color = "#FF6B35"
    elif pct >= 60:
        color = "#FFA62B"
    elif pct >= 45:
        color = "#E6C168"
    else:
        color = "#6B7280"
    return (f'<div class="score-bar-wrap">'
            f'<div class="score-bar-fill" style="width:{pct}%;background:{color}"></div>'
            f'</div>')


# ── Report data builder ────────────────────────────────────────────────────────

def build_speculative_report_data(
    ticker: str,
    data: dict,
    scorecard: dict,
    scenario: dict,
) -> dict:
    """Build the template variable dict for the speculative HTML report."""

    scores  = scorecard.get("scores", {})
    total   = scorecard.get("total_score", 0)
    verdict = scorecard.get("verdict", "Pass")
    sc      = scenario

    today_str = datetime.date.today().strftime("%-d %B %Y") if hasattr(datetime.date.today(), 'strftime') else str(datetime.date.today())
    try:
        today_str = datetime.date.today().strftime("%d %B %Y").lstrip("0")
    except Exception:
        today_str = str(datetime.date.today())

    # ── Scorecard rows HTML ────────────────────────────────────────────────────
    SIGNAL_LABELS = [
        ("momentum",          "Price Momentum",           "RSI + vs 50d/200d MA"),
        ("volume",            "Volume Signal",            "10d vs 50d avg volume"),
        ("short_interest",    "Short Interest",           "Short float % + days to cover"),
        ("analyst_revisions", "Analyst Revision Momentum","Upgrades/downgrades + estimate drift"),
        ("float_size",        "Float / Market Cap",       "Smaller = more explosive"),
        ("insider_buying",    "Insider Buying",           "Open-market purchases, 90d"),
        ("downside_floor",    "Downside Floor",           "Net cash / debt position"),
        ("options_activity",  "Options Activity",         "Put/Call ratio"),
        ("technical_setup",   "Technical Setup",          "MACD crossover · 52w high proximity · EMA stack"),
        ("social_trend",      "Social / Trend Momentum",  "News sentiment · Reddit mentions"),
        ("narrative",         "Narrative Theme",          "User-supplied + theme heat"),
        ("catalyst",          "Catalyst Quality",         "User-supplied + timing anchor"),
    ]

    scorecard_rows = ""
    for key, label, sub in SIGNAL_LABELS:
        s    = scores.get(key, {})
        tier = s.get("tier", "MOD")
        pts  = s.get("pts", 5)
        note = s.get("note", "")
        is_manual = key in ("narrative", "catalyst")
        manual_tag = '<span class="manual-tag">manual</span>' if is_manual else ""
        scorecard_rows += f"""
        <tr class="sc-row">
          <td class="sc-signal">
            <div class="sc-label">{label} {manual_tag}</div>
            <div class="sc-sub">{sub}</div>
          </td>
          <td class="sc-tier">{_tier_badge(tier)}</td>
          <td class="sc-pts">{pts}</td>
          <td class="sc-note">{note}</td>
        </tr>"""

    # ── Scenario rows HTML ─────────────────────────────────────────────────────
    def _scenario_row(label, mult, price, ret, css_class):
        mult_str  = _fmt_x(mult)  if mult  else "N/A"
        price_str = _fmt_price(price) if price else "N/A"
        ret_str   = _fmt_pct(ret)     if ret   is not None else "N/A"
        ret_css   = "pos" if ret is not None and ret >= 0 else "neg"
        return f"""
        <tr class="sc-row {css_class}">
          <td class="scenario-label">{label}</td>
          <td class="scenario-val">{mult_str}</td>
          <td class="scenario-val">{price_str}</td>
          <td class="scenario-val {ret_css}">{ret_str}</td>
        </tr>"""

    scenario_rows = ""
    scenario_rows += _scenario_row(
        "BEAR — Narrative Fails",
        sc.get("bear_mult"), sc.get("bear_price"), sc.get("bear_ret"), "row-bear"
    )
    scenario_rows += _scenario_row(
        "BASE — Revenue Growth Only",
        sc.get("base_mult"), sc.get("base_price"), sc.get("base_ret"), "row-base"
    )
    scenario_rows += _scenario_row(
        "BULL — Narrative Plays Out",
        sc.get("bull_mult"), sc.get("bull_price"), sc.get("bull_ret"), "row-bull"
    )

    # Probability-weighted expected return row — the headline number
    p_bear = sc.get("p_bear")
    p_base = sc.get("p_base")
    p_bull = sc.get("p_bull")
    exp_ret = sc.get("expected_ret")
    exp_px  = sc.get("expected_price")
    if p_bear is not None and p_base is not None and p_bull is not None:
        prob_label = (f"EXPECTED — Prob-Weighted "
                      f"({int(round(p_bear*100))}/{int(round(p_base*100))}/{int(round(p_bull*100))})")
    else:
        prob_label = "EXPECTED — Prob-Weighted"
    exp_ret_str = _fmt_pct(exp_ret) if exp_ret is not None else "N/A"
    exp_px_str  = _fmt_price(exp_px) if exp_px is not None else "N/A"
    exp_css     = "pos" if (exp_ret is not None and exp_ret >= 0) else "neg"
    scenario_rows += f"""
        <tr class="sc-row" style="border-top:2px solid var(--border)">
          <td class="scenario-label" style="color:var(--orange);font-weight:600">{prob_label}</td>
          <td class="scenario-val" style="color:var(--ink-3)">—</td>
          <td class="scenario-val" style="color:var(--orange)">{exp_px_str}</td>
          <td class="scenario-val {exp_css}" style="font-weight:700">{exp_ret_str}</td>
        </tr>"""

    # 1.5x note
    bull_hits_1_5x = sc.get("bull_reaches_1_5x", False)
    target_1_5x    = _fmt_price(sc.get("target_1_5x_price"))
    req_mult       = _fmt_x(sc.get("req_mult_for_1_5x"))
    if bull_hits_1_5x:
        target_1_5x_html = (f'<div class="target-note target-hit">'
                            f'✓ Bull scenario reaches 1.5× target ({target_1_5x}) — '
                            f'requires EV/Rev of {req_mult}</div>')
    else:
        target_1_5x_html = (f'<div class="target-note target-miss">'
                            f'✗ Bull scenario does not reach 1.5× ({target_1_5x}) at current assumptions — '
                            f'narrative must be exceptionally strong or multiple expansion more aggressive</div>')

    # ── Catalyst block ────────────────────────────────────────────────────────
    cat_desc    = scorecard.get("catalyst_desc", "") or "Not specified"
    cat_timing  = scorecard.get("catalyst_timing", "vague") or "vague"
    timing_map  = {"near": "Near-term (< 3 months)", "medium": "Medium-term (3–9 months)", "vague": "Vague / unspecified"}
    cat_timing_label = timing_map.get(cat_timing, cat_timing)

    narr_theme  = scorecard.get("narrative_theme", "") or "Not specified"
    narr_str    = scorecard.get("narrative_strength", "MOD") or "MOD"

    # ── Mock warning ─────────────────────────────────────────────────────────
    mock_banner = ""
    if data.get("mock"):
        mock_banner = """
        <div class="mock-banner">
          ⚠ MOCK DATA MODE — all quantitative signals are illustrative placeholders.
          Re-run with live FMP API access for real analysis.
        </div>"""

    return {
        "TICKER":               ticker,
        "COMPANY_NAME":         data.get("company_name") or ticker,
        "SECTOR":               data.get("sector") or "Unknown",
        "REPORT_DATE":          today_str,
        "CURRENT_PRICE":        _fmt_price(data.get("current_price")),
        "MARKET_CAP":           _fmt_b((data.get("market_cap") or 0) / 1e9),
        "FLOAT_SHARES_M":       _fmt_num((data.get("float_shares") or 0) / 1e6, 0),
        "NET_DEBT_M":           _fmt_m((data.get("net_debt") or 0) / 1e6),
        "TRAIL_REV_B":          _fmt_b((data.get("trailing_rev") or 0) / 1e9),
        "FWD_REV_B":            _fmt_b((data.get("fwd_rev_1yr") or 0) / 1e9),
        "CURRENT_EV_B":         _fmt_b((sc.get("current_ev_b") or 0)),
        "CURRENT_EV_REV":       _fmt_x(sc.get("current_ev_rev_mult")),
        # Scorecard
        "TOTAL_SCORE":          str(total),
        "VERDICT":              verdict,
        "VERDICT_CLASS":        _verdict_class(verdict),
        "SCORE_BAR":            _score_bar(total),
        "SCORECARD_ROWS":       scorecard_rows,
        # Scenario
        "HOLD_MONTHS":          str(sc.get("hold_months", 6)),
        "SCENARIO_ROWS":        scenario_rows,
        "TARGET_1_5X":          target_1_5x,
        "REQ_MULT_1_5X":        req_mult,
        "TARGET_1_5X_HTML":     target_1_5x_html,
        "BEAR_PRICE":           _fmt_price(sc.get("bear_price")),
        "BASE_PRICE":           _fmt_price(sc.get("base_price")),
        "BULL_PRICE":           _fmt_price(sc.get("bull_price")),
        "BEAR_RET":             _fmt_pct(sc.get("bear_ret")),
        "BASE_RET":             _fmt_pct(sc.get("base_ret")),
        "BULL_RET":             _fmt_pct(sc.get("bull_ret")),
        "EXPECTED_RET":         _fmt_pct(sc.get("expected_ret")),
        "EXPECTED_PRICE":       _fmt_price(sc.get("expected_price")),
        "SECTOR_BUCKET":        (sc.get("sector_bucket") or "default").replace("_", " "),
        "PROB_WEIGHTS":         (f"{int(round((sc.get('p_bear') or 0)*100))} / "
                                  f"{int(round((sc.get('p_base') or 0)*100))} / "
                                  f"{int(round((sc.get('p_bull') or 0)*100))}"),
        "BEAR_FACTOR":          f"{sc.get('bear_factor'):.2f}×" if sc.get('bear_factor') else "—",
        "BULL_FACTOR":          f"{sc.get('bull_factor'):.2f}×" if sc.get('bull_factor') else "—",
        # Catalyst / Narrative
        "NARRATIVE_THEME":      narr_theme,
        "NARRATIVE_STRENGTH":   narr_str,
        "CATALYST_DESC":        cat_desc,
        "CATALYST_TIMING":      cat_timing_label,
        # Misc
        "MOCK_BANNER":          mock_banner,
        "RSI":                  _fmt_num(data.get("rsi_14"), 1),
        "SHORT_FLOAT":          f"{data.get('short_float_pct'):.1f}%" if data.get('short_float_pct') else "N/A",
        "PUT_CALL":             _fmt_num(data.get("put_call_ratio"), 2),
        "UPGRADES_90D":         str(data.get("upgrades_90d") or 0),
        "DOWNGRADES_90D":       str(data.get("downgrades_90d") or 0),
        "INSIDER_BUYS":         str(data.get("insider_buys") or 0),
    }


def render_speculative_report(report_data: dict) -> str:
    """Load template, substitute variables, return completed HTML."""
    tpl_path = os.path.join(os.path.dirname(__file__), "Speculative_Report_Template.html")
    with open(tpl_path, "r", encoding="utf-8") as f:
        html = f.read()

    for key, val in report_data.items():
        html = html.replace("{{" + key + "}}", str(val) if val is not None else "")

    return html
