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

# ── Parabolic precursor checklist ─────────────────────────────────────────────
# The 8 leading indicators that, in combination, historically precede outsized
# (1.5x+) moves over 3–12 months. Each is rendered as a colored tile so the
# user can scan setup quality at a glance.

def _build_parabolic_checklist(data: dict) -> dict:
    """Build the precursor checklist dict + rendered HTML.

    Returns:
        {
          "html":         rendered <div class="precursor-grid">…
          "n_firing":     int (0–8)
          "n_partial":    int
          "verdict":      "Setup Active" | "Partial Setup" | "Setup Not Present"
          "verdict_cls":  "setup-active" | "setup-partial" | "setup-dormant"
        }
    """
    # status: "fire" (firing — bullish), "part" (partial / watch), "off" (dormant / no signal)
    items: list[dict] = []

    # 1. Volatility coiling / breakout (BB squeeze → expansion)
    bb_sq, bb_exp, bb_w = data.get("bb_squeeze"), data.get("bb_expansion"), data.get("bb_width_pct")
    if bb_exp:
        items.append({"name": "Coiled-Spring Breakout", "status": "fire",
                      "detail": f"BB width expanding from squeeze · {bb_w*100:.1f}%" if bb_w else "BB expanding from squeeze",
                      "tooltip": "Bollinger Bands compressed to a low-volatility regime, now expanding — the classic 'coiled spring' release before a directional move."})
    elif bb_sq:
        items.append({"name": "Volatility Coiling", "status": "part",
                      "detail": f"BB squeeze active · width {bb_w*100:.1f}%" if bb_w else "BB squeeze active",
                      "tooltip": "Bollinger Bands in lowest 25% of recent volatility — coiled spring forming. Waiting for the breakout candle to confirm direction."})
    elif bb_w is not None:
        items.append({"name": "Volatility Coiling", "status": "off",
                      "detail": f"normal-range volatility · width {bb_w*100:.1f}%",
                      "tooltip": "BB width is in its normal range — no compression. Parabolic moves typically launch from a contracted volatility regime, not from normal-range chop."})
    else:
        items.append({"name": "Volatility Coiling", "status": "off", "detail": "data unavailable", "tooltip": "Need ≥60 days of OHLCV to compute BB squeeze."})

    # 2. OBV Divergence (smart money accumulation)
    obv_div, obv_rise, obv_slope = data.get("obv_divergence"), data.get("obv_rising"), data.get("obv_slope_pct")
    if obv_div:
        items.append({"name": "Smart-Money Divergence", "status": "fire",
                      "detail": f"OBV {obv_slope*100:+.1f}% with price flat",
                      "tooltip": "On-Balance Volume rising while price stays flat means net buying volume is accumulating without lifting the tape — institutional accumulation footprint."})
    elif obv_rise:
        items.append({"name": "OBV Accumulation", "status": "part",
                      "detail": f"OBV trend {obv_slope*100:+.1f}%",
                      "tooltip": "Cumulative on-balance volume trending up. Less bullish than a pure divergence (price is moving with it), but still a healthy accumulation signal."})
    elif obv_slope is not None and obv_slope < -0.05:
        items.append({"name": "OBV Accumulation", "status": "off",
                      "detail": f"OBV {obv_slope*100:+.1f}% — distribution",
                      "tooltip": "Cumulative volume signal is negative — net selling pressure dominates."})
    else:
        items.append({"name": "OBV Accumulation", "status": "off", "detail": "flat / unavailable",
                      "tooltip": "OBV is flat — no clear accumulation or distribution signal from the volume tape."})

    # 3. Up-Day Volume Dominance
    tier = data.get("acc_dist_tier")
    share = data.get("up_vol_share")
    if tier == "accumulation":
        items.append({"name": "Up-Day Volume Dominance", "status": "fire",
                      "detail": f"{share*100:.0f}% of 20d vol on up days",
                      "tooltip": "Over 60% of last 20 days' volume occurred on up-close days — buyers in control of the tape on a day-by-day basis."})
    elif tier == "neutral":
        items.append({"name": "Up-Day Volume Dominance", "status": "part",
                      "detail": f"{share*100:.0f}% up-day vol · neutral",
                      "tooltip": "Volume split roughly 50/50 between up and down days — no clear daily-tape directional bias."})
    elif tier == "distribution":
        items.append({"name": "Up-Day Volume Dominance", "status": "off",
                      "detail": f"{share*100:.0f}% up-day vol · distribution",
                      "tooltip": "Most of recent volume occurred on down-close days — distribution signature."})
    else:
        items.append({"name": "Up-Day Volume Dominance", "status": "off", "detail": "data unavailable", "tooltip": ""})

    # 4. EMA stack alignment
    price, ema21, ema50 = data.get("current_price"), data.get("ema_21"), data.get("ema_50")
    if price and ema21 and ema50:
        if price > ema21 > ema50:
            items.append({"name": "EMA Stack Aligned", "status": "fire",
                          "detail": f"${price:.2f} > 21EMA ${ema21:.2f} > 50EMA ${ema50:.2f}",
                          "tooltip": "Price above the 21-EMA, which is above the 50-EMA — short-term and intermediate trends both pointing up. The first thing trend-followers screen for."})
        elif price > ema21 and ema21 < ema50:
            items.append({"name": "EMA Stack Aligned", "status": "part",
                          "detail": "21EMA below 50EMA — early recovery",
                          "tooltip": "Price above the 21-EMA but the 21 is still below the 50 — early-stage trend reversal forming. Watch for the 21 to cross above the 50."})
        else:
            items.append({"name": "EMA Stack Aligned", "status": "off", "detail": "stack broken",
                          "tooltip": "EMA stack is not bullish — short-term trend is below intermediate trend or price below both."})
    else:
        items.append({"name": "EMA Stack Aligned", "status": "off", "detail": "data unavailable", "tooltip": ""})

    # 5. MACD momentum
    macd, sig, hist = data.get("macd"), data.get("macd_signal"), data.get("macd_hist")
    if macd is not None and sig is not None and hist is not None:
        if hist > 0 and macd > sig:
            items.append({"name": "MACD Momentum", "status": "fire",
                          "detail": f"hist {hist:+.2f} · MACD > signal",
                          "tooltip": "MACD line above its signal line with a positive (and ideally expanding) histogram — momentum is positive and improving."})
        elif hist > 0:
            items.append({"name": "MACD Momentum", "status": "part",
                          "detail": f"hist {hist:+.2f} · crossover pending",
                          "tooltip": "MACD histogram is positive but the MACD line hasn't yet crossed above its signal — momentum building."})
        else:
            items.append({"name": "MACD Momentum", "status": "off",
                          "detail": f"hist {hist:+.2f} · negative",
                          "tooltip": "MACD histogram is negative — momentum has not yet turned bullish."})
    else:
        items.append({"name": "MACD Momentum", "status": "off", "detail": "data unavailable", "tooltip": ""})

    # 6. RSI in launch zone (45-70, rising)
    rsi = data.get("rsi_14")
    if rsi is not None:
        if 50 <= rsi <= 70:
            items.append({"name": "RSI Launch Zone", "status": "fire", "detail": f"RSI {rsi:.0f} · 50-70",
                          "tooltip": "RSI between 50 and 70 is the 'sweet spot' for a building parabolic move — momentum positive but not yet overbought. Above 75 is extended; below 50 means no momentum."})
        elif 40 <= rsi < 50:
            items.append({"name": "RSI Launch Zone", "status": "part", "detail": f"RSI {rsi:.0f} · recovering",
                          "tooltip": "RSI recovering from the lower half — watching for a cross above 50 to confirm momentum has flipped bullish."})
        elif rsi > 70:
            items.append({"name": "RSI Launch Zone", "status": "part", "detail": f"RSI {rsi:.0f} · extended",
                          "tooltip": "RSI above 70 — already in overbought territory. Strong trend, but the easy 'expansion from a low base' move is behind us."})
        else:
            items.append({"name": "RSI Launch Zone", "status": "off", "detail": f"RSI {rsi:.0f} · weak",
                          "tooltip": "RSI below 40 — no positive momentum to work with. Parabolic moves don't start from here without an extreme catalyst."})
    else:
        items.append({"name": "RSI Launch Zone", "status": "off", "detail": "data unavailable", "tooltip": ""})

    # 7. Near 52-week high
    pct_off = data.get("pct_from_52w_high")
    if pct_off is not None:
        if pct_off >= -0.08:
            items.append({"name": "Breakout Proximity", "status": "fire", "detail": f"{pct_off*100:+.1f}% from 52w high",
                          "tooltip": "Within 8% of the 52-week high — in the breakout zone. New highs attract trend-followers and remove all 'stuck-in-loss' overhead supply."})
        elif pct_off >= -0.20:
            items.append({"name": "Breakout Proximity", "status": "part", "detail": f"{pct_off*100:+.1f}% from 52w high",
                          "tooltip": "Within 20% of the 52w high — climbing back but still significant overhead supply from the prior highs to absorb."})
        else:
            items.append({"name": "Breakout Proximity", "status": "off", "detail": f"{pct_off*100:+.1f}% from 52w high",
                          "tooltip": "More than 20% below the 52w high — significant base-building required before a breakout to new highs is realistic."})
    else:
        items.append({"name": "Breakout Proximity", "status": "off", "detail": "data unavailable", "tooltip": ""})

    # 8. ROC accelerating
    roc, roc_acc = data.get("roc_20d"), data.get("roc_accelerating")
    if roc_acc:
        items.append({"name": "ROC Acceleration", "status": "fire", "detail": f"20d ROC {roc*100:+.1f}% · accelerating",
                      "tooltip": "20-day rate of change is positive AND higher than it was 5 days ago — momentum isn't just positive, it's strengthening. The 'second derivative' signal."})
    elif roc is not None and roc > 0:
        items.append({"name": "ROC Acceleration", "status": "part", "detail": f"20d ROC {roc*100:+.1f}% · positive flat",
                      "tooltip": "ROC positive but not accelerating — momentum is holding steady rather than building. Watch for an inflection."})
    elif roc is not None:
        items.append({"name": "ROC Acceleration", "status": "off", "detail": f"20d ROC {roc*100:+.1f}%",
                      "tooltip": "20-day rate of change is negative — momentum is working against the setup."})
    else:
        items.append({"name": "ROC Acceleration", "status": "off", "detail": "data unavailable", "tooltip": ""})

    n_fire    = sum(1 for it in items if it["status"] == "fire")
    n_partial = sum(1 for it in items if it["status"] == "part")

    if n_fire >= 6:
        verdict, vc = "Setup Active — Multi-Signal Confluence", "setup-active"
    elif n_fire >= 4:
        verdict, vc = "Setup Building — Watch List Quality", "setup-active"
    elif n_fire + n_partial >= 4:
        verdict, vc = "Partial Setup — Wait for Confirmation", "setup-partial"
    else:
        verdict, vc = "Setup Not Present — Structure Too Weak", "setup-dormant"

    # Render HTML grid (4x2)
    tiles = []
    for it in items:
        st = it["status"]
        sym = {"fire": "●", "part": "◐", "off": "○"}[st]
        tip = it.get("tooltip", "") or ""
        tip_attr = f' title="{tip.replace(chr(34), chr(39))}"' if tip else ""
        tiles.append(f"""
        <div class="precursor-tile precursor-{st}"{tip_attr}>
          <div class="precursor-dot">{sym}</div>
          <div class="precursor-body">
            <div class="precursor-name">{it['name']}</div>
            <div class="precursor-detail">{it['detail']}</div>
          </div>
        </div>""")

    html = f"""
    <div class="precursor-header">
      <div class="precursor-summary">
        <div class="precursor-summary-num">{n_fire}<span class="precursor-summary-of"> of 8</span></div>
        <div class="precursor-summary-lbl">Precursors Firing</div>
      </div>
      <div class="precursor-verdict precursor-verdict-{vc}">{verdict}</div>
    </div>
    <div class="precursor-grid">
      {''.join(tiles)}
    </div>"""

    return {
        "html":        html,
        "n_firing":    n_fire,
        "n_partial":   n_partial,
        "verdict":     verdict,
        "verdict_cls": vc,
    }


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

    # %-d is Linux-only; %#d is Windows-only. Use %d + lstrip("0") for portability.
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
        ("technical_setup",   "Technical & Parabolic Setup", "MACD · 52w high · EMA stack · BB squeeze · OBV · acc/dist · ROC"),
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

    # ── Parabolic Precursor Dashboard ────────────────────────────────────────
    precursor = _build_parabolic_checklist(data)

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
        # Parabolic precursor dashboard
        "PRECURSOR_DASHBOARD":  precursor["html"],
        "PRECURSOR_FIRING":     str(precursor["n_firing"]),
        "PRECURSOR_VERDICT":    precursor["verdict"],
        "PRECURSOR_VERDICT_CLS": precursor["verdict_cls"],
    }


def render_speculative_report(report_data: dict) -> str:
    """Load template, substitute variables, return completed HTML."""
    tpl_path = os.path.join(os.path.dirname(__file__), "Speculative_Report_Template.html")
    with open(tpl_path, "r", encoding="utf-8") as f:
        html = f.read()

    for key, val in report_data.items():
        html = html.replace("{{" + key + "}}", str(val) if val is not None else "")

    return html
