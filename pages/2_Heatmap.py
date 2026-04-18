"""
Portfolio Heatmap — reads outputs.csv from GitHub and renders a colour-coded table
plus a horizontal stacked verdict chart.
Must NOT call st.set_page_config() — only app.py does that.
"""

import streamlit as st
import pandas as pd
import requests
from io import StringIO

# ── Password gate ─────────────────────────────────────────────────────────────
password = st.text_input("Password", type="password", key="heatmap_pw")
if not password:
    st.stop()
if password != st.secrets["APP_PASSWORD"]:
    st.error("Incorrect password.")
    st.stop()

# ── Config ────────────────────────────────────────────────────────────────────
REPO    = st.secrets.get("GITHUB_REPO", "jaysang2908/Investment-Automation")
CSV_URL = f"https://raw.githubusercontent.com/{REPO}/main/outputs.csv"

AUTO_MAX   = 87.5   # max points from the 11 auto-scored criteria
MANUAL_MAX = 12.5   # Business Clarity (2.5) + Long-Term Potential (10.0)

# Verdict thresholds scaled to the 87.5 auto-score range
# (proportional to the full 100-pt scale: 80→70, 65→56.9, 50→43.75, 35→30.6)
THRESHOLDS = [
    (70.0, "STRONG BUY", "#1B5E20", "white"),
    (56.9, "BUY",        "#43A047", "white"),
    (43.8, "HOLD",       "#FB8C00", "white"),
    (30.6, "REDUCE",     "#EF6C00", "white"),
    (0.0,  "SELL",       "#B71C1C", "white"),
]

def _verdict(score):
    try:
        s = float(score)
    except (TypeError, ValueError):
        return "N/A"
    for threshold, label, *_ in THRESHOLDS:
        if s >= threshold:
            return label
    return "SELL"

def _verdict_color(val):
    for _, label, bg, fg in THRESHOLDS:
        if val == label:
            return f"background-color:{bg};color:{fg};font-weight:bold"
    return ""

def _score_color(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v >= 70.0: return "background-color:#C8E6C9"
    if v >= 56.9: return "background-color:#BBDEFB"
    if v >= 43.8: return "background-color:#FFE0B2"
    if v >= 30.6: return "background-color:#FFCDD2"
    return "background-color:#EF9A9A"

def _roic_color(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v >= 0.20: return "color:#1B5E20;font-weight:bold"
    if v >= 0.12: return "color:#2E7D32"
    if v >= 0.08: return "color:#E65100"
    return "color:#B71C1C"

def _de_color(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v <= 1.0: return "color:#1B5E20;font-weight:bold"
    if v <= 2.5: return "color:#2E7D32"
    if v <= 4.0: return "color:#E65100"
    return "color:#B71C1C"

def _pe_color(val):
    """Green if current P/E is below 5yr avg, red if above."""
    return ""   # neutral — context needed; just display value

# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    r = requests.get(CSV_URL, timeout=8)
    if r.status_code != 200:
        return pd.DataFrame()
    return pd.read_csv(StringIO(r.text))

st.title("Portfolio Heatmap")
st.caption(
    f"**Auto Score** = points earned from 11 quantitative criteria out of **{AUTO_MAX}** max.  "
    f"Remaining **{MANUAL_MAX} pts** (Business Clarity 2.5 + Long-Term Potential 10.0) "
    f"are manual judgment — overlay them yourself to reach a 100-pt total."
)

df = load_data()

if df.empty:
    st.info("No model runs recorded yet. Generate your first model on the Home page.")
    st.stop()

# ── Clean & derive ────────────────────────────────────────────────────────────
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
# Keep latest run per ticker
df = (df.sort_values("Date", ascending=False)
        .drop_duplicates("Ticker")
        .sort_values("Auto_Score", ascending=False)
        .reset_index(drop=True))

df["Verdict"] = df["Auto_Score"].apply(_verdict)

# ── Verdict distribution chart ────────────────────────────────────────────────
st.subheader("Verdict Distribution")

verdict_order  = ["STRONG BUY", "BUY", "HOLD", "REDUCE", "SELL", "N/A"]
verdict_colors = {
    "STRONG BUY": "#1B5E20",
    "BUY":        "#43A047",
    "HOLD":       "#FB8C00",
    "REDUCE":     "#EF6C00",
    "SELL":       "#B71C1C",
    "N/A":        "#BDBDBD",
}

counts   = df["Verdict"].value_counts()
total    = len(df)
segments = [(v, counts.get(v, 0)) for v in verdict_order if counts.get(v, 0) > 0]

# Build horizontal stacked bar using st.columns proportionally
if segments:
    bar_cols = st.columns([c for _, c in segments])
    for col, (label, count) in zip(bar_cols, segments):
        pct = count / total * 100
        col.markdown(
            f"<div style='background:{verdict_colors[label]};border-radius:6px;"
            f"padding:10px 4px;text-align:center;color:white;font-weight:bold;"
            f"font-size:13px;'>"
            f"{label}<br><span style='font-size:18px'>{count}</span>"
            f"<br><span style='font-size:11px;opacity:0.85'>{pct:.0f}%</span></div>",
            unsafe_allow_html=True,
        )

st.markdown("---")

# ── Heatmap table ─────────────────────────────────────────────────────────────
st.subheader("Company Scores")

# Columns to display (only include those present in CSV)
DISPLAY_COLS = [
    "Ticker", "Auto_Score", "Verdict",
    "Price", "MktCap_B",
    "ROIC", "Rev_CAGR", "FCF_NI", "D_EBITDA",
    "PE_Current", "PE_5yr", "PFCF_Current", "PFCF_5yr",
    "Floor_Cap", "Date",
]
show = [c for c in DISPLAY_COLS if c in df.columns]
display = df[show].copy()

if "Date" in display.columns:
    display["Date"] = display["Date"].dt.strftime("%Y-%m-%d")

fmt = {}
if "Auto_Score"    in display.columns: fmt["Auto_Score"]    = "{:.1f}"
if "Price"         in display.columns: fmt["Price"]         = "${:.2f}"
if "MktCap_B"      in display.columns: fmt["MktCap_B"]      = "${:.1f}B"
if "ROIC"          in display.columns: fmt["ROIC"]          = "{:.1%}"
if "Rev_CAGR"      in display.columns: fmt["Rev_CAGR"]      = "{:.1%}"
if "FCF_NI"        in display.columns: fmt["FCF_NI"]        = "{:.1%}"
if "D_EBITDA"      in display.columns: fmt["D_EBITDA"]      = "{:.1f}x"
if "PE_Current"    in display.columns: fmt["PE_Current"]    = "{:.1f}x"
if "PE_5yr"        in display.columns: fmt["PE_5yr"]        = "{:.1f}x"
if "PFCF_Current"  in display.columns: fmt["PFCF_Current"]  = "{:.1f}x"
if "PFCF_5yr"      in display.columns: fmt["PFCF_5yr"]      = "{:.1f}x"

styled = display.style
if "Verdict"    in display.columns: styled = styled.map(_verdict_color, subset=["Verdict"])
if "Auto_Score" in display.columns: styled = styled.map(_score_color,   subset=["Auto_Score"])
if "ROIC"       in display.columns: styled = styled.map(_roic_color,    subset=["ROIC"])
if "D_EBITDA"   in display.columns: styled = styled.map(_de_color,      subset=["D_EBITDA"])
styled = styled.format(fmt, na_rep="—")

st.dataframe(styled, use_container_width=True, height=min(80 + 35 * len(display), 750))
st.caption(
    "Score thresholds (auto only):  "
    "≥70 STRONG BUY  |  ≥56.9 BUY  |  ≥43.8 HOLD  |  ≥30.6 REDUCE  |  <30.6 SELL  "
    "— proportional to full 100-pt thresholds (80/65/50/35).  "
    "Add up to 12.5 manual pts to reach final verdict."
)
