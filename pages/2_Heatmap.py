"""
Portfolio Heatmap — reads outputs.csv from GitHub and renders a colour-coded table.
Must NOT call st.set_page_config() — only app.py does that.
"""

import streamlit as st
import pandas as pd
import requests
from io import StringIO

# ── Password gate (same as home page) ────────────────────────────────────────
password = st.text_input("Password", type="password", key="heatmap_pw")
if not password:
    st.stop()
if password != st.secrets["APP_PASSWORD"]:
    st.error("Incorrect password.")
    st.stop()

# ── Config ────────────────────────────────────────────────────────────────────
REPO    = st.secrets.get("GITHUB_REPO", "jaysang2908/Investment-Automation")
CSV_URL = f"https://raw.githubusercontent.com/{REPO}/main/outputs.csv"

# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    r = requests.get(CSV_URL, timeout=8)
    if r.status_code != 200:
        return pd.DataFrame()
    return pd.read_csv(StringIO(r.text))

st.title("Portfolio Heatmap")
st.caption("Auto-score = weighted avg of 11 quantitative criteria (0–100 scale). "
           "Refresh every 5 min. Two manual criteria (Business Clarity, Long-Term Potential) "
           "are excluded — full scorecard is in the downloaded Excel.")

df = load_data()

if df.empty:
    st.info("No model runs recorded yet. Generate your first model on the Home page.")
    st.stop()

# ── Derived columns ───────────────────────────────────────────────────────────
def _verdict(score):
    try:
        s = float(score)
    except (TypeError, ValueError):
        return "N/A"
    if s >= 80: return "STRONG BUY"
    if s >= 65: return "BUY"
    if s >= 50: return "HOLD"
    if s >= 35: return "REDUCE"
    return "SELL"

df["Verdict"] = df["Auto_Score"].apply(_verdict)

# Keep latest run per ticker (in case re-run)
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df = df.sort_values("Date", ascending=False).drop_duplicates("Ticker").sort_values("Auto_Score", ascending=False)

# ── Styling helpers ───────────────────────────────────────────────────────────
_VERDICT_STYLE = {
    "STRONG BUY": "background-color:#1B5E20;color:white;font-weight:bold",
    "BUY":        "background-color:#43A047;color:white;font-weight:bold",
    "HOLD":       "background-color:#FB8C00;color:white;font-weight:bold",
    "REDUCE":     "background-color:#EF6C00;color:white;font-weight:bold",
    "SELL":       "background-color:#B71C1C;color:white;font-weight:bold",
    "N/A":        "",
}

def _style_verdict(val):
    return _VERDICT_STYLE.get(val, "")

def _style_score(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v >= 80: return "background-color:#C8E6C9"
    if v >= 65: return "background-color:#BBDEFB"
    if v >= 50: return "background-color:#FFE0B2"
    if v >= 35: return "background-color:#FFCDD2"
    return "background-color:#EF9A9A"

def _style_roic(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v >= 0.20: return "color:#1B5E20;font-weight:bold"
    if v >= 0.12: return "color:#2E7D32"
    if v >= 0.08: return "color:#E65100"
    return "color:#B71C1C"

def _style_de(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v <= 1.0: return "color:#1B5E20;font-weight:bold"
    if v <= 2.5: return "color:#2E7D32"
    if v <= 4.0: return "color:#E65100"
    return "color:#B71C1C"

# ── Build display frame ───────────────────────────────────────────────────────
COLS = ["Ticker", "Auto_Score", "Verdict", "ROIC", "Rev_CAGR", "FCF_NI", "D_EBITDA", "Floor_Cap", "Date"]
show = [c for c in COLS if c in df.columns]
display = df[show].copy()

# Format date back to string for display
if "Date" in display.columns:
    display["Date"] = display["Date"].dt.strftime("%Y-%m-%d")

fmt = {}
if "ROIC"      in display.columns: fmt["ROIC"]      = "{:.1%}"
if "Rev_CAGR"  in display.columns: fmt["Rev_CAGR"]  = "{:.1%}"
if "FCF_NI"    in display.columns: fmt["FCF_NI"]    = "{:.1%}"
if "D_EBITDA"  in display.columns: fmt["D_EBITDA"]  = "{:.1f}x"
if "Auto_Score" in display.columns: fmt["Auto_Score"] = "{:.1f}"

styled = display.style

if "Verdict"    in display.columns: styled = styled.map(_style_verdict,  subset=["Verdict"])
if "Auto_Score" in display.columns: styled = styled.map(_style_score,    subset=["Auto_Score"])
if "ROIC"       in display.columns: styled = styled.map(_style_roic,     subset=["ROIC"])
if "D_EBITDA"   in display.columns: styled = styled.map(_style_de,       subset=["D_EBITDA"])

styled = styled.format(fmt, na_rep="N/A")

st.dataframe(styled, use_container_width=True, height=min(80 + 35 * len(display), 700))

# ── Summary counts ────────────────────────────────────────────────────────────
counts = df["Verdict"].value_counts()
order  = ["STRONG BUY", "BUY", "HOLD", "REDUCE", "SELL", "N/A"]
cols   = st.columns(len(order))
for col, label in zip(cols, order):
    col.metric(label, counts.get(label, 0))
