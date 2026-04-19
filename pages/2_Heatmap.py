"""
Portfolio Heatmap — reads outputs.csv from GitHub, renders heatmap + manual score overlay.
Must NOT call st.set_page_config() — only app.py does that.
"""

import streamlit as st
import pandas as pd
import requests
from io import StringIO
import base64
import datetime

# ── Password gate ─────────────────────────────────────────────────────────────
password = st.text_input("Password", type="password", key="heatmap_pw")
if not password:
    st.stop()
if password != st.secrets["APP_PASSWORD"]:
    st.error("Incorrect password.")
    st.stop()

# ── Config ────────────────────────────────────────────────────────────────────
REPO      = st.secrets.get("GITHUB_REPO", "jaysang2908/Investment-Automation")
BRANCH    = st.secrets.get("GITHUB_BRANCH", "main")
CSV_URL   = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}/outputs.csv"
AUTO_MAX  = 87.5
MANUAL_MAX = 12.5

THRESHOLDS = [
    (70.0, "STRONG BUY", "#1B5E20"),
    (56.9, "BUY",        "#43A047"),
    (43.8, "HOLD",       "#FB8C00"),
    (30.6, "REDUCE",     "#EF6C00"),
    (0.0,  "SELL",       "#B71C1C"),
]

def _verdict(score):
    try:
        s = float(score)
    except (TypeError, ValueError):
        return "N/A"
    for threshold, label, _ in THRESHOLDS:
        if s >= threshold:
            return label
    return "SELL"

VERDICT_COLORS = {l: c for _, l, c in THRESHOLDS}
VERDICT_COLORS["N/A"] = "#BDBDBD"

# ── Load / save CSV ───────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    r = requests.get(CSV_URL, timeout=8)
    if r.status_code != 200:
        return pd.DataFrame()
    return pd.read_csv(StringIO(r.text), on_bad_lines="skip")

def save_data(df):
    token = st.secrets.get("GITHUB_TOKEN")
    if not token:
        st.warning("GITHUB_TOKEN not set — cannot save.")
        return False
    try:
        api = f"https://api.github.com/repos/{REPO}/contents/outputs.csv"
        headers = {"Authorization": f"token {token}",
                   "Accept": "application/vnd.github.v3+json"}
        r = requests.get(api, headers=headers, params={"ref": BRANCH}, timeout=8)
        sha = r.json().get("sha") if r.status_code == 200 else None
        csv_str = df.to_csv(index=False)
        payload = {"message": "Update manual scores from heatmap",
                   "branch": BRANCH,
                   "content": base64.b64encode(csv_str.encode()).decode()}
        if sha:
            payload["sha"] = sha
        resp = requests.put(api, headers=headers, json=payload, timeout=10)
        return resp.status_code in (200, 201)
    except Exception as e:
        st.warning(f"Save error: {e}")
        return False

# ── Page ──────────────────────────────────────────────────────────────────────
st.title("Portfolio Heatmap")
st.caption(
    f"**Auto Score** = points from 11 quantitative criteria, max **{AUTO_MAX}**.  "
    f"Add up to **{MANUAL_MAX} manual pts** (Business Clarity 2.5 + Long-Term Potential 10.0) "
    f"in the table below to reach a 100-pt total.  "
    f"GG / EM prices computed from Python-side DCF (mirrors Excel defaults: g=3%, exit 20×)."
)

df = load_data()
if df.empty:
    st.info("No model runs recorded yet. Generate your first model on the Home page.")
    st.stop()

# ── Ensure manual columns exist ───────────────────────────────────────────────
for col, default in [("Manual_Clarity", 0.0), ("Manual_LTP", 0.0)]:
    if col not in df.columns:
        df[col] = default
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

# ── Clean & deduplicate ───────────────────────────────────────────────────────
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df = (df.sort_values("Date", ascending=False)
        .drop_duplicates("Ticker")
        .reset_index(drop=True))

df["Auto_Score"]  = pd.to_numeric(df.get("Auto_Score"),  errors="coerce")
df["Total_Score"] = (df["Auto_Score"]
                     .add(df["Manual_Clarity"], fill_value=0)
                     .add(df["Manual_LTP"],     fill_value=0)
                     .round(1))
df["Verdict"]     = df["Total_Score"].apply(_verdict)
df = df.sort_values("Total_Score", ascending=False).reset_index(drop=True)

# ── Verdict distribution bar ──────────────────────────────────────────────────
st.subheader("Verdict Distribution")
counts  = df["Verdict"].value_counts()
order   = ["STRONG BUY", "BUY", "HOLD", "REDUCE", "SELL", "N/A"]
present = [(v, counts.get(v, 0)) for v in order if counts.get(v, 0) > 0]
total   = len(df)

if present:
    # HTML flex bar — minimum 8% width per segment so labels always fit
    MIN_W = 8
    raw_pcts = [c / total * 100 for _, c in present]
    # Scale so minimum segments get MIN_W, rest fill proportionally
    n_small = sum(1 for p in raw_pcts if p < MIN_W)
    reserved = n_small * MIN_W
    remaining = 100 - reserved
    big_total = sum(p for p in raw_pcts if p >= MIN_W) or 1
    final_pcts = [
        MIN_W if p < MIN_W else p / big_total * remaining
        for p in raw_pcts
    ]
    segments_html = "".join(
        f"<div style='width:{w:.1f}%;background:{VERDICT_COLORS[v]};display:flex;"
        f"flex-direction:column;align-items:center;justify-content:center;"
        f"padding:8px 2px;border-radius:4px;margin:1px;'>"
        f"<span style='color:white;font-weight:bold;font-size:12px;white-space:nowrap'>{v}</span>"
        f"<span style='color:white;font-size:18px;font-weight:bold'>{c}</span>"
        f"<span style='color:rgba(255,255,255,0.85);font-size:11px'>{c/total*100:.0f}%</span>"
        f"</div>"
        for (v, c), w in zip(present, final_pcts)
    )
    st.markdown(
        f"<div style='display:flex;width:100%;gap:2px;margin-bottom:16px'>{segments_html}</div>",
        unsafe_allow_html=True,
    )

st.markdown("---")

# ── Manual score overlay ──────────────────────────────────────────────────────
st.subheader("Company Scores")
st.caption(
    "Edit **Clarity** (0–2.5) and **LT Potential** (0–10) directly in the table. "
    "Total Score and Verdict update live. Press **Save Manual Scores** to persist."
)

# Columns to show in the editor
NUM_COLS = {
    "Total_Score": "Total Score",
    "Auto_Score":  "Auto Score",
    "Verdict":     "Verdict",
    "Manual_Clarity": "Clarity",
    "Manual_LTP":  "LT Potential",
}

DISPLAY_COLS = ["Ticker", "Total_Score", "Verdict", "Auto_Score",
                "Manual_Clarity", "Manual_LTP",
                "Price", "MktCap_B",
                "GG_Price", "GG_Upside", "EM_Price", "EM_Upside",
                "PE_Current", "PE_5yr", "PFCF_Current", "PFCF_5yr",
                "ROIC", "Rev_CAGR", "FCF_NI", "D_EBITDA",
                "Floor_Cap", "Date"]
show = [c for c in DISPLAY_COLS if c in df.columns]
display = df[show].copy()
if "Date" in display.columns:
    display["Date"] = display["Date"].dt.strftime("%Y-%m-%d")

# Configure editable columns
col_config = {
    "Ticker":        st.column_config.TextColumn("Ticker",       disabled=True),
    "Total_Score":   st.column_config.NumberColumn("Total Score", disabled=True, format="%.1f"),
    "Auto_Score":    st.column_config.NumberColumn("Auto Score",  disabled=True, format="%.1f"),
    "Verdict":       st.column_config.TextColumn("Verdict",       disabled=True),
    "Manual_Clarity":st.column_config.NumberColumn("Clarity",     min_value=0.0, max_value=2.5,  step=0.5,  format="%.1f"),
    "Manual_LTP":    st.column_config.NumberColumn("LT Potential",min_value=0.0, max_value=10.0, step=1.0,  format="%.1f"),
    "Price":         st.column_config.NumberColumn("Price",        disabled=True, format="$%.2f"),
    "MktCap_B":      st.column_config.NumberColumn("Mkt Cap ($B)", disabled=True, format="$%.1f"),
    "GG_Price":      st.column_config.NumberColumn("GG Price",     disabled=True, format="$%.2f"),
    "GG_Upside":     st.column_config.NumberColumn("GG Upside",    disabled=True, format="{:.1%}"),
    "EM_Price":      st.column_config.NumberColumn("EM Price",     disabled=True, format="$%.2f"),
    "EM_Upside":     st.column_config.NumberColumn("EM Upside",    disabled=True, format="{:.1%}"),
    "PE_Current":    st.column_config.NumberColumn("P/E",          disabled=True, format="%.1fx"),
    "PE_5yr":        st.column_config.NumberColumn("P/E 5yr avg",  disabled=True, format="%.1fx"),
    "PFCF_Current":  st.column_config.NumberColumn("P/FCF",        disabled=True, format="%.1fx"),
    "PFCF_5yr":      st.column_config.NumberColumn("P/FCF 5yr avg",disabled=True, format="%.1fx"),
    "ROIC":          st.column_config.NumberColumn("ROIC",          disabled=True, format="{:.1%}"),
    "Rev_CAGR":      st.column_config.NumberColumn("Rev CAGR",     disabled=True, format="{:.1%}"),
    "FCF_NI":        st.column_config.NumberColumn("FCF/NI",       disabled=True, format="{:.1%}"),
    "D_EBITDA":      st.column_config.NumberColumn("D/EBITDA",     disabled=True, format="%.1fx"),
    "Floor_Cap":     st.column_config.NumberColumn("Floor Cap",    disabled=True),
    "Date":          st.column_config.TextColumn("Date",           disabled=True),
}
col_config = {k: v for k, v in col_config.items() if k in show}

edited = st.data_editor(
    display,
    column_config=col_config,
    use_container_width=True,
    hide_index=True,
    height=min(80 + 35 * len(display), 750),
    key="heatmap_editor",
)

if st.button("💾 Save Manual Scores", type="primary"):
    # Merge edits back into full df
    for col in ["Manual_Clarity", "Manual_LTP"]:
        if col in edited.columns:
            df.loc[df["Ticker"].isin(edited["Ticker"]), col] = (
                edited.set_index("Ticker")[col]
            )
    # Recompute totals
    df["Total_Score"] = (df["Auto_Score"]
                         .add(df["Manual_Clarity"], fill_value=0)
                         .add(df["Manual_LTP"],     fill_value=0)
                         .round(1))
    cols_to_save = [c for c in df.columns if c != "Verdict"]
    if save_data(df[cols_to_save]):
        st.success("Saved. Scores will refresh within 5 minutes.")
        st.cache_data.clear()
    else:
        st.error("Save failed — check GITHUB_TOKEN.")

# ── Legend ────────────────────────────────────────────────────────────────────
st.caption(
    "Score thresholds (0–100 total):  "
    "≥70 STRONG BUY  |  ≥56.9 BUY  |  ≥43.8 HOLD  |  ≥30.6 REDUCE  |  <30.6 SELL  "
    "— proportional to classic 80/65/50/35 thresholds, scaled to auto-only max of 87.5.  "
    "GG/EM blanks = model not yet regenerated through app."
)
