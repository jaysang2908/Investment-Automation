"""
Portfolio Heatmap — reads outputs.csv from GitHub, renders heatmap + manual score overlay.
Must NOT call st.set_page_config() — only app.py does that.
"""

import streamlit as st
import pandas as pd
import re
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
    f"**Auto Score** (max **{AUTO_MAX}**) = 11 quantitative criteria.  "
    f"Add **Business Clarity** (0–2.5) + **LT Potential** (0–10) in the editor below — "
    f"when both are filled the **Full Score (/100)** column updates.  "
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
    MIN_W = 8
    raw_pcts = [c / total * 100 for _, c in present]
    n_small  = sum(1 for p in raw_pcts if p < MIN_W)
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

# ── Colour-coded main table (read-only, pandas Styler) ────────────────────────
st.subheader("Company Scores")
st.info(
    "**Tip:** Scroll to the far right of the table to see **Clarity (/2.5)** and **LT Potential (/10)** columns. "
    "Enter your manual scores in the **Edit Manual Scores** editor below and press **Save Manual Scores** — "
    "they will be reflected in the Total Score (/100) column."
)

# Columns shown in the main styled table — manual inputs at far right
MAIN_COLS = [
    "Ticker", "Verdict", "Total_Score", "Auto_Score",
    "Price", "MktCap_B",
    "GG_Price", "GG_Upside", "EM_Price", "EM_Upside",
    "PE_Current", "PE_5yr", "PFCF_Current", "PFCF_5yr",
    "ROIC", "Rev_CAGR", "FCF_NI",
    "D_EBITDA", "Revenue_B", "OCF_B", "FCF_B",
    "Floor_Cap", "Date",
    "Manual_Clarity", "Manual_LTP",
]
show_cols = [c for c in MAIN_COLS if c in df.columns]
disp = df[show_cols].copy()
if "Date" in disp.columns:
    disp["Date"] = df["Date"].dt.strftime("%Y-%m-%d")

# Rename columns for display
rename_map = {
    "Total_Score":    "Total Score (/100)",
    "Auto_Score":     "Auto Score (/87.5)",
    "Manual_Clarity": "Clarity (/2.5)",
    "Manual_LTP":     "LT Potential (/10)",
    "MktCap_B":       "Mkt Cap ($B)",
    "GG_Price":       "GG Price",
    "GG_Upside":      "GG Upside/(Down)",
    "EM_Price":       "EM Price",
    "EM_Upside":      "EM Upside/(Down)",
    "PE_Current":     "P/E",
    "PE_5yr":         "P/E 5yr avg",
    "PFCF_Current":   "P/FCF",
    "PFCF_5yr":       "P/FCF 5yr avg",
    "ROIC":           "ROIC",
    "Rev_CAGR":       "Rev CAGR",
    "FCF_NI":         "FCF/NPAT",
    "D_EBITDA":       "D/EBITDA",
    "Revenue_B":      "Revenue ($B)",
    "OCF_B":          "OCF ($B)",
    "FCF_B":          "FCF ($B)",
    "Floor_Cap":      "Floor Cap",
}
disp = disp.rename(columns=rename_map)

# Convert numeric columns
for col in ["Total Score (/100)", "Auto Score (/87.5)", "Clarity (/2.5)", "LT Potential (/10)",
            "Price", "Mkt Cap ($B)", "GG Price", "EM Price",
            "GG Upside/(Down)", "EM Upside/(Down)",
            "P/E", "P/E 5yr avg", "P/FCF", "P/FCF 5yr avg",
            "ROIC", "Rev CAGR", "FCF/NPAT", "D/EBITDA",
            "Revenue ($B)", "OCF ($B)", "FCF ($B)", "Floor Cap"]:
    if col in disp.columns:
        disp[col] = pd.to_numeric(disp[col], errors="coerce")


# ── Styling helpers ───────────────────────────────────────────────────────────
def _style_verdict(val):
    color = VERDICT_COLORS.get(str(val), "#BDBDBD")
    return f"background-color:{color};color:white;font-weight:bold;text-align:center"

def _style_score(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v != v:     return ""   # NaN guard (float('nan') != float('nan'))
    if v >= 70:   return "background-color:#1B5E20;color:white;font-weight:bold"
    if v >= 56.9: return "background-color:#43A047;color:white"
    if v >= 43.8: return "background-color:#FB8C00;color:white"
    if v >= 30.6: return "background-color:#EF6C00;color:white"
    return "background-color:#B71C1C;color:white"

def _style_roic(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v >= 0.20:  return "background-color:#1B5E20;color:white"
    if v >= 0.12:  return "background-color:#43A047;color:white"
    if v >= 0.08:  return "background-color:#FB8C00;color:white"
    return "background-color:#B71C1C;color:white"

def _style_upside(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v >= 0.30:  return "background-color:#1B5E20;color:white"
    if v >= 0.10:  return "background-color:#43A047;color:white"
    if v >= -0.10: return "background-color:#FB8C00;color:white"
    return "background-color:#B71C1C;color:white"

def _style_de(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if v <= 1.0:  return "background-color:#1B5E20;color:white"
    if v <= 2.5:  return "background-color:#43A047;color:white"
    if v <= 4.0:  return "background-color:#FB8C00;color:white"
    return "background-color:#B71C1C;color:white"

def _style_manual(val):
    return "background-color:#1565C0;color:white"


# Build format dict — only for columns that exist in disp
fmt = {}
for col, fmt_str in [
    ("Price",             "${:.2f}"),
    ("Mkt Cap ($B)",      "${:.1f}B"),
    ("GG Price",          "${:.2f}"),
    ("EM Price",          "${:.2f}"),
    ("GG Upside/(Down)",  "{:.1%}"),
    ("EM Upside/(Down)",  "{:.1%}"),
    ("P/E",               "{:.1f}x"),
    ("P/E 5yr avg",       "{:.1f}x"),
    ("P/FCF",             "{:.1f}x"),
    ("P/FCF 5yr avg",     "{:.1f}x"),
    ("ROIC",              "{:.1%}"),
    ("Rev CAGR",          "{:.1%}"),
    ("FCF/NPAT",          "{:.1%}"),
    ("D/EBITDA",          "{:.1f}x"),
    ("Revenue ($B)",      "${:.1f}B"),
    ("OCF ($B)",          "${:.1f}B"),
    ("FCF ($B)",          "${:.1f}B"),
    ("Total Score (/100)",  "{:.1f}"),
    ("Auto Score (/87.5)", "{:.1f}"),
    ("Clarity (/2.5)",      "{:.1f}"),
    ("LT Potential (/10)",  "{:.1f}"),
]:
    if col in disp.columns:
        fmt[col] = fmt_str

_score_cols = [c for c in ["Total Score (/100)", "Auto Score (/87.5)"] if c in disp.columns]
styled = (
    disp.style
    .format(fmt, na_rep="—")
    .map(_style_verdict,  subset=["Verdict"])
    .map(_style_score,    subset=_score_cols)
    .map(_style_roic,     subset=["ROIC"] if "ROIC" in disp.columns else [])
    .map(_style_upside,   subset=[c for c in ["GG Upside/(Down)", "EM Upside/(Down)"] if c in disp.columns])
    .map(_style_de,       subset=["D/EBITDA"] if "D/EBITDA" in disp.columns else [])
    .map(_style_manual,   subset=[c for c in ["Clarity (/2.5)", "LT Potential (/10)"] if c in disp.columns])
    .set_properties(**{"white-space": "nowrap"})
    .set_table_styles([
        {"selector": "th", "props": [
            ("background-color", "#1a1a1a"),
            ("color", "white"),
            ("font-weight", "bold"),
            ("padding", "8px 10px"),
            ("white-space", "nowrap"),
        ]},
        {"selector": "th.col_heading", "props": [
            ("background-color", "#1a1a1a"),
            ("color", "white"),
            ("font-weight", "bold"),
        ]},
        {"selector": "th.blank", "props": [
            ("background-color", "#1a1a1a"),
        ]},
    ])
)

_table_html = styled.to_html(escape=False)

# Strip any existing style attr from <th>, then stamp our own inline style.
# Inline styles have the highest CSS specificity — nothing in any <style> block can override them.
_TH_STYLE = (
    "color:#ffffff;"
    "font-size:15px;"
    "font-weight:800;"
    "background-color:#1a1a1a;"
    "padding:10px 14px;"
    "white-space:nowrap;"
    "border-bottom:2px solid #444;"
    "letter-spacing:0.04em;"
)
_table_html = re.sub(r'(<th\b[^>]*?)\s+style="[^"]*"', r'\1', _table_html)
_table_html = re.sub(r'(<th\b)([^>]*?>)', rf'\1 style="{_TH_STYLE}"\2', _table_html)

st.markdown(
    f'<div style="overflow-x:auto;overflow-y:auto;max-height:750px;">{_table_html}</div>',
    unsafe_allow_html=True,
)

st.markdown("---")

# ── Manual score editor ───────────────────────────────────────────────────────
st.subheader("Edit Manual Scores")
st.caption(
    "**Clarity** (0–2.5 pts): business model clarity.  "
    "**LT Potential** (0–10 pts): long-term upside potential.  "
    "Press **Save Manual Scores** to persist — scores refresh within 5 minutes."
)

editor_df = df[["Ticker", "Auto_Score", "Manual_Clarity", "Manual_LTP", "Total_Score"]].copy()

edited = st.data_editor(
    editor_df,
    column_config={
        "Ticker":         st.column_config.TextColumn("Ticker",              disabled=True),
        "Auto_Score":     st.column_config.NumberColumn("Auto Score (/87.5)", disabled=True, format="%.1f"),
        "Manual_Clarity": st.column_config.NumberColumn("Clarity (/2.5)",
                              min_value=0.0, max_value=2.5, step=0.5, format="%.1f"),
        "Manual_LTP":     st.column_config.NumberColumn("LT Potential (/10)",
                              min_value=0.0, max_value=10.0, step=1.0, format="%.1f"),
        "Total_Score":    st.column_config.NumberColumn("Total Score (/100)", disabled=True, format="%.1f"),
    },
    use_container_width=True,
    hide_index=True,
    height=min(80 + 35 * len(editor_df), 600),
    key="manual_editor",
)

if st.button("Save Manual Scores", type="primary"):
    for col in ["Manual_Clarity", "Manual_LTP"]:
        if col in edited.columns:
            df.loc[df["Ticker"].isin(edited["Ticker"]), col] = (
                edited.set_index("Ticker")[col].values
            )
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
