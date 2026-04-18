"""
Investment Automation — Streamlit Web App
Wraps fmp_3statementv6.py and serves the Excel model as a download.
"""

import base64
import datetime
import io
import sys
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Investment Model Generator",
    page_icon="📊",
    layout="centered",
)

# ── Password gate ─────────────────────────────────────────────────────────────
password = st.text_input("Password", type="password")
if not password:
    st.stop()
if password != st.secrets["APP_PASSWORD"]:
    st.error("Incorrect password.")
    st.stop()

# ── Inject API keys from Streamlit secrets into the module ───────────────────
import fmp_3statementv6 as mdl
mdl.API_KEY = st.secrets["FMP_API_KEY"]

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("📊 Investment Model Generator")
st.caption("Generates a full financial model Excel workbook from FMP data.")

col1, col2 = st.columns(2)
with col1:
    ticker = st.text_input("Ticker symbol", placeholder="e.g. AAPL, MSFT, NVDA").strip().upper()
with col2:
    manual_rating_raw = st.text_input(
        "S&P / Moody's credit rating (optional)",
        placeholder="e.g. AA+  or  Aa1"
    ).strip()

# Normalise manual rating
manual_rating = None
if manual_rating_raw:
    tok = manual_rating_raw.strip().split()[0].strip(".,;:()")
    manual_rating = mdl.MOODY_TO_SP.get(tok) or (
        tok.upper() if tok.upper() in mdl.VALID_SP_RATINGS else None
    )
    if not manual_rating:
        st.warning(f"'{manual_rating_raw}' not recognised as a valid rating — will be ignored.")

run = st.button("Generate Model", type="primary", disabled=not ticker)


def _write_outputs_row(ticker, metrics):
    """Append one row to outputs.csv in the GitHub repo."""
    token = st.secrets.get("GITHUB_TOKEN")
    repo  = st.secrets.get("GITHUB_REPO", "jaysang2908/Investment-Automation")
    branch = st.secrets.get("GITHUB_BRANCH", "main")
    if not token:
        st.warning("GITHUB_TOKEN secret not found — heatmap row not saved.")
        return
    try:
        import requests as _req
        api     = f"https://api.github.com/repos/{repo}/contents/outputs.csv"
        headers = {
            "Authorization": f"token {token}",
            "Accept":        "application/vnd.github.v3+json",
        }
        r = _req.get(api, headers=headers, params={"ref": branch}, timeout=8)
        if r.status_code == 200:
            info    = r.json()
            sha     = info["sha"]
            content = base64.b64decode(info["content"]).decode()
        else:
            sha     = None
            content = "Ticker,ROIC,Rev_CAGR,FCF_NI,D_EBITDA,Auto_Score,Floor_Cap,Date\n"

        def _f(v):
            return "" if v is None else f"{v:.4f}"

        row = ",".join([
            ticker,
            _f(metrics.get("roic")),
            _f(metrics.get("rev_cagr")),
            _f(metrics.get("fcf_ni")),
            _f(metrics.get("d_ebitda")),
            "" if metrics.get("auto_score") is None else str(metrics["auto_score"]),
            "" if metrics.get("floor_cap")  is None else str(metrics["floor_cap"]),
            datetime.date.today().isoformat(),
        ]) + "\n"
        content += row

        payload = {
            "message": f"Add {ticker} scorecard results",
            "branch":  branch,
            "content": base64.b64encode(content.encode()).decode(),
        }
        if sha:
            payload["sha"] = sha
        resp = _req.put(api, headers=headers, json=payload, timeout=10)
        if resp.status_code not in (200, 201):
            st.warning(f"Heatmap CSV write failed: {resp.status_code} — {resp.json().get('message','')}")
    except Exception as e:
        st.warning(f"Heatmap CSV write error: {e}")

if run and ticker:
    log = st.empty()
    messages = []

    def log_print(*args, **kwargs):
        """Capture print() output into the Streamlit log box."""
        msg = " ".join(str(a) for a in args)
        messages.append(msg)
        log.code("\n".join(messages))

    # Redirect stdout so existing print() calls in the module show up
    import builtins
    original_print = builtins.print
    builtins.print = log_print

    try:
        from openpyxl import Workbook

        log_print(f"Fetching data for {ticker}...")

        is_data = mdl.fetch("income-statement",        ticker)[:mdl.YEARS][::-1]
        bs_data = mdl.fetch("balance-sheet-statement", ticker)[:mdl.YEARS][::-1]
        cf_data = mdl.fetch("cash-flow-statement",     ticker)[:mdl.YEARS][::-1]

        if not is_data:
            st.error("No data returned — check the ticker symbol.")
            st.stop()

        # Fetch current price explicitly so DCF equity bridge always has it
        import requests as _req
        current_price = None
        try:
            _prof = _req.get(
                f"https://financialmodelingprep.com/stable/profile"
                f"?symbol={ticker}&apikey={mdl.API_KEY}", timeout=8
            ).json()
            _rec = _prof[0] if isinstance(_prof, list) else _prof
            current_price = float(_rec.get("price") or 0) or None
            log_print(f"  Current price: ${current_price}")
        except Exception:
            log_print("  Warning: could not fetch current price.")

        years = [
            d.get("fiscalYear") or d.get("calendarYear") or d["date"][:4]
            for d in is_data
        ]
        log_print(f"Years: {years}")

        wb = Workbook()
        mdl.build_cover(wb, ticker, years, is_data)
        pl_refs = mdl.build_pl(wb, is_data, years, ticker)
        bs_refs = mdl.build_bs(wb, bs_data, years, ticker)
        cf_refs = mdl.build_cf(wb, cf_data, years, ticker)
        mdl.build_ratios(wb, is_data, bs_data, cf_data, years, ticker, pl_refs, bs_refs, cf_refs)
        mdl.build_segments(wb, ticker, years)
        wacc_refs = mdl.build_wacc(wb, ticker, is_data, bs_data, manual_rating)
        mdl.build_dcf(wb, ticker, is_data, bs_data, cf_data, years, pl_refs, bs_refs, wacc_refs,
                      current_price=current_price)
        _, scorecard_metrics = mdl.build_scorecard(wb, ticker, is_data, bs_data, cf_data, years)

        # Save to memory buffer instead of disk
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        log_print("Done.")

        st.success("Model generated successfully.")
        _write_outputs_row(ticker, scorecard_metrics)
        st.download_button(
            label="⬇️ Download Excel Model",
            data=buf,
            file_name=f"{ticker}_FinancialModel_{years[-1]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.code(traceback.format_exc())

    finally:
        builtins.print = original_print
