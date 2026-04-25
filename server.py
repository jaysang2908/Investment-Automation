"""
server.py — Flask backend for the Investment Research web app.
Wraps fmp_3statementv6.py (Excel) + report_bridge.py (HTML report).

Endpoints:
  GET  /                      → serves static/index.html
  POST /generate              → runs model, returns report_id + metrics
  GET  /report/<rid>          → view HTML report in browser
  GET  /download/excel/<rid>  → download Excel workbook
  GET  /download/html/<rid>   → download HTML report file
"""

import io
import os
import uuid
import builtins
import datetime
import traceback

from flask import Flask, request, jsonify, Response
import requests as _req

import fmp_3statementv6 as mdl
from report_bridge import build_report_data, render_html_report
from data_store import save_ticker_data, load_ticker_data

# ── App setup ─────────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder="static", static_url_path="")

# ── Config from environment ───────────────────────────────────────────────────
mdl.API_KEY  = os.environ.get("FMP_API_KEY", mdl.API_KEY)
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

# ── In-memory report store (2-hour TTL) ───────────────────────────────────────
_reports: dict = {}

def _prune():
    cutoff = datetime.datetime.now() - datetime.timedelta(hours=2)
    for rid in [k for k, v in _reports.items() if v["ts"] < cutoff]:
        del _reports[rid]


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return app.send_static_file("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    _prune()

    body         = request.get_json(force=True) or {}
    ticker       = body.get("ticker", "").strip().upper()
    rating_raw   = body.get("rating", "").strip()
    password     = body.get("password", "").strip()
    biz_clarity  = body.get("biz_clarity", "").strip().upper()
    ltp          = body.get("ltp", "").strip().upper()

    if APP_PASSWORD and password != APP_PASSWORD:
        return jsonify({"error": "Incorrect password."}), 401
    if not ticker:
        return jsonify({"error": "Ticker required."}), 400

    # Normalise credit rating (same logic as app.py)
    manual_rating = None
    if rating_raw:
        tok = rating_raw.strip().split()[0].strip(".,;:()")
        manual_rating = mdl.MOODY_TO_SP.get(tok) or (
            tok.upper() if tok.upper() in mdl.VALID_SP_RATINGS else None
        )

    # Capture print() output for live log
    logs = []
    _orig_print = builtins.print
    builtins.print = lambda *a, **k: logs.append(" ".join(str(x) for x in a))

    try:
        from openpyxl import Workbook

        # ── Fetch financials ──────────────────────────────────────────────────
        is_data = mdl.fetch("income-statement",        ticker)[:mdl.YEARS][::-1]
        bs_data = mdl.fetch("balance-sheet-statement", ticker)[:mdl.YEARS][::-1]
        cf_data = mdl.fetch("cash-flow-statement",     ticker)[:mdl.YEARS][::-1]

        if not is_data:
            return jsonify({"error": f"No financial data returned for {ticker}."}), 400

        # ── Company profile (price, mktCap, name, …) ─────────────────────────
        profile       = {}
        current_price = None
        market_cap    = None
        try:
            _p = _req.get(
                f"https://financialmodelingprep.com/stable/profile"
                f"?symbol={ticker}&apikey={mdl.API_KEY}", timeout=8
            ).json()
            profile       = (_p[0] if isinstance(_p, list) and _p else _p or {})
            current_price = float(profile.get("price") or 0) or None
            market_cap    = float(profile.get("mktCap") or profile.get("marketCap") or 0) or None
        except Exception:
            pass

        years = [
            d.get("fiscalYear") or d.get("calendarYear") or d["date"][:4]
            for d in is_data
        ]

        # ── Build Excel workbook ──────────────────────────────────────────────
        wb       = Workbook()
        pl_refs  = mdl.build_pl(wb, is_data, years, ticker)
        mdl.build_cover(wb, ticker, years, is_data)
        bs_refs  = mdl.build_bs(wb, bs_data, years, ticker)
        cf_refs  = mdl.build_cf(wb, cf_data, years, ticker)
        mdl.build_ratios(wb, is_data, bs_data, cf_data, years, ticker, pl_refs, bs_refs, cf_refs)
        mdl.build_segments(wb, ticker, years)
        wacc_refs = mdl.build_wacc(wb, ticker, is_data, bs_data, manual_rating)
        dcf_refs  = mdl.build_dcf(
            wb, ticker, is_data, bs_data, cf_data, years,
            pl_refs, bs_refs, wacc_refs, current_price=current_price
        )
        _, scorecard_metrics = mdl.build_scorecard(wb, ticker, is_data, bs_data, cf_data, years)

        buf = io.BytesIO()
        wb.save(buf)
        excel_bytes = buf.getvalue()

        # ── Compute adjusted score first (need it for HTML report) ───────────
        TIER_PTS    = {"HIGH": 10, "MOD-HIGH": 7, "MOD-LOW": 3, "LOW": 0}
        auto_score  = scorecard_metrics.get("auto_score") or 0
        bc_pts      = TIER_PTS.get(biz_clarity, 0) * 2.5 / 10   # max 2.5
        ltp_pts     = TIER_PTS.get(ltp, 0) * 10.0 / 10           # max 10.0
        adj_score   = round(auto_score + bc_pts + ltp_pts, 1)
        floor_cap   = scorecard_metrics.get("floor_cap")
        if floor_cap is not None:
            adj_score = min(adj_score, floor_cap)

        # ── Build HTML report ─────────────────────────────────────────────────
        report_data = build_report_data(
            ticker            = ticker,
            profile           = profile,
            is_data           = is_data,
            bs_data           = bs_data,
            cf_data           = cf_data,
            years             = years,
            wacc_val          = wacc_refs.get("wacc_val"),
            dcf_prices        = (dcf_refs or {}).get("dcf_prices") or {},
            scorecard_metrics = scorecard_metrics,
            manual_rating     = manual_rating,
            current_price     = current_price,
            market_cap        = market_cap,
            biz_clarity       = biz_clarity or None,
            ltp               = ltp or None,
            adj_score         = adj_score,
        )
        html_content = render_html_report(report_data)

        # ── Persist raw data for DCF calculator (avoids future FMP calls) ───
        try:
            save_ticker_data(
                ticker=ticker, is_data=is_data, bs_data=bs_data, cf_data=cf_data,
                profile=profile, years=years,
                wacc_val=wacc_refs.get("wacc_val"),
                dcf_prices=(dcf_refs or {}).get("dcf_prices") or {},
                scorecard_metrics=scorecard_metrics,
                analyst_ests=None,
            )
        except Exception:
            pass

        # ── Store with short ID ───────────────────────────────────────────────
        rid = uuid.uuid4().hex[:10]
        _reports[rid] = {
            "ticker":    ticker,
            "html":      html_content,
            "excel":     excel_bytes,
            "year":      years[-1] if years else "2025",
            "ts":        datetime.datetime.now(),
            "score":     adj_score,
            "auto_score": auto_score,
        }

        return jsonify({
            "report_id":   rid,
            "ticker":      ticker,
            "auto_score":  auto_score,
            "adj_score":   adj_score,
            "biz_clarity": biz_clarity or None,
            "ltp":         ltp or None,
            "logs":        logs[-15:],
        })

    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500
    finally:
        builtins.print = _orig_print


@app.route("/report/<rid>")
def view_report(rid):
    r = _reports.get(rid)
    if not r:
        return "<h2 style='font-family:sans-serif;padding:40px'>Report not found or expired (2-hour TTL).</h2>", 404
    return Response(r["html"], mimetype="text/html")


@app.route("/download/excel/<rid>")
def download_excel(rid):
    r = _reports.get(rid)
    if not r:
        return "Not found", 404
    fname = f"{r['ticker']}_FinancialModel_{r['year']}.xlsx"
    return Response(
        r["excel"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


@app.route("/download/html/<rid>")
def download_html(rid):
    r = _reports.get(rid)
    if not r:
        return "Not found", 404
    fname = f"{r['ticker']}_Report.html"
    return Response(
        r["html"],
        mimetype="text/html",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


@app.route("/dcf")
def dcf_page():
    return app.send_static_file("dcf.html")


@app.route("/api/dcf-data/<ticker>")
def dcf_data(ticker):
    ticker = ticker.upper().strip()
    try:
        # ── Try stored data first — zero FMP calls for known tickers ─────────
        stored = load_ticker_data(ticker)
        if stored:
            return jsonify(_build_dcf_response(stored))

        # ── Unknown ticker: fetch from FMP (costs quota) ──────────────────────
        is_data = mdl.fetch("income-statement",        ticker)[:5][::-1]
        bs_data = mdl.fetch("balance-sheet-statement", ticker)[:5][::-1]
        cf_data = mdl.fetch("cash-flow-statement",     ticker)[:5][::-1]
        if not is_data:
            return jsonify({"error": f"No financial data for {ticker}."}), 404

        profile = {}
        try:
            _p = _req.get(
                f"https://financialmodelingprep.com/stable/profile"
                f"?symbol={ticker}&apikey={mdl.API_KEY}", timeout=8
            ).json()
            profile = _p[0] if isinstance(_p, list) and _p else _p or {}
        except Exception:
            pass

        years_hist = [
            d.get("fiscalYear") or d.get("calendarYear") or d["date"][:4]
            for d in is_data
        ]
        analyst_ests = []
        try:
            _ae = _req.get(
                f"https://financialmodelingprep.com/stable/analyst-estimates"
                f"?symbol={ticker}&period=annual&limit=5&apikey={mdl.API_KEY}", timeout=8
            ).json()
            if isinstance(_ae, list):
                analyst_ests = sorted(
                    [e for e in _ae if str(e.get("date",""))[:4] > str(years_hist[-1])],
                    key=lambda x: x.get("date","")
                )[:5]
        except Exception:
            pass

        # Save so future calls are free
        save_ticker_data(ticker, is_data, bs_data, cf_data, profile, years_hist,
                         None, {}, {}, analyst_ests)

        stored = load_ticker_data(ticker)
        return jsonify(_build_dcf_response(stored))

    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


def _build_dcf_response(stored):
    """Build the DCF API response from a stored data dict (no FMP calls)."""
    is_data  = stored["is_data"]
    bs_data  = stored["bs_data"]
    cf_data  = stored["cf_data"]
    profile  = stored.get("profile") or {}
    years    = stored.get("years") or []
    wacc_val = stored.get("wacc_val")
    dcf_px   = stored.get("dcf_prices") or {}
    ae_raw   = stored.get("analyst_ests") or []

    ticker        = stored["ticker"]
    current_price = float(profile.get("price") or 0) or None
    market_cap    = float(profile.get("mktCap") or 0) or None
    shares        = float(profile.get("sharesOutstanding") or 0)
    beta          = float(profile.get("beta") or 1.0) or 1.0

    # Build history rows
    history = []
    for is_, bs_, cf_ in zip(is_data, bs_data, cf_data):
        rev        = is_.get("revenue") or 0
        da         = abs(is_.get("depreciationAndAmortization") or
                         cf_.get("depreciationAndAmortization") or 0)
        ebit       = is_.get("operatingIncome") or 0
        ebitda_raw = is_.get("ebitda") or (ebit + da)
        capex      = abs(cf_.get("capitalExpenditure") or 0)
        ocf        = cf_.get("operatingCashFlow") or 0
        fcf        = cf_.get("freeCashFlow") or (ocf - capex)
        pti        = is_.get("incomeBeforeTax") or is_.get("pretaxIncome") or 0
        te         = abs(is_.get("incomeTaxExpense") or 0)
        tax_r      = min(te / pti, 0.50) if pti > 0 else 0.21
        history.append({
            "year":          is_.get("fiscalYear") or is_.get("calendarYear") or is_["date"][:4],
            "revenue_m":     round(rev / 1e6, 1),
            "rev_growth":    None,
            "ebitda_margin": round(ebitda_raw / rev, 4) if rev else 0,
            "da_pct":        round(da / rev, 4)     if rev else 0,
            "capex_pct":     round(capex / rev, 4)  if rev else 0,
            "tax_rate":      round(tax_r, 4),
            "fcf_m":         round(fcf / 1e6, 1),
        })
    for i in range(1, len(history)):
        prev = history[i-1]["revenue_m"]
        curr = history[i]["revenue_m"]
        history[i]["rev_growth"] = round(curr / prev - 1, 4) if prev else None

    # Net debt
    bs0      = bs_data[-1]
    cash     = bs0.get("cashAndCashEquivalents") or 0
    debt     = (bs0.get("shortTermDebt") or 0) + (bs0.get("longTermDebt") or 0)
    net_debt = debt - cash

    # WACC components
    RF  = 0.043
    ERP = 0.045
    ke  = RF + beta * ERP
    cap_total = (market_cap + debt) if market_cap else max(debt, 1)
    ew  = market_cap / cap_total if market_cap else 1.0
    dw  = debt / cap_total if debt else 0.0
    is0     = is_data[-1]
    pti0    = is0.get("incomeBeforeTax") or is0.get("pretaxIncome") or 0
    te0     = abs(is0.get("incomeTaxExpense") or 0)
    eff_tax = min(te0 / pti0, 0.50) if pti0 > 0 else 0.21
    int_exp = abs(is0.get("interestExpense") or 0)
    kd_pre  = max(0.02, min(int_exp / debt if debt > 0 else RF * 0.8, 0.15))
    wacc    = wacc_val or (ew * ke + dw * kd_pre * (1 - eff_tax))

    # Default projections — anchor to the Excel-engine WACC and DCF prices
    last_g  = history[-1]["rev_growth"] or 0.05
    last_m  = history[-1]["ebitda_margin"] or 0.20
    last_da = history[-1]["da_pct"] or 0.03
    last_cx = history[-1]["capex_pct"] or 0.04

    analyst_ests = ae_raw
    def_rev = []
    for i in range(5):
        if i < len(analyst_ests):
            er  = analyst_ests[i].get("estimatedRevenueAvg") or 0
            epr = (analyst_ests[i-1].get("estimatedRevenueAvg")
                   if i > 0 else history[-1]["revenue_m"] * 1e6) or 0
            if er and epr:
                def_rev.append(round(er / epr - 1, 4))
                continue
        tgt = 0.05
        g   = last_g + (tgt - last_g) * (i / 4) if last_g != tgt else tgt
        def_rev.append(round(max(-0.10, min(g, 0.60)), 4))

    defaults = {
        "rev_growth":    def_rev,
        "ebitda_margin": [round(last_m, 4)]  * 5,
        "da_pct":        [round(last_da, 4)] * 5,
        "capex_pct":     [round(last_cx, 4)] * 5,
        "nwc_pct":       [0.005] * 5,
        "tax_rate":      [round(eff_tax, 4)] * 5,
        "tgr":           0.030,
        "exit_multiple": 15.0,
        "rf":            RF,
        "beta":          round(beta, 3),
        "erp":           ERP,
        "kd_pretax":     round(kd_pre, 4),
        "eff_tax":       round(eff_tax, 4),
        "equity_weight": round(ew, 4),
        "wacc":          round(wacc, 4),
    }

    last_year = int(years[-1]) if years else 2024

    return {
        "ticker":         ticker,
        "name":           profile.get("companyName") or ticker,
        "price":          current_price,
        "shares_m":       round(shares / 1e6, 2),
        "net_debt_m":     round(net_debt / 1e6, 1),
        "last_year":      last_year,
        "history":        history,
        "defaults":       defaults,
        "dcf_base_price": dcf_px.get("gg_price"),
        "fetched_date":   stored.get("fetched", ""),
        "analyst_ests": [
            {
                "year":      str(e.get("date",""))[:4],
                "rev_avg_m": round((e.get("estimatedRevenueAvg") or 0) / 1e6, 1),
                "eps_avg":   round(e.get("estimatedEpsAvg") or 0, 2),
                "ebitda_m":  round((e.get("estimatedEbitdaAvg") or 0) / 1e6, 1),
            }
            for e in analyst_ests
        ],
    }


@app.route("/api/reports/discovered")
def discovered_reports():
    """Returns discovery tickers (not in core 24) that have rendered reports."""
    CORE = {
        "NVDA","MSFT","AAPL","ADBE","COST","AMD","JNJ","META","TSM","V",
        "KO","NFLX","ABBV","CSCO","WMT","F","WFC","INTC","TSLA","SOFI",
        "JPM","C","BAC","UAL",
    }
    reports_dir = os.path.join(os.path.dirname(__file__), "static", "reports")
    result = []
    if os.path.isdir(reports_dir):
        for fname in sorted(os.listdir(reports_dir), reverse=True):
            if fname.endswith("_report.html"):
                t = fname.replace("_report.html", "")
                if t not in CORE:
                    result.append({"ticker": t, "url": f"/reports/{fname}"})
    return jsonify(result)


@app.route("/api/reports")
def api_reports():
    _prune()
    return jsonify([
        {
            "rid":    rid,
            "ticker": r["ticker"],
            "score":  r["score"],
            "ts":     r["ts"].isoformat(),
        }
        for rid, r in sorted(_reports.items(), key=lambda x: x[1]["ts"], reverse=True)
    ])


# ── Dev entry point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
