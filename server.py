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

        # ── Build HTML report ─────────────────────────────────────────────────
        report_data = build_report_data(
            ticker           = ticker,
            profile          = profile,
            is_data          = is_data,
            bs_data          = bs_data,
            cf_data          = cf_data,
            years            = years,
            wacc_val         = wacc_refs.get("wacc_val"),
            dcf_prices       = (dcf_refs or {}).get("dcf_prices") or {},
            scorecard_metrics= scorecard_metrics,
            manual_rating    = manual_rating,
            current_price    = current_price,
            market_cap       = market_cap,
        )
        html_content = render_html_report(report_data)

        # ── Store with short ID ───────────────────────────────────────────────
        rid = uuid.uuid4().hex[:10]
        _reports[rid] = {
            "ticker":  ticker,
            "html":    html_content,
            "excel":   excel_bytes,
            "year":    years[-1] if years else "2025",
            "ts":      datetime.datetime.now(),
            "score":   scorecard_metrics.get("auto_score"),
        }

        return jsonify({
            "report_id":  rid,
            "ticker":     ticker,
            "auto_score": scorecard_metrics.get("auto_score"),
            "logs":       logs[-15:],
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


# ── Dev entry point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
