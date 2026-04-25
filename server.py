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
import json
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
from scenarios_db import init_db, save_scenario, list_scenarios, delete_scenario, get_scenario

# ── App setup ─────────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder="static", static_url_path="")

# ── Config from environment ───────────────────────────────────────────────────
mdl.API_KEY  = os.environ.get("FMP_API_KEY", mdl.API_KEY)
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

# ── Initialise scenario database ─────────────────────────────────────────────
init_db()

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
        # If this ticker has an existing report, its data should be in the store.
        # If not seeded yet, fetch from FMP.
        report_path = os.path.join(os.path.dirname(__file__), "static", "reports",
                                   f"{ticker}_report.html")
        if os.path.exists(report_path):
            return jsonify({
                "error": f"Data for {ticker} is not yet cached. "
                         f"Please re-run batch_reports.py or seed_data_store.py to populate the data store.",
                "hint": "report_exists_but_not_seeded"
            }), 503

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
    """Build the DCF API response from stored data (no FMP calls).
    Prefers Excel model data (excel_dcf key) when available — exact assumptions.
    Falls back to FMP-derived data for new tickers not yet in Excel models.
    """
    excel    = stored.get("excel_dcf")           # present when read from Excel
    profile  = stored.get("profile") or {}
    ae_raw   = stored.get("analyst_ests") or []
    dcf_px   = stored.get("dcf_prices") or {}
    ticker   = stored["ticker"]

    # ── Excel path: use exact model assumptions ────────────────────────────
    if excel:
        hist_rows = excel.get("hist") or []
        proj_rows = excel.get("proj") or []
        wi        = excel.get("wacc_inputs") or {}

        history = []
        for i, h in enumerate(hist_rows):
            history.append({
                "year":          h.get("year", ""),
                "revenue_m":     round(h["rev_mm"], 1) if h.get("rev_mm") else 0,
                "rev_growth":    h.get("rev_growth"),
                "ebitda_margin": h.get("ebitda_margin") or 0,
                "da_pct":        h.get("da_pct") or 0,
                "capex_pct":     h.get("capex_pct") or 0,
                "tax_rate":      h.get("tax_rate") or 0.21,
                "fcf_m":         round(h["ufcf_mm"], 1) if h.get("ufcf_mm") else 0,
            })
        # Fill in rev_growth for first year if missing
        for i in range(1, len(history)):
            if history[i]["rev_growth"] is None:
                prev = history[i-1]["revenue_m"]
                curr = history[i]["revenue_m"]
                history[i]["rev_growth"] = round(curr / prev - 1, 4) if prev else None

        # Projection defaults — exact from Excel
        def_rev    = [p.get("rev_growth")    or 0.05 for p in proj_rows]
        def_margin = [p.get("ebitda_margin") or (history[-1]["ebitda_margin"] if history else 0.20) for p in proj_rows]
        def_da     = [p.get("da_pct")        or (history[-1]["da_pct"] if history else 0.03) for p in proj_rows]
        def_cx     = [p.get("capex_pct")     or (history[-1]["capex_pct"] if history else 0.03) for p in proj_rows]
        def_nwc    = [p.get("nwc_pct")       or 0.005 for p in proj_rows]
        def_tax    = [p.get("tax_rate")       or 0.21  for p in proj_rows]
        # Pad to 5 if fewer projection years
        while len(def_rev) < 5:
            def_rev.append(0.05); def_margin.append(def_margin[-1] if def_margin else 0.20)
            def_da.append(def_da[-1] if def_da else 0.03); def_cx.append(def_cx[-1] if def_cx else 0.03)
            def_nwc.append(0.005); def_tax.append(0.21)

        wacc = wi.get("wacc") or stored.get("wacc_val") or 0.09
        defaults = {
            "rev_growth":    [round(v, 4) for v in def_rev[:5]],
            "ebitda_margin": [round(v, 4) for v in def_margin[:5]],
            "da_pct":        [round(v, 4) for v in def_da[:5]],
            "capex_pct":     [round(v, 4) for v in def_cx[:5]],
            "nwc_pct":       [round(v, 4) for v in def_nwc[:5]],
            "tax_rate":      [round(v, 4) for v in def_tax[:5]],
            "tgr":           excel.get("tgr") or 0.03,
            "exit_multiple": excel.get("exit_multiple") or 15.0,
            "rf":            wi.get("rf")       or 0.043,
            "beta":          wi.get("beta")      or 1.0,
            "erp":           wi.get("erp")       or 0.045,
            "kd_pretax":     wi.get("kd_pretax") or 0.04,
            "eff_tax":       wi.get("tax_rate")  or 0.21,
            "equity_weight": wi.get("equity_weight") or 0.90,
            "wacc":          round(wacc, 4),
        }

        current_price = excel.get("current_price") or float(profile.get("price") or 0) or None
        shares_m      = excel.get("shares_mm") or (float(profile.get("sharesOutstanding") or 0) / 1e6)
        net_debt_m    = excel.get("net_debt_mm") or 0
        last_year     = int(str(hist_rows[-1].get("year", 2024))[:4]) if hist_rows else 2024
        gg_price      = dcf_px.get("gg_price")
        source        = "excel"

    # ── FMP fallback: derive from raw 3-statement data ─────────────────────
    else:
        is_data  = stored["is_data"]
        bs_data  = stored["bs_data"]
        cf_data  = stored["cf_data"]
        years    = stored.get("years") or []
        wacc_val = stored.get("wacc_val")
        beta     = float(profile.get("beta") or 1.0) or 1.0
        market_cap = float(profile.get("mktCap") or 0) or None
        shares     = float(profile.get("sharesOutstanding") or 0)

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

        bs0      = bs_data[-1]
        cash     = bs0.get("cashAndCashEquivalents") or 0
        debt     = (bs0.get("shortTermDebt") or 0) + (bs0.get("longTermDebt") or 0)
        net_debt = debt - cash

        RF  = 0.043; ERP = 0.045
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

        last_g  = history[-1]["rev_growth"] or 0.05
        last_m  = history[-1]["ebitda_margin"] or 0.20
        last_da = history[-1]["da_pct"] or 0.03
        last_cx = history[-1]["capex_pct"] or 0.04

        def_rev = []
        for i in range(5):
            if i < len(ae_raw):
                er  = ae_raw[i].get("estimatedRevenueAvg") or 0
                epr = (ae_raw[i-1].get("estimatedRevenueAvg")
                       if i > 0 else history[-1]["revenue_m"] * 1e6) or 0
                if er and epr:
                    def_rev.append(round(er / epr - 1, 4)); continue
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

        current_price = float(profile.get("price") or 0) or None
        shares_m      = round(shares / 1e6, 2)
        net_debt_m    = round(net_debt / 1e6, 1)
        last_year     = int(years[-1]) if years else 2024
        gg_price      = dcf_px.get("gg_price")
        source        = "fmp"

    return {
        "ticker":         ticker,
        "name":           profile.get("companyName") or ticker,
        "price":          current_price,
        "shares_m":       round(shares_m, 2),
        "net_debt_m":     round(net_debt_m, 1),
        "last_year":      last_year,
        "history":        history,
        "defaults":       defaults,
        "dcf_base_price": gg_price,
        "data_source":    source,
        "fetched_date":   stored.get("fetched", ""),
        "analyst_ests": [
            {
                "year":      str(e.get("date",""))[:4],
                "rev_avg_m": round((e.get("estimatedRevenueAvg") or 0) / 1e6, 1),
                "eps_avg":   round(e.get("estimatedEpsAvg") or 0, 2),
                "ebitda_m":  round((e.get("estimatedEbitdaAvg") or 0) / 1e6, 1),
            }
            for e in ae_raw
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


@app.route("/news")
def news_page():
    return app.send_static_file("news.html")


@app.route("/api/news")
def api_news():
    ticker = request.args.get("ticker", "").upper().strip()
    cache_path = os.path.join(os.path.dirname(__file__), "static", "data", "news_cache.json")

    if not os.path.exists(cache_path):
        return jsonify({"articles": [], "tickers": [], "fetched": None, "stale": True})

    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            cache = json.load(f)
    except (json.JSONDecodeError, IOError):
        return jsonify({"articles": [], "tickers": [], "fetched": None, "stale": True})

    fetched = cache.get("fetched")
    stale = True
    if fetched:
        try:
            fetched_dt = datetime.datetime.fromisoformat(fetched)
            stale = (datetime.datetime.utcnow() - fetched_dt).total_seconds() > 26 * 3600
        except (ValueError, TypeError):
            stale = True

    articles = cache.get("articles", [])
    if ticker:
        articles = [a for a in articles if (a.get("symbol") or "").upper() == ticker]

    return jsonify({
        "articles": articles,
        "tickers":  cache.get("tickers", []),
        "fetched":  fetched,
        "stale":    stale,
    })


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


# ── Scenario API ──────────────────────────────────────────────────────────────

@app.route("/api/scenarios", methods=["GET"])
def api_list_scenarios():
    ticker = request.args.get("ticker", "").strip().upper()
    if not ticker:
        return jsonify({"error": "ticker query param required"}), 400
    return jsonify(list_scenarios(ticker))


@app.route("/api/scenarios", methods=["POST"])
def api_save_scenario():
    body = request.get_json(force=True) or {}
    ticker  = (body.get("ticker") or "").strip().upper()
    name    = (body.get("name") or "").strip()
    inputs  = body.get("inputs") or {}
    outputs = body.get("outputs") or {}
    if not ticker or not name:
        return jsonify({"error": "ticker and name required"}), 400
    sid = save_scenario(ticker, name, inputs, outputs)
    return jsonify({"id": sid, "ok": True})


@app.route("/api/scenarios", methods=["DELETE"])
def api_delete_scenario():
    body = request.get_json(force=True) or {}
    ticker = (body.get("ticker") or "").strip().upper()
    name   = (body.get("name") or "").strip()
    if not ticker or not name:
        return jsonify({"error": "ticker and name required"}), 400
    delete_scenario(ticker, name)
    return jsonify({"ok": True})


@app.route("/api/scenarios/compare", methods=["GET"])
def api_compare_scenarios():
    ticker = request.args.get("ticker", "").strip().upper()
    names  = request.args.get("names", "")
    if not ticker:
        return jsonify({"error": "ticker query param required"}), 400
    scenarios = list_scenarios(ticker)
    if names and ticker != "ALL":
        name_set = set(n.strip() for n in names.split(",") if n.strip())
        scenarios = [s for s in scenarios if s["name"] in name_set]
    return jsonify(scenarios)


# ── Heatmap ───────────────────────────────────────────────────────────────────

# Static sector buckets for the core coverage universe
_SECTORS = {
    "AAPL":  "Technology",  "ADBE":  "Technology",   "AMD":   "Technology",
    "CSCO":  "Technology",  "INTC":  "Technology",   "MSFT":  "Technology",
    "NVDA":  "Technology",  "TSM":   "Technology",
    "META":  "Comm & Media","NFLX":  "Comm & Media",
    "BAC":   "Financials",  "C":     "Financials",   "JPM":   "Financials",
    "SOFI":  "Financials",  "V":     "Financials",   "WFC":   "Financials",
    "ABBV":  "Healthcare",  "JNJ":   "Healthcare",
    "F":     "Consumer",    "TSLA":  "Consumer",     "UAL":   "Consumer",
    "COST":  "Staples",     "KO":    "Staples",      "WMT":   "Staples",
}

_COMPANY_NAMES = {
    "AAPL":"Apple","ADBE":"Adobe","AMD":"AMD","CSCO":"Cisco","INTC":"Intel",
    "MSFT":"Microsoft","NVDA":"NVIDIA","TSM":"TSMC","META":"Meta","NFLX":"Netflix",
    "BAC":"BofA","C":"Citigroup","JPM":"JPMorgan","SOFI":"SoFi","V":"Visa","WFC":"Wells Fargo",
    "ABBV":"AbbVie","JNJ":"J&J","F":"Ford","TSLA":"Tesla","UAL":"United Airlines",
    "COST":"Costco","KO":"Coca-Cola","WMT":"Walmart",
}

@app.route("/heatmap")
def heatmap_page():
    return app.send_static_file("heatmap.html")


@app.route("/api/heatmap-data")
def api_heatmap_data():
    csv_path = os.path.join(os.path.dirname(__file__), "outputs.csv")
    if not os.path.exists(csv_path):
        return jsonify({"tickers": []})

    import csv as _csv
    def _f(v):
        try: return float(v) if v not in ("", None) else None
        except: return None

    tickers = []
    with open(csv_path, "r", encoding="utf-8") as f:
        for row in _csv.DictReader(f):
            t = row.get("Ticker", "").strip().upper()
            if not t:
                continue
            mktcap  = _f(row.get("MktCap_B"))
            rev     = _f(row.get("Revenue_B"))
            price   = _f(row.get("Price"))
            score   = _f(row.get("Auto_Score"))
            gg_up   = _f(row.get("GG_Upside"))
            em_up   = _f(row.get("EM_Upside"))
            roic    = _f(row.get("ROIC"))
            cagr    = _f(row.get("Rev_CAGR"))
            pe      = _f(row.get("PE_Current"))
            pfcf    = _f(row.get("PFCF_Current"))
            d_eb    = _f(row.get("D_EBITDA"))
            fcf_ni  = _f(row.get("FCF_NI"))
            gg_px   = _f(row.get("GG_Price"))
            em_px   = _f(row.get("EM_Price"))

            # Size fallback: use revenue if no mktcap
            size = mktcap or rev or 10.0

            tickers.append({
                "ticker":    t,
                "name":      _COMPANY_NAMES.get(t, t),
                "sector":    _SECTORS.get(t, "Other"),
                "size":      size,
                "mktcap_b":  mktcap,
                "revenue_b": rev,
                "price":     price,
                "gg_px":     gg_px,
                "em_px":     em_px,
                "score":     score,
                "gg_upside": round(gg_up * 100, 1) if gg_up is not None else None,
                "em_upside": round(em_up * 100, 1) if em_up is not None else None,
                "roic":      round(roic * 100, 1) if roic is not None else None,
                "rev_cagr":  round(cagr * 100, 1) if cagr is not None else None,
                "pe":        pe,
                "pfcf":      pfcf,
                "d_ebitda":  d_eb,
                "fcf_ni":    fcf_ni,
            })

    return jsonify({"tickers": tickers})


# ── Dev entry point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
