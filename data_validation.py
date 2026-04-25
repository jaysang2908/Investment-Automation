"""
data_validation.py
FMP data anomaly detection layer — runs after financial data is fetched
and flags suspicious entries before they corrupt reports or DCF models.

Returns informational warnings only; never raises exceptions or blocks
report generation.
"""

import os
import json
import datetime


def validate_fmp_data(ticker, is_data, bs_data, cf_data) -> list[dict]:
    """
    Returns list of warning dicts: {severity, check, year, message}
    severity: "ERROR" | "WARNING" | "INFO"
    Never raises exceptions -- always returns a list (possibly empty).
    """
    warnings = []

    try:
        warnings.extend(_check_revenue_discontinuity(is_data))
        warnings.extend(_check_fcf_ni_divergence(is_data, cf_data))
        warnings.extend(_check_balance_sheet_identity(bs_data))
        warnings.extend(_check_missing_critical_fields(is_data, bs_data, cf_data))
        warnings.extend(_check_temporal_consistency(is_data))
        warnings.extend(_check_ebitda_sanity(is_data, cf_data))
        warnings.extend(_check_negative_equity(bs_data))
    except Exception:
        # Catch-all: validation must never blow up the pipeline
        pass

    return warnings


# ── Individual checks ────────────────────────────────────────────────────────

def _get_year(record):
    """Extract fiscal year string from a statement record."""
    # FMP data typically has calendarYear or date field
    yr = record.get("calendarYear") or ""
    if not yr:
        date_str = record.get("date") or record.get("fillingDate") or ""
        yr = date_str[:4] if len(date_str) >= 4 else "?"
    return str(yr)


def _check_revenue_discontinuity(is_data):
    """YoY revenue change > +/-50% for companies with revenue > $1B."""
    results = []
    try:
        for i in range(1, len(is_data)):
            cur_rev = is_data[i].get("revenue") or 0
            prv_rev = is_data[i - 1].get("revenue") or 0
            if prv_rev == 0 or abs(prv_rev) < 1e9:
                continue
            yoy = (cur_rev / prv_rev) - 1
            if abs(yoy) > 0.50:
                yr = _get_year(is_data[i])
                prv_yr = _get_year(is_data[i - 1])
                results.append({
                    "severity": "WARNING",
                    "check": "revenue_discontinuity",
                    "year": yr,
                    "message": (
                        f"FY{yr} revenue ${cur_rev/1e9:.1f}B vs FY{prv_yr} "
                        f"${prv_rev/1e9:.1f}B ({yoy:+.1%}) "
                        f"-- verify for M&A or restatement"
                    ),
                })
    except Exception:
        pass
    return results


def _check_fcf_ni_divergence(is_data, cf_data):
    """FCF/NI divergence: abs(FCF/NI) > 3.0 or FCF/NI < -1.0 for profitable companies."""
    results = []
    try:
        for is_rec, cf_rec in zip(is_data, cf_data):
            ni = is_rec.get("netIncome") or 0
            if ni <= 0:
                continue  # only check profitable years
            ocf = cf_rec.get("operatingCashFlow") or 0
            capex = abs(cf_rec.get("capitalExpenditure") or 0)
            fcf = cf_rec.get("freeCashFlow") or (ocf - capex)
            ratio = fcf / ni
            if abs(ratio) > 3.0 or ratio < -1.0:
                yr = _get_year(is_rec)
                results.append({
                    "severity": "WARNING",
                    "check": "fcf_ni_divergence",
                    "year": yr,
                    "message": (
                        f"FY{yr} FCF/NI ratio {ratio:.1f}x "
                        f"(FCF ${fcf/1e9:.1f}B, NI ${ni/1e9:.1f}B) "
                        f"-- large divergence may indicate working capital swings or non-cash items"
                    ),
                })
    except Exception:
        pass
    return results


def _check_balance_sheet_identity(bs_data):
    """abs(totalAssets - totalLiabilities - totalStockholdersEquity) / totalAssets > 2%."""
    results = []
    try:
        for bs_rec in bs_data:
            ta = bs_rec.get("totalAssets") or 0
            tl = bs_rec.get("totalLiabilities") or 0
            te = bs_rec.get("totalStockholdersEquity") or 0
            if ta == 0 or tl == 0 or te == 0:
                continue
            diff = abs(ta - tl - te)
            if diff / abs(ta) > 0.02:
                yr = _get_year(bs_rec)
                results.append({
                    "severity": "WARNING",
                    "check": "balance_sheet_identity",
                    "year": yr,
                    "message": (
                        f"FY{yr} balance sheet identity gap: "
                        f"Assets ${ta/1e9:.1f}B - Liabilities ${tl/1e9:.1f}B "
                        f"- Equity ${te/1e9:.1f}B = ${diff/1e6:.0f}M residual "
                        f"({diff/abs(ta)*100:.1f}% of assets)"
                    ),
                })
    except Exception:
        pass
    return results


def _check_missing_critical_fields(is_data, bs_data, cf_data):
    """Check that critical fields are non-null and non-zero for the most recent year."""
    results = []
    try:
        if not is_data or not bs_data or not cf_data:
            return results

        is0 = is_data[-1]
        bs0 = bs_data[-1]
        cf0 = cf_data[-1]
        yr = _get_year(is0)

        checks = [
            ("revenue", is0.get("revenue")),
            ("operatingIncome", is0.get("operatingIncome")),
            ("netIncome", is0.get("netIncome")),
            ("totalAssets", bs0.get("totalAssets")),
            ("totalStockholdersEquity", bs0.get("totalStockholdersEquity")),
            ("operatingCashFlow", cf0.get("operatingCashFlow")),
        ]
        for field, val in checks:
            if val is None or val == 0:
                results.append({
                    "severity": "WARNING",
                    "check": "missing_critical_field",
                    "year": yr,
                    "message": (
                        f"FY{yr} {field} is {'null' if val is None else 'zero'} "
                        f"-- may indicate incomplete data from FMP"
                    ),
                })
    except Exception:
        pass
    return results


def _check_temporal_consistency(is_data):
    """Check fiscal years are sequential with no gaps > 1 year."""
    results = []
    try:
        years_int = []
        for rec in is_data:
            yr_str = _get_year(rec)
            try:
                years_int.append(int(yr_str))
            except (ValueError, TypeError):
                continue
        for i in range(1, len(years_int)):
            gap = abs(years_int[i] - years_int[i - 1])
            if gap > 1:
                results.append({
                    "severity": "WARNING",
                    "check": "temporal_gap",
                    "year": str(years_int[i]),
                    "message": (
                        f"Gap of {gap} years between FY{years_int[i-1]} "
                        f"and FY{years_int[i]} -- missing annual filings?"
                    ),
                })
    except Exception:
        pass
    return results


def _check_ebitda_sanity(is_data, cf_data):
    """EBITDA should be >= EBIT (operating income) since D&A is positive."""
    results = []
    try:
        for is_rec, cf_rec in zip(is_data, cf_data):
            ebitda = is_rec.get("ebitda")
            if ebitda is None:
                continue
            ebit = is_rec.get("operatingIncome") or 0
            if ebitda < ebit:
                yr = _get_year(is_rec)
                results.append({
                    "severity": "WARNING",
                    "check": "ebitda_below_ebit",
                    "year": yr,
                    "message": (
                        f"FY{yr} EBITDA ${ebitda/1e9:.1f}B < EBIT ${ebit/1e9:.1f}B "
                        f"-- EBITDA should be >= EBIT (D&A is additive); "
                        f"possible data error"
                    ),
                })
    except Exception:
        pass
    return results


def _check_negative_equity(bs_data):
    """Negative stockholders' equity -- legitimate for some companies but affects ratios."""
    results = []
    try:
        for bs_rec in bs_data:
            te = bs_rec.get("totalStockholdersEquity")
            if te is not None and te < 0:
                yr = _get_year(bs_rec)
                results.append({
                    "severity": "INFO",
                    "check": "negative_equity",
                    "year": yr,
                    "message": (
                        f"FY{yr} stockholders' equity is negative "
                        f"(${te/1e9:.1f}B) -- may affect ROIC and "
                        f"leverage ratios; common for buyback-heavy companies"
                    ),
                })
    except Exception:
        pass
    return results


# ── Persistence helper ───────────────────────────────────────────────────────

def persist_anomalies(ticker, anomalies, data_dir):
    """
    Append to static/data/anomalies.json. Prune entries older than 90 days.
    Never raises exceptions.
    """
    try:
        os.makedirs(data_dir, exist_ok=True)
        path = os.path.join(data_dir, "anomalies.json")

        # Load existing
        existing = []
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                existing = json.load(f)

        today = datetime.date.today().isoformat()
        cutoff = (datetime.date.today() - datetime.timedelta(days=90)).isoformat()

        # Prune entries older than 90 days
        existing = [e for e in existing if e.get("date", "") >= cutoff]

        # Remove previous entries for this ticker (replace with fresh run)
        existing = [e for e in existing if e.get("ticker") != ticker]

        # Append new
        for a in anomalies:
            existing.append({
                "ticker": ticker,
                "date": today,
                **a,
            })

        with open(path, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=2)
    except Exception:
        pass  # persistence is best-effort; never block the pipeline
