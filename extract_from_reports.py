"""
extract_from_reports.py
Rebuilds static/data/TICKER_data.json for every existing HTML report
by parsing the chart data that is already embedded in the report HTML.

No FMP API calls. Uses the data that was already computed at report generation time.

Usage:  python extract_from_reports.py
        python extract_from_reports.py AAPL MSFT   # specific tickers only
"""

import os, sys, re, json, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_store import save_ticker_data, load_ticker_data, DATA_DIR

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "static", "reports")


def _parse_js_array(html, var_name):
    """Extract a JS array like `const rev = [1.2, 3.4];` → [1.2, 3.4]"""
    pattern = rf"const {re.escape(var_name)}\s*=\s*\[([^\]]+)\]"
    m = re.search(pattern, html)
    if not m:
        return []
    try:
        return [float(x.strip()) for x in m.group(1).split(",") if x.strip()]
    except ValueError:
        return []


def _parse_js_str_array(html, var_name):
    """Extract a JS string array like `const finLabels = ['FY2021','FY2022'];`"""
    pattern = rf"const {re.escape(var_name)}\s*=\s*\[([^\]]+)\]"
    m = re.search(pattern, html)
    if not m:
        return []
    return re.findall(r"'([^']+)'", m.group(1))


def _parse_float(html, var_name):
    """Extract `const waccValue = 9.01;` → 9.01"""
    pattern = rf"const {re.escape(var_name)}\s*=\s*([\d.]+)"
    m = re.search(pattern, html)
    return float(m.group(1)) if m else None


def _parse_current_price(html):
    """Extract current price from 'At $271.06,' pattern."""
    m = re.search(r"At \$([\d,]+\.?\d*),", html)
    if m:
        return float(m.group(1).replace(",", ""))
    # fallback: look for price in valuation paragraph
    m = re.search(r"trades at [\d.]+x.*?At \$([\d,]+\.?\d*)", html)
    if m:
        return float(m.group(1).replace(",", ""))
    return None


def _parse_dcf_prices(html):
    """Extract GG base price, bear, bull from report HTML."""
    # Pattern: "DCF base case (Gordon Growth, WACC X%): $NNN"
    gg = None
    m = re.search(r"DCF base case.*?Gordon Growth.*?\$([\d,]+)", html)
    if m:
        gg = float(m.group(1).replace(",", ""))

    # Bear / Base / Bull from the scenario table price cells
    prices = re.findall(r'class="price-[^"]*"[^>]*>\$([\d,]+)<', html)
    prices = [float(p.replace(",", "")) for p in prices]

    # The three scenario prices are bear, base, bull in order
    bear = prices[0] if len(prices) >= 1 else None
    base = prices[1] if len(prices) >= 2 else gg
    bull = prices[2] if len(prices) >= 3 else None

    if gg is None:
        gg = base

    return {"gg_price": gg, "em_price": base, "bear_price": bear, "bull_price": bull}


def _parse_company_name(html, ticker):
    """Try to extract company name from the report header."""
    m = re.search(r'<h1[^>]*>(.*?)</h1>', html, re.DOTALL)
    if m:
        name = re.sub(r'<[^>]+>', '', m.group(1)).strip()
        if name and name != ticker:
            return name
    # From profile block
    m = re.search(r'class="company-name"[^>]*>([^<]+)<', html)
    if m:
        return m.group(1).strip()
    return ticker


def extract_ticker(ticker, html_path, force=False):
    """
    Parse one report HTML and write the data store JSON.
    Returns True on success, False on failure.
    """
    if not force and load_ticker_data(ticker):
        print(f"  {ticker:6s}  already cached — skip (use --force to overwrite)")
        return True

    try:
        with open(html_path, "r", encoding="utf-8", errors="replace") as f:
            html = f.read()

        # ── Core arrays (values in $B) ─────────────────────────────────────────
        labels   = _parse_js_str_array(html, "finLabels")
        rev_b    = _parse_js_array(html, "rev")
        ebitda_b = _parse_js_array(html, "ebitda")
        ni_b     = _parse_js_array(html, "ni")
        ocf_b    = _parse_js_array(html, "ocfData")
        fcf_b    = _parse_js_array(html, "fcfData")
        wacc_pct = _parse_float(html, "waccValue")  # e.g. 9.01
        wacc_val = (wacc_pct / 100.0) if wacc_pct else None

        n = len(rev_b)
        if n == 0:
            print(f"  {ticker:6s}  FAIL — no revenue data found in HTML")
            return False

        # Pad missing arrays to same length as rev
        def _pad(arr, length, default=0.0):
            return (arr + [default] * length)[:length]

        ebitda_b = _pad(ebitda_b, n)
        ni_b     = _pad(ni_b,     n)
        ocf_b    = _pad(ocf_b,    n)
        fcf_b    = _pad(fcf_b,    n)

        # Year labels: strip 'FY' prefix → e.g. '2021'
        if labels:
            years = [lbl.replace("FY", "").replace("CY", "").strip() for lbl in labels]
        else:
            # Derive from report filename / defaults
            years = [str(2020 + i) for i in range(n)]

        # ── Reconstruct synthetic FMP-format data rows ─────────────────────────
        # _build_dcf_response reads these fields from the stored arrays.
        # Estimates for fields not available in HTML:
        #   D&A    ~ 3% of revenue (reasonable cross-sector default)
        #   CapEx  = OCF - FCF (standard FCF definition)
        #   PreTax = NI / 0.79  (assumes 21% effective tax)
        #   Tax    = PreTax * 0.21

        is_data, bs_data, cf_data = [], [], []
        for i in range(n):
            rev     = rev_b[i]    * 1e9
            ebitda  = ebitda_b[i] * 1e9
            ni      = ni_b[i]     * 1e9
            ocf     = ocf_b[i]    * 1e9
            fcf     = fcf_b[i]    * 1e9
            capex   = max(0.0, ocf - fcf)   # OCF - FCF ≈ CapEx (positive)
            da      = rev * 0.03            # 3% D&A estimate
            ebit    = ebitda - da
            pti     = ni / 0.79 if ni != 0 else 0
            tax_exp = pti * 0.21

            yr = years[i] if i < len(years) else str(2020 + i)

            is_data.append({
                "fiscalYear":                    yr,
                "date":                          f"{yr}-09-30",
                "revenue":                       rev,
                "ebitda":                        ebitda,
                "operatingIncome":               ebit,
                "netIncome":                     ni,
                "depreciationAndAmortization":   da,
                "incomeBeforeTax":               pti,
                "incomeTaxExpense":              tax_exp,
                "interestExpense":               0,
                "grossProfit":                   ebitda,     # approximation
            })
            bs_data.append({
                "fiscalYear":                    yr,
                "date":                          f"{yr}-09-30",
                "cashAndCashEquivalents":        0,
                "shortTermDebt":                 0,
                "longTermDebt":                  0,
                "totalStockholdersEquity":       ni * 5,     # rough placeholder
                "totalAssets":                   rev * 1.5,  # rough placeholder
            })
            cf_data.append({
                "fiscalYear":                    yr,
                "date":                          f"{yr}-09-30",
                "operatingCashFlow":             ocf,
                "capitalExpenditure":            -capex,     # FMP stores as negative
                "freeCashFlow":                  fcf,
                "depreciationAndAmortization":   da,
            })

        # ── Profile ────────────────────────────────────────────────────────────
        current_price = _parse_current_price(html)
        company_name  = _parse_company_name(html, ticker)
        dcf_prices    = _parse_dcf_prices(html)

        profile = {
            "symbol":           ticker,
            "companyName":      company_name,
            "price":            current_price or 0,
            "mktCap":           0,              # not in HTML; DCF still works
            "sharesOutstanding":0,              # not in HTML
            "beta":             1.0,            # not in HTML; user can adjust
        }

        # ── Save ──────────────────────────────────────────────────────────────
        save_ticker_data(
            ticker    = ticker,
            is_data   = is_data,
            bs_data   = bs_data,
            cf_data   = cf_data,
            profile   = profile,
            years     = years,
            wacc_val  = wacc_val,
            dcf_prices= dcf_prices,
            scorecard_metrics = {},
            analyst_ests      = [],
        )

        price_str = f"  price=${current_price:.2f}" if current_price else ""
        wacc_str  = f"  WACC={wacc_pct:.1f}%" if wacc_pct else ""
        gg_str    = f"  GG=${dcf_prices['gg_price']:.0f}" if dcf_prices.get("gg_price") else ""
        print(f"  {ticker:6s}  OK  {n}yr{price_str}{wacc_str}{gg_str}")
        return True

    except Exception as e:
        import traceback
        print(f"  {ticker:6s}  FAIL — {e}")
        traceback.print_exc()
        return False


def main():
    force = "--force" in sys.argv
    args  = [a for a in sys.argv[1:] if not a.startswith("--")]

    if args:
        # Specific tickers
        targets = [(t.upper(), os.path.join(REPORTS_DIR, f"{t.upper()}_report.html"))
                   for t in args]
    else:
        # All tickers with existing reports
        targets = []
        if os.path.isdir(REPORTS_DIR):
            for fname in sorted(os.listdir(REPORTS_DIR)):
                if fname.endswith("_report.html"):
                    t = fname.replace("_report.html", "")
                    targets.append((t, os.path.join(REPORTS_DIR, fname)))

    if not targets:
        print("No reports found in static/reports/")
        return

    os.makedirs(DATA_DIR, exist_ok=True)
    print(f"\nExtracting data store from {len(targets)} report(s)")
    print(f"Output: {DATA_DIR}")
    print("=" * 50)

    ok = 0; failed = 0
    for ticker, path in targets:
        if not os.path.exists(path):
            print(f"  {ticker:6s}  SKIP — report not found: {path}")
            continue
        if extract_ticker(ticker, path, force=force):
            ok += 1
        else:
            failed += 1

    print("=" * 50)
    print(f"Done: {ok} extracted, {failed} failed")
    if ok:
        print(f"\nNext: git add static/data/ && git commit -m 'seed data store from reports' && git push")


if __name__ == "__main__":
    main()
