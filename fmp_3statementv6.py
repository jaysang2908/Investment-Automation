import os
import datetime
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# Output files always save next to this script, regardless of working directory
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION — update API_KEY before running
# ═══════════════════════════════════════════════════════════════════════════════
API_KEY      = "tOPLCq7cEELfef0FA6AKNuVoO549gAS1"
YEARS        = 5    # historical years to fetch
YEARS_PROJ   = 5    # minimum projection years in DCF (auto-extended if FMP has more)

# ── Damodaran tables (Jan 2025 US data — update annually) ─────────────────────
# ICR band → (synthetic rating, default spread)
DAMODARAN_SPREADS = [
    (12.5, 1e9,  "AAA",  0.0054),
    ( 9.5, 12.5, "AA",   0.0072),
    ( 7.5,  9.5, "A+",   0.0096),
    ( 6.0,  7.5, "A",    0.0108),
    ( 4.5,  6.0, "A-",   0.0132),
    ( 4.0,  4.5, "BBB+", 0.0156),
    ( 3.5,  4.0, "BBB",  0.0180),
    ( 3.0,  3.5, "BBB-", 0.0240),
    ( 2.5,  3.0, "BB+",  0.0288),
    ( 2.0,  2.5, "BB",   0.0360),
    ( 1.5,  2.0, "B+",   0.0432),
    ( 1.25, 1.5, "B",    0.0540),
    ( 0.8,  1.25,"B-",   0.0648),
    ( 0.5,  0.8, "CCC",  0.0900),
    ( 0.0,  0.5, "CC",   0.1200),
]
# Moody's → S&P rating equivalents
MOODY_TO_SP = {
    "Aaa": "AAA", "Aa1": "AA+", "Aa2": "AA",  "Aa3": "AA-",
    "A1":  "A+",  "A2":  "A",   "A3":  "A-",
    "Baa1":"BBB+","Baa2":"BBB", "Baa3":"BBB-",
    "Ba1": "BB+", "Ba2": "BB",  "Ba3": "BB-",
    "B1":  "B+",  "B2":  "B",   "B3":  "B-",
    "Caa1":"CCC+","Caa2":"CCC", "Caa3":"CCC-","Ca": "CC",
}
VALID_SP_RATINGS = {
    "AAA","AA+","AA","AA-","A+","A","A-",
    "BBB+","BBB","BBB-","BB+","BB","BB-",
    "B+","B","B-","CCC+","CCC","CCC-","CC","C","D",
}
# Industry → unlevered beta (US, Jan 2025)
DAMODARAN_BETAS = {
    "Semiconductor":         1.15,
    "Software":              1.08,
    "Technology":            1.07,
    "Computer":              0.95,
    "Internet":              1.12,
    "Electronics":           1.05,
    "Telecom":               0.68,
    "Retail":                0.82,
    "Healthcare":            0.76,
    "Pharmaceutical":        0.74,
    "Financial":             0.55,
    "Insurance":             0.60,
    "Oil":                   0.78,
    "Energy":                0.82,
    "Automobile":            0.88,
    "Consumer":              0.80,
    "Industrial":            0.88,
    "Default":               1.00,
}
# Damodaran implied ERP — US market (update annually from pages.stern.nyu.edu)
DAMODARAN_ERP_IMPLIED  = 0.0472   # Jan 2026
DAMODARAN_ERP_HIST_AVG = 0.0420   # arithmetic avg 1928–2025
# Peer tickers by sector (for beta comparison)
SECTOR_PEERS = {
    "Semiconductors":  ["AMD", "INTC", "QCOM", "AVGO", "TSM"],
    "Software":        ["MSFT", "CRM", "ORCL", "ADBE", "NOW"],
    "Technology":      ["AAPL", "MSFT", "GOOGL", "META", "AMZN"],
    "Healthcare":      ["JNJ",  "UNH",  "ABT",  "MDT",  "BMY"],
    "Financials":      ["JPM",  "BAC",  "GS",   "MS",   "WFC"],
    "Consumer":        ["AMZN", "HD",   "MCD",  "NKE",  "SBUX"],
    "Energy":          ["XOM",  "CVX",  "COP",  "SLB",  "PSX"],
}

# ── Colours ───────────────────────────────────────────────────────────────────
C_TITLE      = "1F2D3D"
C_SECTION    = "2E4057"
C_SUMMARY_HD = "1A3A5C"
C_SUMMARY_BG = "EAF2FB"
C_DETAIL_HD  = "34495E"
C_ALT        = "F4F8FB"
C_SUBTOTAL   = "D6E4F0"
C_WHITE      = "FFFFFF"
C_BLUE       = "0000FF"   # hardcoded inputs
C_AI_BG      = "FFF9C4"   # amber  — AI recommendation rows
C_AI_RAT     = "FFFDE7"   # pale   — AI rationale rows
C_FLAG_BG    = "FFEBEE"   # red    — warning / flag rows
C_OVR_BG     = "F1F8E9"   # green  — selected / override rows
C_BLACK      = "000000"   # formula outputs
C_GREEN      = "006400"   # cross-sheet links
# DCF-specific colours
C_SUB        = "D6E4F0"   # subtotal / header rows  (same as C_SUBTOTAL)
C_ASSM       = "EBF5FB"   # assumption / input rows
C_HIST       = "F8FBFD"   # historical data rows
C_CONS       = "D4E6F1"   # FMP consensus-driven projection rows
C_BG         = "D4E6F1"   # alias: consensus/projection background in DCF
C_SECT       = "2E4057"   # DCF section header background (same as C_SECTION)
C_HD         = "1A3A5C"   # DCF sub-section header (slightly lighter)

# ── Fonts / fills / borders ───────────────────────────────────────────────────
def fnt(bold=False, color=C_BLACK, size=10, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)

def fll(hex_color):
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)

def brd(color="B0B8C1"):
    t = Side(style="thin", color=color)
    return Border(left=t, right=t, top=t, bottom=t)

def pct_fmt(cell):   cell.number_format = '0.0%;(0.0%);"-"'
def num_fmt(cell):   cell.number_format = '#,##0.0;(#,##0.0);"-"'
def ratio_fmt(cell): cell.number_format = '0.0x;(0.0x);"-"'
def days_fmt(cell):  cell.number_format = '#,##0.0;(#,##0.0);"-"'

def cl(col): return get_column_letter(col)

# ═══════════════════════════════════════════════════════════════════════════════
# API FETCH
# ═══════════════════════════════════════════════════════════════════════════════
def fetch(endpoint, ticker, extra_params=""):
    url = (f"https://financialmodelingprep.com/stable/{endpoint}"
           f"?symbol={ticker}&limit={YEARS}{extra_params}&apikey={API_KEY}")
    print(f"  GET {endpoint}...")
    r = requests.get(url)
    print(f"  -> {r.status_code}")
    if r.status_code != 200:
        raise ValueError(f"HTTP {r.status_code} on {endpoint}")
    if not r.text.strip():
        raise ValueError(f"Empty response: {endpoint}/{ticker}")
    try:
        data = r.json()
    except Exception as e:
        raise ValueError(f"JSON parse failed: {e}\nRaw: {r.text[:200]}")
    if isinstance(data, dict):
        msg = data.get("Error Message") or data.get("message", "")
        if msg:
            raise ValueError(f"API error: {msg}")
    if not isinstance(data, list) or len(data) == 0:
        raise ValueError(f"No data for '{ticker}' on {endpoint}.")
    return data

def fetch_segment(endpoint, ticker):
    """Fetch segmentation — returns None gracefully if not on plan."""
    try:
        url = (f"https://financialmodelingprep.com/stable/{endpoint}"
               f"?symbol={ticker}&apikey={API_KEY}")
        r = requests.get(url)
        if r.status_code != 200:
            return None
        data = r.json()
        if not data or (isinstance(data, dict) and ("Error" in str(data) or "message" in data)):
            return None
        return data if isinstance(data, list) else None
    except:
        return None

# ═══════════════════════════════════════════════════════════════════════════════
# WACC HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def fetch_fred(series_id):
    """Return (value_as_decimal, date_string) for latest FRED observation.
    Uses public CSV endpoint — no API key required."""
    try:
        csv = requests.get(
            f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}",
            timeout=10
        ).text.strip().split("\n")
        last = next(r for r in reversed(csv)
                    if r and r.split(",")[1] not in (".", ""))
        date, val = last.split(",")
        return float(val) / 100, date
    except Exception:
        return None, None

def fetch_analyst_estimates(ticker, last_hist_year):
    """Fetch FMP annual analyst estimates, return only forward years (sorted oldest→newest).
    Each record: {year, rev_avg, rev_low, rev_high, ebitda_avg, ebitda_low, ebitda_high,
                  ni_avg, eps_avg, n_analysts_rev, n_analysts_eps}
    Returns [] gracefully on any failure."""
    try:
        url = (f"https://financialmodelingprep.com/stable/analyst-estimates"
               f"?symbol={ticker}&period=annual&limit=10&apikey={API_KEY}")
        r = requests.get(url, timeout=10)
        if r.status_code != 200:
            print(f"  [Estimates] HTTP {r.status_code}")
            return []
        raw = r.json()
        if not isinstance(raw, list):
            return []
        out = []
        for rec in raw:
            yr = str(rec.get("date", ""))[:4]
            if yr <= str(last_hist_year):
                continue          # skip historical estimate years
            out.append({
                "year":          yr,
                "rev_avg":       (rec.get("revenueAvg")   or 0) / 1e6,
                "rev_low":       (rec.get("revenueLow")   or 0) / 1e6,
                "rev_high":      (rec.get("revenueHigh")  or 0) / 1e6,
                "ebitda_avg":    (rec.get("ebitdaAvg")    or 0) / 1e6,
                "ebitda_low":    (rec.get("ebitdaLow")    or 0) / 1e6,
                "ebitda_high":   (rec.get("ebitdaHigh")   or 0) / 1e6,
                "ni_avg":        (rec.get("netIncomeAvg") or 0) / 1e6,
                "eps_avg":       rec.get("epsAvg"),
                "n_rev":         rec.get("numAnalystsRevenue") or 0,
                "n_eps":         rec.get("numAnalystsEps")     or 0,
            })
        # Sort oldest → newest
        out.sort(key=lambda x: x["year"])
        print(f"  [Estimates] {len(out)} forward years: "
              f"{[e['year'] for e in out]}")
        return out
    except Exception as e:
        print(f"  [Estimates] Failed: {e}")
        return []

def get_synthetic_rating(icr):
    """Map Interest Coverage Ratio to Damodaran synthetic rating + spread."""
    for lo, hi, rating, spread in DAMODARAN_SPREADS:
        if lo <= icr < hi:
            return rating, spread
    return "CC", 0.12


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def g(rec, key):
    v = rec.get(key)
    try:   return float(v) if v is not None else None
    except: return None

def gm(rec, key):
    v = g(rec, key)
    return round(v / 1e6, 2) if v is not None else None

def g_any(rec, *keys):
    """Try multiple field names, return first non-None value found."""
    for k in keys:
        v = g(rec, k)
        if v is not None:
            return v
    return None

def gm_any(rec, *keys):
    """g_any but scaled to $mm."""
    v = g_any(rec, *keys)
    return round(v / 1e6, 2) if v is not None else None

def setup_ws(ws, years, col_a_width=42):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = col_a_width
    for i in range(len(years)):
        ws.column_dimensions[cl(i+2)].width = 16
    ws.freeze_panes = "B3"

def write_tab_title(ws, row, text, ncols, subtitle=None):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font  = fnt(bold=True, color=C_WHITE, size=13)
    c.fill  = fll(C_TITLE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 28
    if subtitle:
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        s = ws.cell(row=row, column=1, value=subtitle)
        s.font = fnt(size=9, italic=True, color="888888")
        ws.row_dimensions[row].height = 14
    return row + 1

def write_section_hdr(ws, row, text, ncols, color=None):
    bg = color or C_SECTION
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font  = fnt(bold=True, color=C_WHITE, size=10)
    c.fill  = fll(bg)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 18
    return row + 1

def write_year_hdr(ws, row, years, ncols, label="Fiscal Year Ending"):
    c = ws.cell(row=row, column=1, value=label)
    c.font  = fnt(bold=True, size=10)
    c.fill  = fll(C_SUBTOTAL)
    c.border = brd()
    c.alignment = Alignment(horizontal="left", indent=1)
    for i, yr in enumerate(years):
        cell = ws.cell(row=row, column=i+2, value=yr)
        cell.font  = fnt(bold=True, size=10)
        cell.fill  = fll(C_SUBTOTAL)
        cell.alignment = Alignment(horizontal="right")
        cell.border = brd()
    ws.row_dimensions[row].height = 18
    return row + 1

def write_data_row(ws, row, label, values, years,
                   bold=False, bg=None, indent=0,
                   is_pct=False, is_ratio=False, is_days=False, color=None):
    bg = bg or C_WHITE
    tc = color or (C_BLUE if not is_pct and not is_ratio and not is_days else C_BLACK)
    c = ws.cell(row=row, column=1, value=label)
    c.font  = fnt(bold=bold, size=10)
    c.fill  = fll(bg)
    c.border = brd()
    c.alignment = Alignment(horizontal="left", indent=1+indent)
    for i in range(len(years)):
        cell = ws.cell(row=row, column=i+2)
        cell.value = values[i] if i < len(values) else None
        cell.font  = fnt(bold=bold, color=tc, size=10)
        cell.fill  = fll(bg)
        cell.border = brd()
        cell.alignment = Alignment(horizontal="right")
        if is_pct:   pct_fmt(cell)
        elif is_ratio: ratio_fmt(cell)
        elif is_days:  days_fmt(cell)
        else:          num_fmt(cell)
    return row + 1

def write_formula_row(ws, row, label, formula_fn, n_years,
                      bold=False, bg=None, indent=0,
                      is_pct=False, is_ratio=False, is_days=False):
    bg = bg or C_WHITE
    c = ws.cell(row=row, column=1, value=label)
    c.font  = fnt(bold=bold, size=10)
    c.fill  = fll(bg)
    c.border = brd()
    c.alignment = Alignment(horizontal="left", indent=1+indent)
    for i in range(n_years):
        col = i+2
        cell = ws.cell(row=row, column=col)
        cell.value = formula_fn(row, col)
        cell.font  = fnt(bold=bold, color=C_BLACK, size=10)
        cell.fill  = fll(bg)
        cell.border = brd()
        cell.alignment = Alignment(horizontal="right")
        if is_pct:    pct_fmt(cell)
        elif is_ratio: ratio_fmt(cell)
        elif is_days:  days_fmt(cell)
        else:          num_fmt(cell)
    return row + 1

def patch_formula_cells(ws, target_row, n_years, formula_fn,
                        bold=False, bg=None,
                        is_pct=False, is_ratio=False, is_days=False):
    """
    Overwrite data cells in target_row with formulas (fix-after pattern).
    Also corrects font color to black so cells look like formula outputs.
    """
    bg = bg or C_WHITE
    for i in range(n_years):
        col = i + 2
        cell = ws.cell(row=target_row, column=col)
        cell.value = formula_fn(target_row, col)
        cell.font  = fnt(bold=bold, color=C_BLACK, size=10)
        cell.fill  = fll(bg)
        cell.border = brd()
        cell.alignment = Alignment(horizontal="right")
        if is_pct:    pct_fmt(cell)
        elif is_ratio: ratio_fmt(cell)
        elif is_days:  days_fmt(cell)
        else:          num_fmt(cell)

def blank_row(ws, row, ncols):
    ws.row_dimensions[row].height = 6
    return row + 1

# ═══════════════════════════════════════════════════════════════════════════════
# P&L TAB
# v4 changes:
#   1. Net Interest formula fixed: Interest Income − Interest Expense
#      (positive = net earner, e.g. NVDA cash-rich)
#   2. Added "Other Non-Operating Income / (Expenses)" line in summary
#      between Net Interest and EBT, formula = EBT − EBIT − Net Interest
# ═══════════════════════════════════════════════════════════════════════════════
def build_pl(wb, data, years, ticker):
    ws = wb.create_sheet("P&L")
    n  = len(years)
    nc = n + 1
    setup_ws(ws, years)

    row = write_tab_title(ws, 1,
        f"{ticker} — Income Statement ($mm)",
        nc, subtitle="All figures in USD millions. Blue = source data, Black = formula.")
    row = write_year_hdr(ws, row, years, nc)

    def v(key): return [gm(d, key) for d in data]

    # ── SUMMARY ───────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "SUMMARY — INCOME STATEMENT", nc, C_SUMMARY_HD)

    rev_row = row
    row = write_data_row(ws, row, "(1)  Revenue", v("revenue"), years, bold=True)

    cogs_row = row
    row = write_data_row(ws, row, "(2)  Cost of Revenue (COGS)", v("costOfRevenue"), years)

    gp_row = row
    row = write_formula_row(ws, row, "(3)  Gross Profit", bold=True, bg=C_SUMMARY_BG,
        formula_fn=lambda r,c: f"={cl(c)}{rev_row}-{cl(c)}{cogs_row}", n_years=n)
    row = write_formula_row(ws, row, "     Gross Margin %", indent=1, is_pct=True,
        formula_fn=lambda r,c: f"=IFERROR({cl(c)}{gp_row}/{cl(c)}{rev_row},\"\")", n_years=n)

    sga_row = row
    row = write_data_row(ws, row, "(4)  SG&A", v("sellingGeneralAndAdministrativeExpenses"), years)

    # Other OPEX placeholder — fixed after opex_row is known
    _other_opex_r = row
    row = write_formula_row(ws, row, "(5)  Other OPEX (ex-SG&A)",
        lambda r,c: '=""', n_years=n)

    opex_row = row
    row = write_data_row(ws, row, "     Total Operating Expenses", v("operatingExpenses"), years, indent=1)

    # Fix Other OPEX formula
    patch_formula_cells(ws, _other_opex_r, n,
        lambda r,c: f"=IFERROR({cl(c)}{opex_row}-{cl(c)}{sga_row},\"\")")

    # v4 FIX: EBITDA as formula = EBIT + D&A (pure GAAP EBITDA).
    # FMP's ebitda field adds back SBC and other non-cash items (Adjusted EBITDA).
    # Using formula ensures internal consistency. FMP's adjusted figure is in the Detail section.
    ebitda_row = row
    row = write_formula_row(ws, row, "(6)  EBITDA", bold=True, bg=C_SUMMARY_BG,
        formula_fn=lambda r,c: '=""', n_years=n)   # placeholder — fixed after ebit_row
    row = write_formula_row(ws, row, "     EBITDA Margin %", indent=1, is_pct=True,
        formula_fn=lambda r,c: f"=IFERROR({cl(c)}{ebitda_row}/{cl(c)}{rev_row},\"\")", n_years=n)

    da_row = row
    row = write_data_row(ws, row, "(7)  Depreciation & Amortisation", v("depreciationAndAmortization"), years)

    ebit_row = row
    row = write_data_row(ws, row, "(8)  EBIT (Operating Income)", v("operatingIncome"), years, bold=True, bg=C_SUMMARY_BG)
    row = write_formula_row(ws, row, "     EBIT Margin %", indent=1, is_pct=True,
        formula_fn=lambda r,c: f"=IFERROR({cl(c)}{ebit_row}/{cl(c)}{rev_row},\"\")", n_years=n)

    # Fix EBITDA = EBIT + D&A
    patch_formula_cells(ws, ebitda_row, n,
        lambda r,c: f"={cl(c)}{ebit_row}+{cl(c)}{da_row}",
        bold=True, bg=C_SUMMARY_BG)

    int_exp_row = row
    row = write_data_row(ws, row, "     Interest Expense", v("interestExpense"), years, indent=1)
    int_inc_row = row
    row = write_data_row(ws, row, "     Interest Income", v("interestIncome"), years, indent=1)

    # v4 FIX: Interest Income (+) minus Interest Expense
    # Positive = net earner (e.g. cash-rich companies like NVDA)
    net_int_row = row
    row = write_formula_row(ws, row, "(9)  Net Interest Income / (Expense)",
        formula_fn=lambda r,c: f"=IFERROR({cl(c)}{int_inc_row}-{cl(c)}{int_exp_row},\"\")", n_years=n)

    # v4 NEW: Other Non-Operating Income / (Expenses) placeholder — fixed after ebt_row
    _other_no_r = row
    row = write_formula_row(ws, row, "     Other Non-Operating Income / (Expenses)",
        lambda r,c: '=""', n_years=n, indent=1)

    ebt_row = row
    row = write_data_row(ws, row, "(10) EBT / NPBT", v("incomeBeforeTax"), years, bold=True, bg=C_SUMMARY_BG)

    # Fix Other Non-Operating: EBT − EBIT − Net Interest
    patch_formula_cells(ws, _other_no_r, n,
        lambda r,c: f"=IFERROR({cl(c)}{ebt_row}-{cl(c)}{ebit_row}-{cl(c)}{net_int_row},\"\")")

    tax_row = row
    row = write_data_row(ws, row, "(11) Income Tax Expense", v("incomeTaxExpense"), years)
    row = write_formula_row(ws, row, "     Effective Tax Rate %", indent=1, is_pct=True,
        formula_fn=lambda r,c: f"=IFERROR({cl(c)}{tax_row}/{cl(c)}{ebt_row},\"\")", n_years=n)

    ni_row = row
    row = write_data_row(ws, row, "(12) Net Income / NPAT", v("netIncome"), years, bold=True, bg=C_SUMMARY_BG)
    row = write_formula_row(ws, row, "     Net Margin %", indent=1, is_pct=True,
        formula_fn=lambda r,c: f"=IFERROR({cl(c)}{ni_row}/{cl(c)}{rev_row},\"\")", n_years=n)

    row = blank_row(ws, row, nc)

    # ── DETAIL ────────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "DETAIL — ALL LINE ITEMS FROM FMP", nc, C_DETAIL_HD)
    row = write_year_hdr(ws, row, years, nc)

    row = write_section_hdr(ws, row, "Revenue & Cost", nc)
    row = write_data_row(ws, row, "Revenue",                            v("revenue"),                                    years, bold=True)
    row = write_data_row(ws, row, "Cost of Revenue",                    v("costOfRevenue"),                              years)
    row = write_data_row(ws, row, "Gross Profit",                       v("grossProfit"),                                years, bold=True, bg=C_ALT)

    row = write_section_hdr(ws, row, "Operating Expenses", nc)
    row = write_data_row(ws, row, "R&D Expenses",                       v("researchAndDevelopmentExpenses"),             years)
    row = write_data_row(ws, row, "General & Administrative",           v("generalAndAdministrativeExpenses"),           years)
    row = write_data_row(ws, row, "Selling & Marketing",                v("sellingAndMarketingExpenses"),                years)
    row = write_data_row(ws, row, "SG&A (Combined)",                    v("sellingGeneralAndAdministrativeExpenses"),    years)
    row = write_data_row(ws, row, "Other Expenses",                     v("otherExpenses"),                              years)
    row = write_data_row(ws, row, "Total Operating Expenses",           v("operatingExpenses"),                          years, bold=True, bg=C_ALT)
    row = write_data_row(ws, row, "Cost & Expenses (COGS + Opex)",      v("costAndExpenses"),                            years)

    row = write_section_hdr(ws, row, "Operating & EBITDA", nc)
    row = write_data_row(ws, row, "EBIT (Operating Income)",            v("operatingIncome"),                            years, bold=True)
    row = write_data_row(ws, row, "EBIT (FMP field)",                   v("ebit"),                                       years)
    row = write_data_row(ws, row, "EBITDA",                             v("ebitda"),                                     years, bold=True, bg=C_ALT)
    row = write_data_row(ws, row, "Depreciation & Amortisation",        v("depreciationAndAmortization"),                years)

    row = write_section_hdr(ws, row, "Below the Line", nc)
    row = write_data_row(ws, row, "Interest Income",                    v("interestIncome"),                             years)
    row = write_data_row(ws, row, "Interest Expense",                   v("interestExpense"),                            years)
    row = write_data_row(ws, row, "Net Interest Income",                v("netInterestIncome"),                          years)
    row = write_data_row(ws, row, "Non-Operating Income (ex-interest)", v("nonOperatingIncomeExcludingInterest"),        years)
    total_other_row = row
    row = write_data_row(ws, row, "Total Other Income / (Expenses)",    v("totalOtherIncomeExpensesNet"),                years)
    row = write_data_row(ws, row, "EBT / Income Before Tax",            v("incomeBeforeTax"),                            years, bold=True, bg=C_ALT)
    row = write_data_row(ws, row, "Income Tax Expense",                 v("incomeTaxExpense"),                           years)

    row = write_section_hdr(ws, row, "Net Income", nc)
    row = write_data_row(ws, row, "Net Income from Continuing Ops",     v("netIncomeFromContinuingOperations"),          years)
    row = write_data_row(ws, row, "Net Income from Discontinued Ops",   v("netIncomeFromDiscontinuedOperations"),        years)
    row = write_data_row(ws, row, "Other Adjustments to Net Income",    v("otherAdjustmentsToNetIncome"),                years)
    row = write_data_row(ws, row, "Net Income Deductions",              v("netIncomeDeductions"),                        years)
    row = write_data_row(ws, row, "Net Income",                         v("netIncome"),                                  years, bold=True, bg=C_ALT)
    row = write_data_row(ws, row, "Bottom Line Net Income",             v("bottomLineNetIncome"),                        years, bold=True)

    row = write_section_hdr(ws, row, "Per Share", nc)
    row = write_data_row(ws, row, "EPS (Basic)",                        [g(d,"eps") for d in data],                      years)
    row = write_data_row(ws, row, "EPS (Diluted)",                      [g(d,"epsdiluted") for d in data],               years)
    row = write_data_row(ws, row, "Shares Outstanding — Basic (mm)",    [gm(d,"weightedAverageShsOut") for d in data],   years)
    row = write_data_row(ws, row, "Shares Outstanding — Diluted (mm)",  [gm(d,"weightedAverageShsOutDil") for d in data],years)

    row = write_section_hdr(ws, row, "Metadata", nc)
    row = write_data_row(ws, row, "Reported Currency",                  [data[i].get("reportedCurrency","") for i in range(min(n,len(data)))], years)

    return {"rev": rev_row, "cogs": cogs_row, "gp": gp_row,
            "sga": sga_row, "opex": opex_row,
            "ebitda": ebitda_row, "da": da_row,
            "ebit": ebit_row, "int_exp": int_exp_row, "int_inc": int_inc_row,
            "net_int": net_int_row,
            "ebt": ebt_row, "tax": tax_row, "ni": ni_row}

# ═══════════════════════════════════════════════════════════════════════════════
# BALANCE SHEET TAB
# v4 changes (summary):
#   3.  (4)  Other Current Assets  → formula plug: TCA − Cash − Rec − Inv
#   4.  (9)  Other LT Assets       → formula plug: TLTA − PPE − Goodwill − DTA
#   5.  (14) Short-Term Leases     → uses correct FMP current-lease field
#   6.  (15) Other Current Liabs   → formula plug: TCL − AP − STDebt − STLeases
#   7.  (18) Long-Term Leases      → uses correct FMP LT-lease field
#   8.  (19) Other LT Liabilities  → formula plug: TL − TCL − LTDebt − LTLeases
#   9.  (23) Other Equity          → formula plug: TE − CommonStock − RetainedEarnings
#   10. (25) Total L&E             → formula: TL + TE
# v4 changes (detail):
#   11. Added "Accrued & Other Current Liabilities" to current liabilities section
#   12. Moved "Minority Interest" from Non-Current Liabilities → Shareholders' Equity
#   13. Added "Long-Term Operating Lease Liabilities" in non-current liabilities
#   14. Total Non-Current Liabilities → formula sum of components
#   15. Total Stockholders' Equity (detail) → formula sum of components
# ═══════════════════════════════════════════════════════════════════════════════
def build_bs(wb, data, years, ticker):
    ws = wb.create_sheet("Balance Sheet")
    n  = len(years)
    nc = n + 1
    setup_ws(ws, years)

    row = write_tab_title(ws, 1, f"{ticker} — Balance Sheet ($mm)", nc,
        subtitle="All figures in USD millions. Blue = source data, Black = formula.")
    row = write_year_hdr(ws, row, years, nc)

    def v(key): return [gm(d, key) for d in data]

    # Lease field helpers: try the most specific FMP fields first,
    # then fall back to broader fields.  Returns $mm list.
    def v_st_leases():
        return [gm_any(d,
            "shortTermCapitalLeaseObligation",   # FMP newer naming
            "currentPortionLeaseLiabilities",
            "shortTermLeaseLiabilities",
            "operatingLeaseLiabilityCurrentPortion",
        ) for d in data]

    def v_lt_leases():
        return [gm_any(d,
            "longTermCapitalLeaseObligation",    # FMP newer naming
            "longTermLeaseLiabilities",
            "operatingLeaseLiabilityNoncurrentPortion",
            "operatingLeaseLiabilityNonCurrent",
        ) for d in data]

    # ── SUMMARY ───────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "SUMMARY — BALANCE SHEET", nc, C_SUMMARY_HD)

    cash_row = row
    row = write_data_row(ws, row, "(1)  Cash & Cash Equivalents",      v("cashAndCashEquivalents"),    years)
    rec_row  = row
    row = write_data_row(ws, row, "(2)  Receivables",                  v("netReceivables"),            years)
    inv_row  = row
    row = write_data_row(ws, row, "(3)  Inventory",                    v("inventory"),                 years)

    # v4: (4) Other Current Assets = TCA − Cash − Rec − Inv (plug)
    _oca_r = row
    row = write_formula_row(ws, row, "(4)  Other Current Assets",
        lambda r,c: '=""', n_years=n)

    tca_row  = row
    row = write_data_row(ws, row, "(5)  Total Current Assets",         v("totalCurrentAssets"),        years, bold=True, bg=C_SUMMARY_BG)

    # Fix OCA plug now that tca_row is known
    patch_formula_cells(ws, _oca_r, n,
        lambda r,c: f"=IFERROR({cl(c)}{tca_row}-{cl(c)}{cash_row}-{cl(c)}{rec_row}-{cl(c)}{inv_row},\"\")")

    ppe_row  = row
    row = write_data_row(ws, row, "(6)  PP&E (Net)",                   v("propertyPlantEquipmentNet"), years)
    gw_row   = row
    row = write_data_row(ws, row, "(7)  Goodwill",                     v("goodwill"),                  years)
    dta_row  = row
    row = write_data_row(ws, row, "(8)  Deferred Tax Assets",          v("taxAssets"),                 years)

    # v4: (9) Other LT Assets = TLTA − PPE − Goodwill − DTA (plug)
    _olta_r = row
    row = write_formula_row(ws, row, "(9)  Other LT Assets",
        lambda r,c: '=""', n_years=n)

    tlta_row = row
    row = write_data_row(ws, row, "(10) Total LT Assets",              v("totalNonCurrentAssets"),     years, bold=True, bg=C_SUMMARY_BG)

    # Fix Other LT Assets plug
    patch_formula_cells(ws, _olta_r, n,
        lambda r,c: f"=IFERROR({cl(c)}{tlta_row}-{cl(c)}{ppe_row}-{cl(c)}{gw_row}-{cl(c)}{dta_row},\"\")")

    tot_assets_row = row
    row = write_data_row(ws, row, "(11) Total Assets",                 v("totalAssets"),               years, bold=True, bg=C_SUMMARY_BG)

    row = blank_row(ws, row, nc)

    ap_row   = row
    row = write_data_row(ws, row, "(12) Accounts Payable",             v("accountPayables"),           years)
    std_row  = row
    row = write_data_row(ws, row, "(13) Short-Term Borrowings",        v("shortTermDebt"),             years)

    # v4: Short-Term Leases uses proper current-portion field
    stl_row  = row
    row = write_data_row(ws, row, "(14) Short-Term Leases",            v_st_leases(),                  years)

    # v4: (15) Other Current Liabilities = TCL − AP − STDebt − STLeases (plug)
    _ocl_r = row
    row = write_formula_row(ws, row, "(15) Other Current Liabilities",
        lambda r,c: '=""', n_years=n)

    tcl_row  = row
    row = write_data_row(ws, row, "(16) Total Current Liabilities",    v("totalCurrentLiabilities"),   years, bold=True, bg=C_SUMMARY_BG)

    # Fix Other CL plug
    patch_formula_cells(ws, _ocl_r, n,
        lambda r,c: f"=IFERROR({cl(c)}{tcl_row}-{cl(c)}{ap_row}-{cl(c)}{std_row}-{cl(c)}{stl_row},\"\")")

    ltd_row  = row
    row = write_data_row(ws, row, "(17) Long-Term Debt",               v("longTermDebt"),              years)

    # v4: Long-Term Leases uses proper LT-lease field
    ltl_row  = row
    row = write_data_row(ws, row, "(18) Long-Term Leases",             v_lt_leases(),                  years)

    # v4: (19) Other LT Liabilities = TL − TCL − LTDebt − LTLeases (plug)
    _oltl_r = row
    row = write_formula_row(ws, row, "(19) Other LT Liabilities",
        lambda r,c: '=""', n_years=n)

    tl_row   = row
    row = write_data_row(ws, row, "(20) Total Liabilities",            v("totalLiabilities"),          years, bold=True, bg=C_SUMMARY_BG)

    # Fix Other LT Liabilities plug
    patch_formula_cells(ws, _oltl_r, n,
        lambda r,c: f"=IFERROR({cl(c)}{tl_row}-{cl(c)}{tcl_row}-{cl(c)}{ltd_row}-{cl(c)}{ltl_row},\"\")")

    row = blank_row(ws, row, nc)

    cs_row   = row
    row = write_data_row(ws, row, "(21) Common Stock & APIC",          v("commonStock"),               years)
    re_row   = row
    row = write_data_row(ws, row, "(22) Retained Earnings",            v("retainedEarnings"),          years)

    # v4: (23) Other Equity = TE − Common Stock − Retained Earnings (plug)
    _oe_r = row
    row = write_formula_row(ws, row, "(23) Other Equity",
        lambda r,c: '=""', n_years=n)

    te_row   = row
    row = write_data_row(ws, row, "(24) Total Equity",                 v("totalStockholdersEquity"),   years, bold=True, bg=C_SUMMARY_BG)

    # Fix Other Equity plug
    patch_formula_cells(ws, _oe_r, n,
        lambda r,c: f"=IFERROR({cl(c)}{te_row}-{cl(c)}{cs_row}-{cl(c)}{re_row},\"\")")

    # v4: (25) Total Liabilities & Equity = TL + TE (formula, not raw data)
    tle_row  = row
    row = write_formula_row(ws, row, "(25) Total Liabilities & Equity",
        formula_fn=lambda r,c: f"={cl(c)}{tl_row}+{cl(c)}{te_row}",
        n_years=n, bold=True, bg=C_SUMMARY_BG)

    row = blank_row(ws, row, nc)

    # ── DETAIL ────────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "DETAIL — ALL LINE ITEMS FROM FMP", nc, C_DETAIL_HD)
    row = write_year_hdr(ws, row, years, nc)

    row = write_section_hdr(ws, row, "Current Assets", nc)
    row = write_data_row(ws, row, "Cash & Cash Equivalents",               v("cashAndCashEquivalents"),           years)
    row = write_data_row(ws, row, "Short-Term Investments",                v("shortTermInvestments"),             years)
    row = write_data_row(ws, row, "Cash & Short-Term Investments",         v("cashAndShortTermInvestments"),      years, bg=C_ALT)
    row = write_data_row(ws, row, "Accounts Receivable (Trade)",           v("accountsReceivables"),              years)
    row = write_data_row(ws, row, "Other Receivables",                     v("otherReceivables"),                 years)
    row = write_data_row(ws, row, "Net Receivables (Total)",               v("netReceivables"),                   years, bg=C_ALT)
    row = write_data_row(ws, row, "Inventory",                             v("inventory"),                        years)
    row = write_data_row(ws, row, "Prepaids",                              v("prepaids"),                         years)
    row = write_data_row(ws, row, "Other Current Assets",                  v("otherCurrentAssets"),               years)
    row = write_data_row(ws, row, "Total Current Assets",                  v("totalCurrentAssets"),               years, bold=True, bg=C_ALT)

    row = write_section_hdr(ws, row, "Non-Current Assets", nc)
    row = write_data_row(ws, row, "PP&E (Net)",                            v("propertyPlantEquipmentNet"),        years)
    row = write_data_row(ws, row, "Goodwill",                              v("goodwill"),                         years)
    row = write_data_row(ws, row, "Intangible Assets",                     v("intangibleAssets"),                 years)
    row = write_data_row(ws, row, "Long-Term Investments",                 v("longTermInvestments"),              years)
    row = write_data_row(ws, row, "Tax Assets (Deferred)",                 v("taxAssets"),                        years)
    row = write_data_row(ws, row, "Total Investments",                     v("totalInvestments"),                 years)
    row = write_data_row(ws, row, "Other Non-Current Assets",              v("otherNonCurrentAssets"),            years)
    row = write_data_row(ws, row, "Total Non-Current Assets",              v("totalNonCurrentAssets"),            years, bold=True, bg=C_ALT)
    row = write_data_row(ws, row, "TOTAL ASSETS",                          v("totalAssets"),                      years, bold=True, bg=C_SUBTOTAL)

    # v4 DETAIL — Current Liabilities
    row = write_section_hdr(ws, row, "Current Liabilities", nc)
    d_ap_r = row
    row = write_data_row(ws, row, "Accounts Payable",                      v("accountPayables"),                  years)
    d_std_r = row
    row = write_data_row(ws, row, "Short-Term Debt",                       v("shortTermDebt"),                    years)
    d_stl_r = row
    row = write_data_row(ws, row, "Short-Term Lease Liabilities",          v_st_leases(),                         years)
    d_drev_cur_r = row
    row = write_data_row(ws, row, "Deferred Revenue (Current)",            v("deferredRevenue"),                  years)
    d_accrued_r = row
    row = write_data_row(ws, row, "Accrued & Other Current Liabilities",
        [gm_any(d,
            "accruedLiabilities",
            "accruedAndOtherCurrentLiabilities",
            "otherCurrentLiabilities",
        ) for d in data], years)
    # Other Current Liabilities = plug: Total CL − AP − ST Debt − ST Leases − Def Rev − Accrued
    _d_ocl_r = row
    row = write_formula_row(ws, row, "Other Current Liabilities",
        lambda r,c: '=""', n_years=n)
    d_tcl_detail_r = row
    row = write_data_row(ws, row, "Total Current Liabilities",             v("totalCurrentLiabilities"),          years, bold=True, bg=C_ALT)
    # Fix Other CL plug
    patch_formula_cells(ws, _d_ocl_r, n,
        lambda r,c: (
            f"=IFERROR({cl(c)}{d_tcl_detail_r}"
            f"-{cl(c)}{d_ap_r}-{cl(c)}{d_std_r}-{cl(c)}{d_stl_r}"
            f"-{cl(c)}{d_drev_cur_r}-{cl(c)}{d_accrued_r},\"\")"
        ))

    # v4 DETAIL — Non-Current Liabilities
    row = write_section_hdr(ws, row, "Non-Current Liabilities", nc)
    d_ltd_row = row
    row = write_data_row(ws, row, "Long-Term Debt",                        v("longTermDebt"),                     years)
    d_ltl_row = row
    row = write_data_row(ws, row, "Long-Term Operating Lease Liabilities", v_lt_leases(),                         years)
    d_drev_nc_r = row
    row = write_data_row(ws, row, "Deferred Revenue (Non-Current)",        v("deferredRevenueNonCurrent"),        years)
    d_dtl_row  = row
    row = write_data_row(ws, row, "Deferred Tax Liabilities",              v("deferredTaxLiabilitiesNonCurrent"), years)
    d_oltl_row = row
    row = write_data_row(ws, row, "Other Non-Current Liabilities",         v("otherNonCurrentLiabilities"),       years)
    d_tncl_row = row
    row = write_formula_row(ws, row, "Total Non-Current Liabilities",
        formula_fn=lambda r,c: (
            f"=IFERROR({cl(c)}{d_ltd_row}+{cl(c)}{d_ltl_row}"
            f"+{cl(c)}{d_drev_nc_r}+{cl(c)}{d_dtl_row}+{cl(c)}{d_oltl_row},\"\")"
        ), n_years=n, bold=True, bg=C_ALT)
    row = write_data_row(ws, row, "TOTAL LIABILITIES",                     v("totalLiabilities"),                 years, bold=True, bg=C_SUBTOTAL)

    # v4 DETAIL — Shareholders' Equity
    row = write_section_hdr(ws, row, "Shareholders' Equity", nc)
    d_cs_row = row
    row = write_data_row(ws, row, "Common Stock & APIC",                   v("commonStock"),                      years)
    d_re_row = row
    row = write_data_row(ws, row, "Retained Earnings",                     v("retainedEarnings"),                 years)
    # Other Total SE = plug (fix-after d_te_inc_min is written)
    _d_ose_r = row
    row = write_formula_row(ws, row, "Other Total Stockholders Equity",
        lambda r,c: '=""', n_years=n)
    # Total Stockholders' Equity = formula sum of components
    d_tse_row = row
    row = write_formula_row(ws, row, "Total Stockholders Equity",
        formula_fn=lambda r,c: (
            f"=IFERROR({cl(c)}{d_cs_row}+{cl(c)}{d_re_row}+{cl(c)}{_d_ose_r},\"\")"
        ), n_years=n, bold=True, bg=C_ALT)
    d_mi_row = row
    row = write_data_row(ws, row, "Minority Interest",                     v("minorityInterest"),                 years)
    d_te_inc_min_row = row
    row = write_data_row(ws, row, "Total Equity (inc. Minority)",          v("totalEquity"),                      years, bold=True, bg=C_ALT)
    # Fix Other Total SE plug = Total Equity (inc. minority) − Common − Retained − Minority
    patch_formula_cells(ws, _d_ose_r, n,
        lambda r,c: (
            f"=IFERROR({cl(c)}{d_te_inc_min_row}"
            f"-{cl(c)}{d_cs_row}-{cl(c)}{d_re_row}-{cl(c)}{d_mi_row},\"\")"
        ))

    row = write_section_hdr(ws, row, "Key Derived Balances", nc)
    row = write_data_row(ws, row, "Total Debt (ST + LT)",                  v("totalDebt"),                        years, bold=True)
    row = write_data_row(ws, row, "Net Debt",                              v("netDebt"),                          years, bold=True)
    # Working Capital = formula: Total CA − Total CL (from summary rows, same sheet)
    row = write_formula_row(ws, row, "Working Capital",
        formula_fn=lambda r,c: f"=IFERROR({cl(c)}{tca_row}-{cl(c)}{tcl_row},\"\")",
        n_years=n, bold=True)
    # Total L&E = reference the summary Total L&E formula row
    row = write_formula_row(ws, row, "TOTAL LIABILITIES & EQUITY",
        formula_fn=lambda r,c: f"={cl(c)}{tle_row}",
        n_years=n, bold=True, bg=C_SUBTOTAL)

    return {"cash": cash_row, "rec": rec_row, "inv": inv_row, "tca": tca_row,
            "ppe": ppe_row, "dta": dta_row, "tlta": tlta_row, "tot_assets": tot_assets_row,
            "ap": ap_row, "tcl": tcl_row, "ltd": ltd_row,
            "tl": tl_row, "te": te_row}

# ═══════════════════════════════════════════════════════════════════════════════
# CASH FLOW TAB
# v4 changes:
#   16. Added "Other Investing Activities" in detail investing section
#       so the sub-items reconcile to total CFI
#   17. Added "Other Financing Activities" in detail financing section
#       so the sub-items reconcile to total CFF
# ═══════════════════════════════════════════════════════════════════════════════
def build_cf(wb, data, years, ticker):
    ws = wb.create_sheet("Cash Flow")
    n  = len(years)
    nc = n + 1
    setup_ws(ws, years)

    row = write_tab_title(ws, 1, f"{ticker} — Cash Flow Statement ($mm)", nc,
        subtitle="All figures in USD millions. Blue = source data, Black = formula.")
    row = write_year_hdr(ws, row, years, nc)

    def v(key): return [gm(d, key) for d in data]

    # ── SUMMARY ───────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "SUMMARY — CASH FLOW STATEMENT", nc, C_SUMMARY_HD)

    cfo_row   = row
    row = write_data_row(ws, row, "(1)  Net Cash from Operations (CFO)", v("netCashProvidedByOperatingActivities"), years, bold=True, bg=C_SUMMARY_BG)

    capex_row = row
    row = write_data_row(ws, row, "(2)  Capital Expenditures",            v("capitalExpenditure"),                  years)

    # Summary CFI and CFF: placeholders — fixed after detail totals are computed as formula sums
    _cfi_summary = row
    cfi_row = row
    row = write_formula_row(ws, row, "(3)  Net Cash from Investing (CFI)",
        lambda r,c: '=""', n_years=n, bold=True, bg=C_SUMMARY_BG)

    row = blank_row(ws, row, nc)

    draw_row  = row
    row = write_data_row(ws, row, "(4)  Debt Drawdowns",                  v("netDebtIssuance"),                     years)
    rep_row   = row
    row = write_data_row(ws, row, "(5)  Debt Repayments",
        [gm_any(d, "debtRepayment", "repaymentOfDebt", "longTermDebtRepayment") for d in data], years)
    iss_row   = row
    row = write_data_row(ws, row, "(6)  Issuance of Common Stock",        v("commonStockIssuance"),                 years)
    div_row   = row
    row = write_data_row(ws, row, "(7)  Dividends Paid",
        [gm_any(d, "dividendsPaid", "commonDividendsPaid", "paymentOfDividends") for d in data], years)
    _cff_summary = row
    cff_row = row
    row = write_formula_row(ws, row, "(8)  Net Cash from Financing (CFF)",
        lambda r,c: '=""', n_years=n, bold=True, bg=C_SUMMARY_BG)

    row = blank_row(ws, row, nc)
    fcf_row   = row
    row = write_data_row(ws, row, "     Free Cash Flow (FCF)",            v("freeCashFlow"),                        years, bold=True, bg=C_SUMMARY_BG)
    row = write_data_row(ws, row, "     Net Change in Cash",              v("netChangeInCash"),                     years, bold=True)

    row = blank_row(ws, row, nc)

    # ── DETAIL ────────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "DETAIL — ALL LINE ITEMS FROM FMP", nc, C_DETAIL_HD)
    row = write_year_hdr(ws, row, years, nc)

    row = write_section_hdr(ws, row, "Operating Activities", nc)
    row = write_data_row(ws, row, "Net Income",                           v("netIncome"),                           years)
    row = write_data_row(ws, row, "Depreciation & Amortisation",          v("depreciationAndAmortization"),         years)
    row = write_data_row(ws, row, "Deferred Income Tax",                  v("deferredIncomeTax"),                   years)
    row = write_data_row(ws, row, "Stock-Based Compensation",             v("stockBasedCompensation"),              years)
    row = write_data_row(ws, row, "Change in Working Capital",            v("changeInWorkingCapital"),              years)
    row = write_data_row(ws, row, "  — Accounts Receivable",              v("accountsReceivables"),                 years, indent=1)
    row = write_data_row(ws, row, "  — Inventory",                        v("inventory"),                           years, indent=1)
    row = write_data_row(ws, row, "  — Accounts Payable",                 v("accountsPayables"),                    years, indent=1)
    row = write_data_row(ws, row, "  — Other Working Capital",            v("otherWorkingCapital"),                 years, indent=1)
    row = write_data_row(ws, row, "Other Non-Cash Items",                 v("otherNonCashItems"),                   years)
    row = write_data_row(ws, row, "Net Cash from Operations (CFO)",       v("netCashProvidedByOperatingActivities"),years, bold=True, bg=C_ALT)

    # Investing detail — track rows for formula sum
    row = write_section_hdr(ws, row, "Investing Activities", nc)
    d_capex_r = row
    row = write_data_row(ws, row, "Capital Expenditures",                 v("capitalExpenditure"),                  years)
    row = write_data_row(ws, row, "  (Alt: Invest. in PP&E)",             v("investmentsInPropertyPlantAndEquipment"), years, indent=1)
    d_acq_r = row
    row = write_data_row(ws, row, "Acquisitions (Net)",                   v("acquisitionsNet"),                     years)
    d_purch_r = row
    row = write_data_row(ws, row, "Purchases of Investments",             v("purchasesOfInvestments"),              years)
    d_sales_r = row
    row = write_data_row(ws, row, "Sales / Maturities of Investments",    v("salesMaturitiesOfInvestments"),        years)
    d_other_inv_r = row
    row = write_data_row(ws, row, "Other Investing Activities",
        [gm_any(d, "otherInvestingActivities", "otherInvestingActivitiesNet") for d in data], years)
    # CFI total = formula sum (Alt PP&E row excluded to avoid double-count with Capex)
    d_cfi_r = row
    row = write_formula_row(ws, row, "Net Cash from Investing (CFI)",
        formula_fn=lambda r,c: (
            f"=IFERROR({cl(c)}{d_capex_r}+{cl(c)}{d_acq_r}"
            f"+{cl(c)}{d_purch_r}+{cl(c)}{d_sales_r}+{cl(c)}{d_other_inv_r},\"\")"
        ), n_years=n, bold=True, bg=C_ALT)

    # Financing detail — track rows for formula sum
    row = write_section_hdr(ws, row, "Financing Activities", nc)
    d_debt_iss_r = row
    row = write_data_row(ws, row, "Debt Issuance (Net)",                  v("netDebtIssuance"),                     years)
    d_debt_rep_r = row
    row = write_data_row(ws, row, "Debt Repayment",
        [gm_any(d, "debtRepayment", "repaymentOfDebt", "longTermDebtRepayment") for d in data], years)
    d_stk_iss_r = row
    row = write_data_row(ws, row, "Common Stock Issuance",                v("commonStockIssuance"),                 years)
    d_buyback_r = row
    row = write_data_row(ws, row, "Common Stock Repurchased (Buybacks)",  v("commonStockRepurchased"),              years)
    d_div_r = row
    row = write_data_row(ws, row, "Dividends Paid",
        [gm_any(d, "dividendsPaid", "commonDividendsPaid", "paymentOfDividends") for d in data], years)
    d_other_fin_r = row
    row = write_data_row(ws, row, "Other Financing Activities",
        [gm_any(d, "otherFinancingActivities", "otherFinancingActivitiesNet") for d in data], years)
    # CFF total = formula sum of all financing line items
    d_cff_r = row
    row = write_formula_row(ws, row, "Net Cash from Financing (CFF)",
        formula_fn=lambda r,c: (
            f"=IFERROR({cl(c)}{d_debt_iss_r}+{cl(c)}{d_debt_rep_r}"
            f"+{cl(c)}{d_stk_iss_r}+{cl(c)}{d_buyback_r}"
            f"+{cl(c)}{d_div_r}+{cl(c)}{d_other_fin_r},\"\")"
        ), n_years=n, bold=True, bg=C_ALT)

    # Patch summary CFI and CFF to reference detail formula totals
    patch_formula_cells(ws, _cfi_summary, n,
        lambda r,c: f"={cl(c)}{d_cfi_r}", bold=True, bg=C_SUMMARY_BG)
    patch_formula_cells(ws, _cff_summary, n,
        lambda r,c: f"={cl(c)}{d_cff_r}", bold=True, bg=C_SUMMARY_BG)

    row = write_section_hdr(ws, row, "Cash Summary", nc)
    row = write_data_row(ws, row, "Effect of Forex on Cash",              v("effectOfForexChangesOnCash"),          years)
    net_change_row = row
    row = write_data_row(ws, row, "Net Change in Cash",                   v("netChangeInCash"),                     years, bold=True)
    row = write_data_row(ws, row, "Cash at Beginning of Period",          v("cashAtBeginningOfPeriod"),             years)
    row = write_data_row(ws, row, "Cash at End of Period",                v("cashAtEndOfPeriod"),                   years, bold=True, bg=C_ALT)
    row = write_data_row(ws, row, "Free Cash Flow",                       v("freeCashFlow"),                        years, bold=True, bg=C_ALT)
    row = write_data_row(ws, row, "Operating Cash Flow per Share",        [g(d,"operatingCashFlowPerShare") for d in data], years)
    row = write_data_row(ws, row, "Free Cash Flow per Share",             [g(d,"freeCashFlowPerShare") for d in data],      years)

    return {"cfo": cfo_row, "capex": capex_row, "cfi": cfi_row,
            "cff": cff_row, "fcf": fcf_row, "div": div_row,
            "net_change": net_change_row}

# ═══════════════════════════════════════════════════════════════════════════════
# RATIOS & FCF BRIDGE TAB  (unchanged from v3)
# ═══════════════════════════════════════════════════════════════════════════════
def build_ratios(wb, is_data, bs_data, cf_data, years, ticker, pl_refs, bs_refs, cf_refs):
    ws = wb.create_sheet("Ratios & FCF")
    n  = len(years)
    nc = n + 1
    setup_ws(ws, years)

    row = write_tab_title(ws, 1, f"{ticker} — Key Ratios & Free Cash Flow Bridge", nc,
        subtitle="All formulas. Black = calculated. Cross-sheet references pull from P&L, Balance Sheet, Cash Flow tabs.")
    row = write_year_hdr(ws, row, years, nc)

    def pl(r, col):  return f"'P&L'!{cl(col)}{r}"
    def bs(r, col):  return f"'Balance Sheet'!{cl(col)}{r}"
    def cf(r, col):  return f"'Cash Flow'!{cl(col)}{r}"

    rev   = pl_refs["rev"];   cogs  = pl_refs["cogs"]; gp    = pl_refs["gp"]
    ebitda= pl_refs["ebitda"];da    = pl_refs["da"];   ebit  = pl_refs["ebit"]
    ebt   = pl_refs["ebt"];   tax   = pl_refs["tax"];  ni    = pl_refs["ni"]

    tca   = bs_refs["tca"];   tcl   = bs_refs["tcl"];  tot_a = bs_refs["tot_assets"]
    te    = bs_refs["te"];    ltd   = bs_refs["ltd"];  cash  = bs_refs["cash"]
    rec   = bs_refs["rec"];   inv   = bs_refs["inv"];  ap    = bs_refs["ap"]

    cfo   = cf_refs["cfo"];   capex = cf_refs["capex"];fcf   = cf_refs["fcf"]

    # ── UNLEVERED FREE CASH FLOW BRIDGE ───────────────────────────────────────
    row = write_section_hdr(ws, row, "UNLEVERED FREE CASH FLOW (UFCF) — STEP-BY-STEP BRIDGE", nc, C_SUMMARY_HD)
    row = write_section_hdr(ws, row, "Note: UFCF = NOPAT + D&A − ΔNWC − Capex  |  Used as input to DCF (unlevered / WACC-based)", nc, "555555")

    ebit_r = row
    row = write_formula_row(ws, row, "EBIT (Operating Income)",
        lambda r,c: f"={pl(ebit, c)}", n, bold=True)

    nopat_r = row
    row = write_formula_row(ws, row, "  (−) Taxes on EBIT  [EBIT × Eff. Tax Rate]",
        lambda r,c: f"=IFERROR(-{cl(c)}{ebit_r}*({pl(tax,c)}/{pl(ebt,c)}),0)",
        n, indent=1)

    nopat_total = row
    row = write_formula_row(ws, row, "NOPAT  (Net Operating Profit After Tax)",
        lambda r,c: f"={cl(c)}{ebit_r}+{cl(c)}{nopat_r}",
        n, bold=True, bg=C_SUMMARY_BG)

    da_r = row
    row = write_formula_row(ws, row, "  (+) Depreciation & Amortisation",
        lambda r,c: f"={pl(da, c)}", n, indent=1)

    ebitda_r = row
    row = write_formula_row(ws, row, "  = EBITDA (cross-check)",
        lambda r,c: f"={cl(c)}{nopat_total}+{cl(c)}{da_r}",
        n, indent=1, bg=C_ALT)

    nwc_r = row
    row = write_formula_row(ws, row, "  (−) Increase in Net Working Capital  [ΔRec + ΔInv − ΔAP]",
        lambda r,c: (
            f"=IFERROR(('Balance Sheet'!{cl(c)}{rec}-'Balance Sheet'!{cl(c-1)}{rec})"
            f"+('Balance Sheet'!{cl(c)}{inv}-'Balance Sheet'!{cl(c-1)}{inv})"
            f"-('Balance Sheet'!{cl(c)}{ap}-'Balance Sheet'!{cl(c-1)}{ap}),0)"
            if c > 2 else "=0"
        ), n, indent=1)

    capex_r = row
    row = write_formula_row(ws, row, "  (−) Capital Expenditures",
        lambda r,c: f"={cf(capex, c)}", n, indent=1)

    ufcf_row = row
    row = write_formula_row(ws, row, "UNLEVERED FREE CASH FLOW (UFCF)",
        lambda r,c: f"={cl(c)}{nopat_total}+{cl(c)}{da_r}-{cl(c)}{nwc_r}+{cl(c)}{capex_r}",
        n, bold=True, bg=C_SUMMARY_BG)
    row = write_formula_row(ws, row, "  UFCF Margin %",
        lambda r,c: f"=IFERROR({cl(c)}{ufcf_row}/{pl(rev,c)},\"\")",
        n, is_pct=True, indent=1)

    row = blank_row(ws, row, nc)

    # ── LEVERED FREE CASH FLOW BRIDGE ─────────────────────────────────────────
    row = write_section_hdr(ws, row, "LEVERED FREE CASH FLOW (LFCF) — STEP-BY-STEP BRIDGE", nc, C_SUMMARY_HD)
    row = write_section_hdr(ws, row, "Note: LFCF = Net Income + D&A − ΔNWC − Capex  |  Represents cash available to equity holders", nc, "555555")

    ni_r = row
    row = write_formula_row(ws, row, "Net Income",
        lambda r,c: f"={pl(ni, c)}", n, bold=True)

    da_lev = row
    row = write_formula_row(ws, row, "  (+) Depreciation & Amortisation",
        lambda r,c: f"={pl(da, c)}", n, indent=1)

    nwc_lev = row
    row = write_formula_row(ws, row, "  (−) Increase in Net Working Capital",
        lambda r,c: (
            f"=IFERROR(('Balance Sheet'!{cl(c)}{rec}-'Balance Sheet'!{cl(c-1)}{rec})"
            f"+('Balance Sheet'!{cl(c)}{inv}-'Balance Sheet'!{cl(c-1)}{inv})"
            f"-('Balance Sheet'!{cl(c)}{ap}-'Balance Sheet'!{cl(c-1)}{ap}),0)"
            if c > 2 else "=0"
        ), n, indent=1)

    capex_lev = row
    row = write_formula_row(ws, row, "  (−) Capital Expenditures",
        lambda r,c: f"={cf(capex, c)}", n, indent=1)

    lfcf_row = row
    row = write_formula_row(ws, row, "LEVERED FREE CASH FLOW (LFCF)",
        lambda r,c: f"={cl(c)}{ni_r}+{cl(c)}{da_lev}-{cl(c)}{nwc_lev}+{cl(c)}{capex_lev}",
        n, bold=True, bg=C_SUMMARY_BG)
    row = write_formula_row(ws, row, "  LFCF Margin %",
        lambda r,c: f"=IFERROR({cl(c)}{lfcf_row}/{pl(rev,c)},\"\")",
        n, is_pct=True, indent=1)
    row = write_formula_row(ws, row, "  FCF Conversion (LFCF / Net Income)",
        lambda r,c: f"=IFERROR({cl(c)}{lfcf_row}/{cl(c)}{ni_r},\"\")",
        n, is_pct=True, indent=1)

    row = blank_row(ws, row, nc)

    # ── PROFITABILITY ─────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "PROFITABILITY RATIOS", nc)
    row = write_formula_row(ws, row, "Gross Margin %",
        lambda r,c: f"=IFERROR({pl(gp,c)}/{pl(rev,c)},\"\")", n, is_pct=True)
    row = write_formula_row(ws, row, "EBITDA Margin %",
        lambda r,c: f"=IFERROR({pl(ebitda,c)}/{pl(rev,c)},\"\")", n, is_pct=True)
    row = write_formula_row(ws, row, "EBIT Margin %",
        lambda r,c: f"=IFERROR({pl(ebit,c)}/{pl(rev,c)},\"\")", n, is_pct=True)
    row = write_formula_row(ws, row, "Net Margin %",
        lambda r,c: f"=IFERROR({pl(ni,c)}/{pl(rev,c)},\"\")", n, is_pct=True)
    row = write_formula_row(ws, row, "Return on Equity (ROE)",
        lambda r,c: f"=IFERROR({pl(ni,c)}/{bs(te,c)},\"\")", n, is_pct=True)
    row = write_formula_row(ws, row, "Return on Assets (ROA)",
        lambda r,c: f"=IFERROR({pl(ni,c)}/{bs(tot_a,c)},\"\")", n, is_pct=True)
    row = write_formula_row(ws, row, "ROIC  [EBIT×(1−t) / (Equity + Net Debt)]",
        lambda r,c: f"=IFERROR(({pl(ebit,c)}*(1-{pl(tax,c)}/{pl(ebt,c)}))/({bs(te,c)}+'Balance Sheet'!{cl(c)}{ltd}-'Balance Sheet'!{cl(c)}{cash}),\"\")",
        n, is_pct=True)

    row = blank_row(ws, row, nc)

    # ── LEVERAGE ──────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "LEVERAGE & CREDIT RATIOS", nc)
    row = write_formula_row(ws, row, "Net Debt / EBITDA",
        lambda r,c: f"=IFERROR(('Balance Sheet'!{cl(c)}{ltd}-'Balance Sheet'!{cl(c)}{cash})/{pl(ebitda,c)},\"\")",
        n, is_ratio=True)
    row = write_formula_row(ws, row, "Total Debt / EBITDA",
        lambda r,c: f"=IFERROR('Balance Sheet'!{cl(c)}{ltd}/{pl(ebitda,c)},\"\")",
        n, is_ratio=True)
    row = write_formula_row(ws, row, "Interest Coverage  (EBIT / Interest Expense)",
        lambda r,c: f"=IFERROR({pl(ebit,c)}/ABS({pl(ebt,c)}-{pl(ebit,c)}),\"\")",
        n, is_ratio=True)
    row = write_formula_row(ws, row, "Debt / Equity",
        lambda r,c: f"=IFERROR('Balance Sheet'!{cl(c)}{ltd}/{bs(te,c)},\"\")",
        n, is_ratio=True)
    row = write_formula_row(ws, row, "Total Debt / Total Assets",
        lambda r,c: f"=IFERROR('Balance Sheet'!{cl(c)}{ltd}/{bs(tot_a,c)},\"\")",
        n, is_pct=True)

    row = blank_row(ws, row, nc)

    # ── LIQUIDITY ─────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "LIQUIDITY RATIOS", nc)
    row = write_formula_row(ws, row, "Current Ratio  (CA / CL)",
        lambda r,c: f"=IFERROR({bs(tca,c)}/{bs(tcl,c)},\"\")", n, is_ratio=True)
    row = write_formula_row(ws, row, "Quick Ratio  (CA − Inventory) / CL",
        lambda r,c: f"=IFERROR(({bs(tca,c)}-'Balance Sheet'!{cl(c)}{inv})/{bs(tcl,c)},\"\")",
        n, is_ratio=True)
    row = write_formula_row(ws, row, "Cash Ratio  (Cash / CL)",
        lambda r,c: f"=IFERROR({bs(cash,c)}/{bs(tcl,c)},\"\")", n, is_ratio=True)

    row = blank_row(ws, row, nc)

    # ── EFFICIENCY ────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "EFFICIENCY & WORKING CAPITAL", nc)
    row = write_formula_row(ws, row, "Asset Turnover  (Revenue / Assets)",
        lambda r,c: f"=IFERROR({pl(rev,c)}/{bs(tot_a,c)},\"\")", n, is_ratio=True)
    # Average balance helpers:
    # For cols 3..n+1 use AVERAGE(EOP, BOP) where BOP = prior column (older year).
    # For col 2 (earliest year) use EOP only — no prior year available.
    def avg_bs(row_ref, c):
        if c > 2:
            return f"AVERAGE('Balance Sheet'!{cl(c)}{row_ref},'Balance Sheet'!{cl(c-1)}{row_ref})"
        return f"'Balance Sheet'!{cl(c)}{row_ref}"

    rec_days_row = row
    row = write_formula_row(ws, row, "Receivables Days  (Avg Rec / Rev × 365)",
        lambda r,c: f"=IFERROR({avg_bs(rec,c)}/{pl(rev,c)}*365,\"\")", n, is_days=True)
    inv_days_row = row
    row = write_formula_row(ws, row, "Inventory Days  (Avg Inv / COGS × 365)",
        lambda r,c: f"=IFERROR({avg_bs(inv,c)}/{pl(cogs,c)}*365,\"\")", n, is_days=True)
    ap_days_row = row
    row = write_formula_row(ws, row, "Payables Days  (Avg AP / COGS × 365)",
        lambda r,c: f"=IFERROR({avg_bs(ap,c)}/{pl(cogs,c)}*365,\"\")", n, is_days=True)
    row = write_formula_row(ws, row, "Cash Conversion Cycle  (Rec Days + Inv Days − AP Days)",
        lambda r,c: f"=IFERROR({cl(c)}{rec_days_row}+{cl(c)}{inv_days_row}-{cl(c)}{ap_days_row},\"\")",
        n, is_days=True, bold=True, bg=C_ALT)

    row = blank_row(ws, row, nc)

    # ── CASH FLOW QUALITY ─────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "CASH FLOW QUALITY", nc)
    row = write_formula_row(ws, row, "FCF Margin  (FCF / Revenue)",
        lambda r,c: f"=IFERROR({cf(fcf,c)}/{pl(rev,c)},\"\")", n, is_pct=True)
    row = write_formula_row(ws, row, "FCF Conversion  (FCF / Net Income)",
        lambda r,c: f"=IFERROR({cf(fcf,c)}/{pl(ni,c)},\"\")", n, is_pct=True)
    row = write_formula_row(ws, row, "Capex as % of Revenue",
        lambda r,c: f"=IFERROR(ABS({cf(capex,c)})/{pl(rev,c)},\"\")", n, is_pct=True)
    row = write_formula_row(ws, row, "CFO / Net Income  (Cash Quality)",
        lambda r,c: f"=IFERROR({cf(cfo,c)}/{pl(ni,c)},\"\")", n, is_pct=True)

    row = blank_row(ws, row, nc)

    # ── PER SHARE ─────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "PER SHARE METRICS", nc)
    dilsh = [gm(d,"weightedAverageShsOutDil") for d in is_data]
    shares_row = row
    row = write_data_row(ws, row, "Diluted Shares (mm)", dilsh, years, color=C_BLUE)
    row = write_formula_row(ws, row, "FCF per Share",
        lambda r,c: f"=IFERROR({cf(fcf,c)}/'Ratios & FCF'!{cl(c)}{shares_row},\"\")", n)
    row = write_formula_row(ws, row, "Book Value per Share",
        lambda r,c: f"=IFERROR({bs(te,c)}/'Ratios & FCF'!{cl(c)}{shares_row},\"\")", n)

    row = blank_row(ws, row, nc)

    # ── MODEL CONTROLS ────────────────────────────────────────────────────────
    # Each check shows the difference between two values that should reconcile.
    # A zero (or blank) difference = PASS.  Any non-zero = investigate.
    row = write_section_hdr(ws, row, "MODEL CONTROLS — KEY RECONCILIATION CHECKS", nc, "8B0000")
    row = write_section_hdr(ws, row,
        "Zero = OK  |  Non-zero = investigate  |  Blank = FMP returned no data for that field",
        nc, "555555")

    # ── Master check placeholder — fixed after all individual checks are written ──
    _master_row = row
    row = write_formula_row(ws, row, "MASTER CHECK", lambda r,c: '=""', n_years=n,
        bold=True, bg=C_SUBTOTAL)

    row = blank_row(ws, row, nc)

    check_rows = []   # accumulate row numbers of numeric check cells

    def add_check(label, fml_fn, note=False):
        nonlocal row
        r = row
        row = write_formula_row(ws, row, label, fml_fn, n)
        if not note:
            check_rows.append(r)

    # ── P&L checks ────────────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "P&L Checks", nc, C_DETAIL_HD)

    add_check("Gross Profit: (Revenue - COGS) vs Reported  [= 0]",
        lambda r,c: f"=IFERROR({pl(rev,c)}-{pl(cogs,c)}-{pl(gp,c)},\"\")")

    # EBITDA is now formula = EBIT + DA so this will always be 0 — kept as integrity check
    add_check("EBITDA: (EBIT + D&A) vs Reported  [= 0  — formula-driven, always passes]",
        lambda r,c: f"=IFERROR({pl(ebit,c)}+{pl(da,c)}-{pl(ebitda,c)},\"\")")

    _int_inc = pl_refs["int_inc"]
    _int_exp = pl_refs["int_exp"]
    add_check("Below-the-Line Residual: (EBT - EBIT) - Net Interest  [non-zero = Other Non-Op items, informational]",
        lambda r,c: (
            f"=IFERROR(({pl(ebt,c)}-{pl(ebit,c)})"
            f"-({pl(_int_inc,c)}-{pl(_int_exp,c)}),\"\")"
        ), note=True)   # informational only — excluded from master check

    add_check("Net Income: (EBT - Tax) vs Reported  [= 0]",
        lambda r,c: f"=IFERROR({pl(ebt,c)}-{pl(tax,c)}-{pl(ni,c)},\"\")")

    # ── Balance Sheet checks ──────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "Balance Sheet Checks", nc, C_DETAIL_HD)

    bs_tot_a  = bs_refs["tot_assets"]
    bs_tl     = bs_refs["tl"]
    bs_te_r   = bs_refs["te"]
    bs_tca_r  = bs_refs["tca"]

    add_check("BS Balanced: Total Assets vs (Total Liabilities + Total Equity)  [= 0]",
        lambda r,c: f"=IFERROR({bs(bs_tot_a,c)}-{bs(bs_tl,c)}-{bs(bs_te_r,c)},\"\")")

    add_check("Assets: (Total CA + Total LT Assets) vs Total Assets  [= 0]",
        lambda r,c: (
            f"=IFERROR({bs(bs_tca_r,c)}"
            f"+('Balance Sheet'!{cl(c)}{bs_refs['tlta']})"
            f"-{bs(bs_tot_a,c)},\"\")"
        ))

    # ── Cash Flow checks ──────────────────────────────────────────────────────
    row = write_section_hdr(ws, row, "Cash Flow Checks", nc, C_DETAIL_HD)

    cf_cfo_r  = cf_refs["cfo"]
    cf_cfi_r  = cf_refs["cfi"]
    cf_cff_r  = cf_refs["cff"]
    cf_capex_r = cf_refs["capex"]
    cf_fcf_r   = cf_refs["fcf"]

    add_check("Cash Change: CFO + CFI + CFF vs Reported Net Change  [= 0]",
        lambda r,c: (
            f"=IFERROR({cf(cf_cfo_r,c)}+{cf(cf_cfi_r,c)}+{cf(cf_cff_r,c)}"
            f"-'Cash Flow'!{cl(c)}{cf_refs['net_change']},\"\")"
        ))

    add_check("FCF: (CFO + Capex) vs Reported FCF  [= 0]",
        lambda r,c: f"=IFERROR({cf(cf_cfo_r,c)}+{cf(cf_capex_r,c)}-{cf(cf_fcf_r,c)},\"\")")

    # ── UFCF directional sense check (informational) ──────────────────────────
    row = write_section_hdr(ws, row, "UFCF Bridge Check  [informational — not exact]", nc, C_DETAIL_HD)
    add_check("UFCF vs (CFO + Capex)  [difference = NWC and tax adjustments]",
        lambda r,c: f"=IFERROR({cl(c)}{ufcf_row}-({cf(cf_cfo_r,c)}+{cf(cf_capex_r,c)}),\"\")",
        note=True)

    # ── Master check: counts how many check cells are non-zero ────────────────
    def master_fml(r, c):
        fail_parts = "+".join(
            f"IFERROR((ABS({cl(c)}{cr})>0.01)*1,0)" for cr in check_rows
        )
        return (
            f'=IF(({fail_parts})=0,'
            f'"ALL PASS","FAIL: "&({fail_parts})&" check(s) non-zero")'
        )

    # Write master check label and formula (overwrite placeholder)
    ws.cell(row=_master_row, column=1).value = "MASTER CHECK — ALL CONTROLS"
    ws.cell(row=_master_row, column=1).font  = fnt(bold=True, color=C_WHITE, size=10)
    ws.cell(row=_master_row, column=1).fill  = fll("8B0000")
    ws.cell(row=_master_row, column=1).border = brd()
    ws.cell(row=_master_row, column=1).alignment = Alignment(horizontal="left", indent=1)
    for i in range(n):
        col = i + 2
        cell = ws.cell(row=_master_row, column=col)
        cell.value = master_fml(_master_row, col)
        cell.font  = fnt(bold=True, color=C_WHITE, size=10)
        cell.fill  = fll("8B0000")
        cell.border = brd()
        cell.alignment = Alignment(horizontal="center")

# ═══════════════════════════════════════════════════════════════════════════════
# SEGMENTATION TAB  (unchanged from v3)
# ═══════════════════════════════════════════════════════════════════════════════
def build_segments(wb, ticker, years):
    ws = wb.create_sheet("Segments")
    n  = len(years)
    nc = max(n + 1, 4)
    setup_ws(ws, years)

    row = write_tab_title(ws, 1, f"{ticker} — Revenue Segmentation", nc,
        subtitle="Requires FMP Starter plan or above. Data sourced from company filings.")

    prod_data = fetch_segment("revenue-product-segmentation", ticker)
    geo_data  = fetch_segment("revenue-geographic-segments",  ticker)

    def render_segment(data, title, start_row):
        r = write_section_hdr(ws, start_row, title, nc, C_SUMMARY_HD)
        if not data:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nc)
            msg = ws.cell(row=r, column=1,
                value="Data not available on your current FMP plan (requires Starter or above). "
                      "Upgrade at financialmodelingprep.com to unlock segment data.")
            msg.font = fnt(italic=True, color="AA0000", size=10)
            ws.row_dimensions[r].height = 24
            return r + 2

        recent = data[:n]
        meta = {"date","symbol","reportedCurrency","cik","filingDate",
                "acceptedDate","fiscalYear","period"}

        def extract_val(d, seg):
            v = d.get(seg)
            if v is None:
                return None
            if isinstance(v, dict):
                v = list(v.values())[0] if v else None
            try:
                return round(float(v) / 1e6, 2) if v is not None else None
            except:
                return None

        segments = []
        for d in recent:
            for k in d:
                if k not in meta and k not in segments:
                    segments.append(k)

        seg_years = [d.get("date","")[:4] for d in recent]
        r = write_year_hdr(ws, r, seg_years, nc)

        for seg in segments:
            vals = [extract_val(d, seg) for d in recent]
            r = write_data_row(ws, r, seg, vals, seg_years)
        return r + 1

    row = render_segment(prod_data, "PRODUCT / BUSINESS SEGMENT REVENUE ($mm)", row)
    row = blank_row(ws, row, nc)
    row = render_segment(geo_data, "GEOGRAPHIC SEGMENT REVENUE ($mm)", row)

# ═══════════════════════════════════════════════════════════════════════════════
# COVER TAB  (unchanged from v3)
# ═══════════════════════════════════════════════════════════════════════════════
def build_cover(wb, ticker, years, is_data):
    ws = wb.active
    ws.title = "Cover"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 22
    nc = 2

    row = write_tab_title(ws, 1, f"{ticker.upper()} — Financial Model", nc,
        subtitle=f"Source: Financial Modeling Prep API  |  All figures USD millions  |  FY {years[0]}–{years[-1]}")

    row += 1
    d = is_data[-1]

    row = write_section_hdr(ws, row, "KEY METRICS — MOST RECENT YEAR", nc, C_SUMMARY_HD)

    def cov_row(ws, r, label, val, fmt):
        c1 = ws.cell(row=r, column=1, value=label)
        c1.font = fnt(size=10); c1.border = brd()
        c1.alignment = Alignment(horizontal="left", indent=1)
        c2 = ws.cell(row=r, column=2, value=val)
        c2.font = fnt(color=C_BLUE, size=10)
        c2.alignment = Alignment(horizontal="right")
        c2.border = brd()
        c2.number_format = fmt
        return r + 1

    row = cov_row(ws, row, f"Fiscal Year",          years[-1],                  "@")
    row = cov_row(ws, row, "Revenue ($mm)",          gm(d,"revenue"),           "#,##0.0")
    row = cov_row(ws, row, "Gross Profit ($mm)",     gm(d,"grossProfit"),       "#,##0.0")
    row = cov_row(ws, row, "EBITDA ($mm)",           gm(d,"ebitda"),            "#,##0.0")
    row = cov_row(ws, row, "EBIT ($mm)",             gm(d,"operatingIncome"),   "#,##0.0")
    row = cov_row(ws, row, "Net Income ($mm)",       gm(d,"netIncome"),         "#,##0.0")
    row = cov_row(ws, row, "Free Cash Flow ($mm)",   gm(d,"freeCashFlow") if d.get("freeCashFlow") else None, "#,##0.0")
    row = cov_row(ws, row, "EPS Diluted",            g(d,"epsdiluted"),         "$#,##0.00")
    row = cov_row(ws, row, "Gross Margin %",         g(d,"grossProfitRatio"),   "0.0%")
    row = cov_row(ws, row, "EBITDA Margin %",        g(d,"ebitdaratio"),        "0.0%")
    row = cov_row(ws, row, "Net Margin %",           g(d,"netIncomeRatio"),     "0.0%")

    row += 1
    row = write_section_hdr(ws, row, "WORKBOOK STRUCTURE", nc, C_DETAIL_HD)
    tabs = [
        ("Cover",         "This page — key metrics snapshot"),
        ("P&L",           "Income statement: summary + all FMP line items"),
        ("Balance Sheet", "Balance sheet: summary + all FMP line items"),
        ("Cash Flow",     "Cash flow: summary + all FMP line items"),
        ("Ratios & FCF",  "UFCF bridge, LFCF bridge, and full ratio suite"),
        ("Segments",      "Product & geographic revenue segments (plan-dependent)"),
    ]
    for tab, desc in tabs:
        c1 = ws.cell(row=row, column=1, value=tab)
        c1.font = fnt(bold=True, size=10); c1.border = brd()
        c1.alignment = Alignment(horizontal="left", indent=1)
        c2 = ws.cell(row=row, column=2, value=desc)
        c2.font = fnt(size=10); c2.border = brd()
        c2.alignment = Alignment(horizontal="left", indent=1)
        row += 1

    row += 1
    note = ws.cell(row=row, column=1,
        value="Colour convention: Blue = raw API data input | Black = formula/calculated | Green = cross-sheet link")
    note.font = fnt(size=9, italic=True, color="666666")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)

# ═══════════════════════════════════════════════════════════════════════════════
# WACC TAB
# ═══════════════════════════════════════════════════════════════════════════════
def build_wacc(wb, ticker, is_data, bs_data, manual_rating=None):
    """Build WACC & Cost of Capital sheet."""

    NC = 3   # columns: label | value | source/note

    ws = wb.create_sheet("WACC")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 42
    ws.freeze_panes = "A3"

    # ── Sheet-local helpers ───────────────────────────────────────────────────
    def wrow(r, label, val, note="", bold=False, bg=C_WHITE,
             val_color=C_BLACK, is_pct=False):
        ca = ws.cell(row=r, column=1, value=label)
        ca.font = fnt(bold=bold, size=10)
        ca.fill = fll(bg); ca.border = brd()
        ca.alignment = Alignment(horizontal="left", indent=1)
        cb = ws.cell(row=r, column=2, value=val)
        cb.font = fnt(bold=bold, color=val_color, size=10)
        cb.fill = fll(bg); cb.border = brd()
        cb.alignment = Alignment(horizontal="right")
        if is_pct:
            cb.number_format = '0.00%;(0.00%);"-"'
        elif isinstance(val, (int, float)):
            cb.number_format = '#,##0.00;(#,##0.00);"-"'
        elif isinstance(val, str) and val.startswith("="):
            cb.number_format = ('#,##0.00;(#,##0.00);"-"'
                                if not is_pct else '0.00%;(0.00%);"-"')
        cc = ws.cell(row=r, column=3, value=note)
        cc.font = fnt(size=9, italic=True, color="777777")
        cc.fill = fll(bg); cc.border = brd()
        cc.alignment = Alignment(horizontal="left", indent=1)
        return r + 1

    def prow(r, label, val, note="", is_pct=False):
        """Percentage formula row — thin wrapper around wrow."""
        return wrow(r, label, val, note, is_pct=True, val_color=C_BLACK)

    def input_row(r, label, val, note="", is_pct=False):
        """Blue override cell pre-filled with suggested value."""
        ca = ws.cell(row=r, column=1, value=label)
        ca.font = fnt(bold=True, size=10)
        ca.fill = fll(C_OVR_BG); ca.border = brd()
        ca.alignment = Alignment(horizontal="left", indent=1)
        cb = ws.cell(row=r, column=2, value=val)
        cb.font = fnt(bold=True, color=C_BLUE, size=10)
        cb.fill = fll(C_OVR_BG); cb.border = brd()
        cb.alignment = Alignment(horizontal="right")
        cb.number_format = ('0.00%;(0.00%);"-"' if is_pct
                            else '0.000;(0.000);"-"')
        cc = ws.cell(row=r, column=3, value=note)
        cc.font = fnt(size=9, italic=True, color="2E7D32")
        cc.fill = fll(C_OVR_BG); cc.border = brd()
        cc.alignment = Alignment(horizontal="left", indent=1)
        return r + 1

    def rat_row(r, text):
        """Merged rationale row (AI output)."""
        ws.merge_cells(start_row=r, start_column=1,
                       end_row=r, end_column=NC)
        c = ws.cell(row=r, column=1, value=text)
        c.font = fnt(size=9, italic=True, color="5D4037")
        c.fill = fll(C_AI_RAT); c.border = brd()
        c.alignment = Alignment(horizontal="left", indent=2,
                                wrap_text=True)
        ws.row_dimensions[r].height = 48
        return r + 1

    def flag_row(r, text):
        ws.merge_cells(start_row=r, start_column=1,
                       end_row=r, end_column=NC)
        c = ws.cell(row=r, column=1, value=text)
        c.font = fnt(size=9, bold=True, color="B71C1C")
        c.fill = fll(C_FLAG_BG); c.border = brd()
        c.alignment = Alignment(horizontal="left", indent=2)
        return r + 1

    def blank(r):
        for col in range(1, NC + 1):
            c = ws.cell(row=r, column=col)
            c.fill = fll(C_WHITE); c.border = brd()
        return r + 1

    def shdr(r, text, color=None):
        return write_section_hdr(ws, r, text, NC, color or C_SECTION)

    # ── Fetch data ────────────────────────────────────────────────────────────
    print("  Fetching WACC inputs...")

    # FMP profile
    prof = {}
    try:
        p = requests.get(
            f"https://financialmodelingprep.com/stable/profile"
            f"?symbol={ticker}&apikey={API_KEY}", timeout=10
        ).json()
        prof = (p[0] if isinstance(p, list) and p
                else p if isinstance(p, dict) else {})
    except Exception:
        pass

    raw_beta = float(prof.get("beta") or 0) or None
    mktcap   = float(prof.get("marketCap") or 0) or None
    sector   = prof.get("industry") or prof.get("sector") or ""
    price    = prof.get("price", "")
    print(f"    Beta={raw_beta}  MktCap={mktcap}  Sector={sector}")

    # Credit rating — manual input takes priority, then FMP, then synthetic fallback
    if manual_rating:
        fmp_rating    = manual_rating
        rating_source = "User input"
        print(f"    Rating={fmp_rating}  (manual)")
    else:
        fmp_rating = None
        rating_source = "FMP /ratings endpoint"
        try:
            rat = requests.get(
                f"https://financialmodelingprep.com/stable/ratings"
                f"?symbol={ticker}&apikey={API_KEY}", timeout=10
            ).json()
            if isinstance(rat, list) and rat:
                fmp_rating = rat[0].get("rating") or rat[0].get("ratingScore")
        except Exception:
            pass
        print(f"    Rating={fmp_rating}")

    # Balance sheet / income statement — most recent year
    is0 = is_data[-1]
    bs0 = bs_data[-1]
    bs1 = bs_data[-2] if len(bs_data) > 1 else bs_data[-1]

    debt0    = ((bs0.get("shortTermDebt") or 0) +
                (bs0.get("longTermDebt")  or 0))
    debt1    = ((bs1.get("shortTermDebt") or 0) +
                (bs1.get("longTermDebt")  or 0))
    avg_debt = (debt0 + debt1) / 2

    ebit     = abs(is0.get("operatingIncome")   or 0)
    int_exp  = abs(is0.get("interestExpense")    or 0)
    int_inc  = abs(is0.get("interestIncome")     or 0)
    tax_exp  = abs(is0.get("incomeTaxExpense")   or 0)
    pretax   = abs(is0.get("incomeBeforeTax")    or 0)
    eff_tax  = tax_exp / pretax if pretax else 0
    icr      = ebit / int_exp if int_exp else 999

    # Capital structure
    E   = (mktcap or 0) / 1e6
    D   = debt0 / 1e6
    V   = E + D
    w_e = E / V if V else 1.0
    w_d = D / V if V else 0.0

    # FRED rates
    rf,     rf_date = fetch_fred("DGS10")
    rd_aaa, _       = fetch_fred("BAMLC0A1CAAAEY")
    rd_aa,  _       = fetch_fred("BAMLC0A2CAAEY")
    rd_a,   _       = fetch_fred("BAMLC0A3CAEY")
    rd_bbb, _       = fetch_fred("BAMLC0A4CBBBEY")
    rd_hy,  _       = fetch_fred("BAMLH0A0HYM2EY")
    rf = rf or 0.043

    RATING_FRED = {
        "AAA":  (rd_aaa, "FRED BAMLC0A1CAAAEY — AAA"),
        "AA+":  (rd_aa,  "FRED BAMLC0A2CAAEY — AA"),
        "AA":   (rd_aa,  "FRED BAMLC0A2CAAEY — AA"),
        "AA-":  (rd_aa,  "FRED BAMLC0A2CAAEY — AA"),
        "A+":   (rd_a,   "FRED BAMLC0A3CAEY — A"),
        "A":    (rd_a,   "FRED BAMLC0A3CAEY — A"),
        "A-":   (rd_a,   "FRED BAMLC0A3CAEY — A"),
        "BBB+": (rd_bbb, "FRED BAMLC0A4CBBBEY — BBB"),
        "BBB":  (rd_bbb, "FRED BAMLC0A4CBBBEY — BBB"),
        "BBB-": (rd_bbb, "FRED BAMLC0A4CBBBEY — BBB"),
    }
    fred_rd, fred_rd_src = RATING_FRED.get(fmp_rating, (None, "No matched rating"))
    if not fred_rd and fmp_rating and fmp_rating[:2] in ("BB", "B-", "B+",
                                                          "B ", "CC"):
        fred_rd, fred_rd_src = rd_hy, "FRED BAMLH0A0HYM2EY — High Yield"

    # Synthetic rating
    synth_rating, synth_spread = get_synthetic_rating(icr)
    rd_synthetic = rf + synth_spread
    rd_acctg     = int_exp / avg_debt if avg_debt else None

    # Peer betas
    peer_list = []
    for key, peers in SECTOR_PEERS.items():
        if key.lower() in sector.lower() or sector.lower() in key.lower():
            peer_list = [p for p in peers if p != ticker]
            break
    peer_betas = []
    for p in peer_list[:5]:
        try:
            pp = requests.get(
                f"https://financialmodelingprep.com/stable/profile"
                f"?symbol={p}&apikey={API_KEY}", timeout=8
            ).json()
            pp = pp[0] if isinstance(pp, list) and pp else pp
            b  = float(pp.get("beta") or 0)
            if b:
                peer_betas.append((p, round(b, 3)))
        except Exception:
            pass
    peer_vals   = sorted([b for _, b in peer_betas])
    peer_median = peer_vals[len(peer_vals) // 2] if peer_vals else None
    print(f"    Peers: {peer_betas}")

    # Damodaran industry beta → re-lever
    dama_unlevered = DAMODARAN_BETAS["Default"]
    for key, val in DAMODARAN_BETAS.items():
        if (key.lower() in sector.lower() or
                sector.lower() in key.lower()):
            dama_unlevered = val
            break
    de_ratio      = D / E if E else 0
    dama_relevered = round(dama_unlevered * (1 + (1 - eff_tax) * de_ratio), 3)
    blume         = round(0.67 * raw_beta + 0.33, 3) if raw_beta else None

    # ── Average-based defaults (no AI) ───────────────────────────────────────
    net_int = int_inc - int_exp

    # Beta: average of all non-None data points
    beta_candidates = [v for v in [raw_beta, blume, dama_relevered, peer_median]
                       if v is not None]
    sel_beta = round(sum(beta_candidates) / len(beta_candidates), 3) \
               if beta_candidates else 1.0

    # ERP: average of Damodaran implied + historical
    sel_erp = round((DAMODARAN_ERP_IMPLIED + DAMODARAN_ERP_HIST_AVG) / 2, 4)

    # Rd: average of all non-None data points
    rd_candidates = [v for v in [fred_rd, rd_synthetic, rd_acctg]
                     if v is not None]
    sel_rd = round(sum(rd_candidates) / len(rd_candidates), 4) \
             if rd_candidates else 0.05

    # ── Write sheet ───────────────────────────────────────────────────────────
    row = 1
    row = write_tab_title(
        ws, row, f"{ticker.upper()} — WACC & COST OF CAPITAL", NC,
        subtitle=("Blue = user override  |  Green = selected input  |  "
                  "Default pre-filled with average of all sources  |  All figures USD millions"))
    row = blank(row)

    # ── Capital structure ─────────────────────────────────────────────────────
    row = shdr(row, "CAPITAL STRUCTURE")
    eq_row = row
    row = wrow(row, "Equity  (market capitalisation, $mm)", E or None,
               f"FMP marketCap  |  price ${price}", val_color=C_BLUE)
    dbt_row = row
    row = wrow(row, "Debt  (gross book value, $mm)", D or None,
               "FMP: shortTermDebt + longTermDebt, most recent yr-end",
               val_color=C_BLUE)
    tot_row = row
    row = wrow(row, "Total capital  (V = E + D)", f"=B{eq_row}+B{dbt_row}",
               "", bold=True, bg=C_SUMMARY_BG)
    ws.cell(row=tot_row, column=2).number_format = '#,##0.0;(#,##0.0);"-"'
    ew_row = row
    row = prow(row, "Equity weight  (E / V)",
               f"=IFERROR(B{eq_row}/B{tot_row},1)",
               "Weight applied to Re in WACC")
    dw_row = row
    row = prow(row, "Debt weight  (D / V)",
               f"=IFERROR(B{dbt_row}/B{tot_row},0)",
               "Weight applied to after-tax Rd in WACC")
    row = blank(row)

    # ── Risk-free rate ────────────────────────────────────────────────────────
    row = shdr(row, "STEP 1 — RISK-FREE RATE")
    row = wrow(row, "10-yr US Treasury yield  (FRED DGS10)",
               rf, f"FRED series DGS10  |  latest: {rf_date}",
               val_color=C_BLACK, is_pct=True)
    rf_row = row
    row = input_row(row, "► Selected Rf  (override if needed)",
                    round(rf, 4),
                    "Pre-filled with live FRED DGS10 rate", is_pct=True)
    row = blank(row)

    # ── Beta ──────────────────────────────────────────────────────────────────
    row = shdr(row, "STEP 2 — BETA")
    row = wrow(row, "Raw beta  (FMP — 5yr monthly vs S&P 500)",
               raw_beta, "FMP company profile endpoint")
    row = wrow(row, "Blume-adjusted  (0.67 × raw + 0.33)",
               blume, "Mean-reversion toward market beta of 1.0")
    row = wrow(row, f"Damodaran industry unlevered  ({sector or 'sector'})",
               dama_unlevered,
               "Damodaran.com — betas by industry, Jan 2025 US")
    row = wrow(row, f"Damodaran re-levered  "
               f"(D/E = {de_ratio:.2f}x,  t = {eff_tax*100:.1f}%)",
               dama_relevered,
               "= unlevered × (1 + (1 − t) × D/E)")
    if peer_betas:
        peers_str = "  |  ".join(f"{p}: {b}" for p, b in peer_betas)
        row = wrow(row, "Peer median beta",
                   peer_median, peers_str[:52])
    if (raw_beta and dama_unlevered and
            raw_beta > dama_unlevered * 1.8):
        row = flag_row(row,
            f"FLAG: Raw beta ({raw_beta:.2f}) is "
            f"{raw_beta/dama_unlevered:.1f}x Damodaran industry avg "
            f"({dama_unlevered:.2f}) — historical window may include "
            f"structural break or regime change.")
    beta_row = row
    row = input_row(row, "► Selected beta  (user override)",
                    round(sel_beta, 3),
                    f"Default = average of {len(beta_candidates)} source(s) above — override freely")
    row = blank(row)

    # ── Equity risk premium ───────────────────────────────────────────────────
    row = shdr(row, "STEP 3 — EQUITY RISK PREMIUM  (ERP)")
    row = wrow(row, "Damodaran implied ERP — US market (current)",
               DAMODARAN_ERP_IMPLIED,
               "Damodaran.com implied ERP — Jan 2026", is_pct=True)
    row = wrow(row, "Damodaran historical avg — US 1928–2025",
               DAMODARAN_ERP_HIST_AVG,
               "Arithmetic average excess return over T-bill", is_pct=True)
    erp_row = row
    row = input_row(row, "► Selected ERP  (user override)",
                    round(sel_erp, 4),
                    "Default = average of Damodaran implied & historical avg — override freely",
                    is_pct=True)
    row = blank(row)

    # ── Cost of equity ────────────────────────────────────────────────────────
    row = shdr(row, "STEP 4 — COST OF EQUITY  (CAPM:  Re = Rf + β × ERP)")
    re_row = row
    re_cell = ws.cell(row=re_row, column=2,
                      value=f"=B{rf_row}+B{beta_row}*B{erp_row}")
    re_label = ws.cell(row=re_row, column=1,
                       value="Cost of equity  (Re)")
    re_note  = ws.cell(row=re_row, column=3,
                       value="CAPM formula — references selected inputs above")
    for c_ in (re_label, re_cell, re_note):
        c_.fill = fll(C_SUMMARY_BG); c_.border = brd()
    re_label.font  = fnt(bold=True, size=10)
    re_label.alignment = Alignment(horizontal="left", indent=1)
    re_cell.font   = fnt(bold=True, color=C_BLACK, size=10)
    re_cell.alignment = Alignment(horizontal="right")
    re_cell.number_format = '0.00%;(0.00%);"-"'
    re_note.font   = fnt(size=9, italic=True, color="777777")
    re_note.alignment = Alignment(horizontal="left", indent=1)
    row += 1
    row = blank(row)

    # ── Cost of debt ──────────────────────────────────────────────────────────
    row = shdr(row, "STEP 5 — COST OF DEBT  (Pre-tax Rd)")
    row = wrow(row, "Credit rating  (S&P / Moody's)",
               fmp_rating or "Not available",
               rating_source)
    if fred_rd:
        row = wrow(row, "FRED matched yield  (rating-tier index)",
                   fred_rd, fred_rd_src, is_pct=True)
    else:
        row = wrow(row, "FRED matched yield",
                   None, "No rating returned — FRED method not applicable")
    row = wrow(row,
               f"Interest Coverage Ratio  (EBIT / Gross Int Exp)",
               round(icr, 1) if icr < 900 else None,
               f"EBIT ${ebit/1e6:.0f}mm  /  Int Exp ${int_exp/1e6:.0f}mm")
    row = wrow(row, "Synthetic rating  (Damodaran ICR table)",
               synth_rating,
               f"ICR {icr:.1f}x  →  {synth_rating}")
    row = wrow(row, "Synthetic Rd  (Rf + Damodaran default spread)",
               rd_synthetic,
               f"Rf {rf*100:.2f}% + spread {synth_spread*100:.2f}%",
               is_pct=True)
    row = wrow(row, "Accounting Rd  (Gross int exp / Avg gross debt)",
               rd_acctg, "Cross-check only — backward-looking", is_pct=True)
    if net_int > 0:
        row = wrow(row,
                   f"  Note: net interest INCOME of ${net_int/1e6:.0f}mm  "
                   f"(int inc ${int_inc/1e6:.0f}mm > int exp ${int_exp/1e6:.0f}mm)",
                   None,
                   "Gross Rd is the correct basis here — net figure is distorted")
    if D < 1:
        row = flag_row(row,
                       "FLAG: Zero / negligible debt detected — "
                       "Rd is immaterial.  WACC ≈ Re.")
    rd_row = row
    row = input_row(row, "► Selected pre-tax Rd  (user override)",
                    round(sel_rd, 4),
                    f"Default = average of {len(rd_candidates)} source(s) above — override freely",
                    is_pct=True)
    row = blank(row)

    # ── Tax rate ──────────────────────────────────────────────────────────────
    row = shdr(row, "STEP 6 — TAX RATE")
    row = wrow(row, "Effective tax rate  (tax expense / pre-tax income)",
               round(eff_tax, 4),
               f"FMP: ${tax_exp/1e6:.0f}mm  /  ${pretax/1e6:.0f}mm",
               is_pct=True)
    tax_row = row
    row = input_row(row, "► Selected tax rate  (user override)",
                    round(eff_tax, 4),
                    "Adjust to normalised / marginal rate if preferred",
                    is_pct=True)
    row = blank(row)

    # ── WACC output ───────────────────────────────────────────────────────────
    row = shdr(row, "WACC OUTPUT", C_SUMMARY_HD)
    wacc_row = row
    wacc_formula = (f"=B{ew_row}*B{re_row}"
                    f"+B{dw_row}*B{rd_row}*(1-B{tax_row})")
    wl = ws.cell(row=wacc_row, column=1,
                 value="WACC  =  (E/V × Re)  +  (D/V × Rd × (1 − t))")
    wv = ws.cell(row=wacc_row, column=2, value=wacc_formula)
    wn = ws.cell(row=wacc_row, column=3,
                 value="Used as discount rate in DCF tab")
    for c_ in (wl, wv, wn):
        c_.fill = fll(C_SUMMARY_BG); c_.border = brd()
    wl.font = fnt(bold=True, size=11)
    wl.alignment = Alignment(horizontal="left", indent=1)
    wv.font = fnt(bold=True, color=C_BLACK, size=11)
    wv.alignment = Alignment(horizontal="right")
    wv.number_format = '0.00%;(0.00%);"-"'
    wn.font = fnt(size=10, italic=True, color="555555")
    wn.alignment = Alignment(horizontal="left", indent=1)
    ws.row_dimensions[wacc_row].height = 22
    row += 1
    row = blank(row)

    # ── Sensitivity table ─────────────────────────────────────────────────────
    row = shdr(row, "SENSITIVITY — WACC (%)  by Beta offset × ERP offset",
               C_DETAIL_HD)
    beta_deltas = [-1.0, -0.5, 0.0, +0.5, +1.0]
    erp_deltas  = [-0.01, -0.005, 0.0, +0.005, +0.01]
    erp_labels  = ["-1.0%", "-0.5%", "Base ERP", "+0.5%", "+1.0%"]
    beta_labels = ["-1.0",  "-0.5",  "Base β",   "+0.5",  "+1.0"]

    # Header row
    hdr_cell = ws.cell(row=row, column=1,
                       value="Beta offset  \\  ERP offset →")
    hdr_cell.font = fnt(bold=True, size=9)
    hdr_cell.fill = fll(C_SUBTOTAL); hdr_cell.border = brd()
    hdr_cell.alignment = Alignment(horizontal="center")
    for ci, lbl in enumerate(erp_labels):
        c_ = ws.cell(row=row, column=ci + 2, value=lbl)
        c_.font = fnt(bold=True, size=9)
        c_.fill = fll(C_SUBTOTAL); c_.border = brd()
        c_.alignment = Alignment(horizontal="center")
    row += 1

    for bd, bl in zip(beta_deltas, beta_labels):
        is_base_row = (bd == 0.0)
        bg_ = C_SUMMARY_BG if is_base_row else C_WHITE
        lc = ws.cell(row=row, column=1, value=bl)
        lc.font = fnt(bold=is_base_row, size=9)
        lc.fill = fll(bg_); lc.border = brd()
        lc.alignment = Alignment(horizontal="center")
        for ci, ed in enumerate(erp_deltas):
            is_base_cell = is_base_row and ed == 0.0
            f = (f"=B{ew_row}*(B{rf_row}+(B{beta_row}+({bd}))"
                 f"*(B{erp_row}+({ed})))"
                 f"+B{dw_row}*B{rd_row}*(1-B{tax_row})")
            vc = ws.cell(row=row, column=ci + 2, value=f)
            vc.font = fnt(bold=is_base_cell, size=9)
            vc.fill = fll(C_SUMMARY_BG if is_base_cell else bg_)
            vc.border = brd()
            vc.alignment = Alignment(horizontal="right")
            vc.number_format = '0.00%;(0.00%);"-"'
        row += 1

    return {
        "wacc_row": wacc_row, "re_row":   re_row,
        "rf_row":   rf_row,   "beta_row": beta_row,
        "erp_row":  erp_row,  "rd_row":   rd_row,
        "tax_row":  tax_row,
    }

# ═══════════════════════════════════════════════════════════════════════════════
# DCF TAB
# ═══════════════════════════════════════════════════════════════════════════════
def build_dcf(wb, ticker, is_data, bs_data, cf_data, years, pl_refs, bs_refs, wacc_refs, current_price=None):
    """Build DCF sheet — consensus years auto-populated from FMP, remainder user input."""

    last_hist_year = years[-1]
    estimates      = fetch_analyst_estimates(ticker, last_hist_year)

    # Projection year list: all FMP forward years, extended to at least YEARS_PROJ
    est_years = [e["year"] for e in estimates]
    last_yr   = int(last_hist_year)
    n_proj    = max(YEARS_PROJ, len(est_years))
    proj_years = []
    for i in range(1, n_proj + 1):
        proj_years.append(str(last_yr + i))

    # Lookup dict: year → estimate record
    est_map = {e["year"]: e for e in estimates}
    n_hist  = len(years)
    n_term  = 1

    # Column layout: A=labels | hist cols | proj cols | terminal | notes
    NC          = 1 + n_hist + n_proj + n_term + 1
    HIST_COLS   = list(range(2, 2 + n_hist))
    PROJ_COLS   = list(range(2 + n_hist, 2 + n_hist + n_proj))
    TERM_COL    = 2 + n_hist + n_proj
    NOTE_COL    = TERM_COL + 1

    ws = wb.create_sheet("DCF")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    for c in range(2, TERM_COL + 1):
        ws.column_dimensions[cl(c)].width = 13
    ws.column_dimensions[cl(NOTE_COL)].width = 40
    ws.freeze_panes = f"B5"

    # ── Local helpers ─────────────────────────────────────────────────────────
    def wcell(r, c, val, bold=False, bg=C_WHITE, color=C_BLACK,
              italic=False, halign="right", fmt=None, indent=0):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font      = fnt(bold=bold, color=color, size=10, italic=italic)
        cell.fill      = fll(bg)
        cell.border    = brd()
        cell.alignment = Alignment(horizontal=halign, vertical="center",
                                   indent=indent, wrap_text=False)
        if fmt:
            cell.number_format = fmt
        return cell

    def note(r, text, bg=C_WHITE):
        c = ws.cell(row=r, column=NOTE_COL, value=text)
        c.font      = fnt(size=9, italic=True, color="555555")
        c.fill      = fll(bg)
        c.border    = brd()
        c.alignment = Alignment(horizontal="left", indent=1, wrap_text=True)
        ws.row_dimensions[r].height = 20
        return c

    def blank(r):
        for c in range(1, NC + 1):
            ws.cell(row=r, column=c).fill   = fll(C_WHITE)
            ws.cell(row=r, column=c).border = brd()
        ws.row_dimensions[r].height = 6
        return r + 1

    def shdr(r, text, bg=None):
        return write_section_hdr(ws, r, text, NC, bg or C_SECT)

    NUM  = '#,##0.0;(#,##0.0);"-"'
    PCT  = '0.0%;(0.0%);"-"'
    PCT2 = '0.00%;(0.00%);"-"'
    DOLS = '$#,##0.00'

    # ── Title & year header ───────────────────────────────────────────────────
    row = 1
    row = write_tab_title(ws, row,
        f"{ticker.upper()} — DCF VALUATION", NC,
        subtitle=("Grey = historical actual  |  Blue (darker) = FMP analyst consensus  |  "
                  "Blue (lighter) = user input  |  Amber = key assumption"))

    wcell(row, 1, "Fiscal Year", bold=True, bg=C_SUB, halign="left", indent=1)
    for i, yr in enumerate(years):
        wcell(row, HIST_COLS[i], yr, bold=True, bg=C_SUB)
    for i, yr in enumerate(proj_years):
        has_est = yr in est_map
        bg_ = C_BG if has_est else C_ALT
        wcell(row, PROJ_COLS[i], f"{yr}E", bold=True, bg=bg_)
    wcell(row, TERM_COL, "Terminal", bold=True, bg=C_ASSM)
    wcell(row, NOTE_COL, "Source / Notes", bold=True, bg=C_SUB,
          halign="left", indent=1)
    ws.row_dimensions[row].height = 16
    row += 1

    # ── Consensus boundary flag ───────────────────────────────────────────────
    if estimates:
        last_est_yr = est_years[-1]
        n_consensus = len([y for y in proj_years if y in est_map])
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=NC)
        flag = ws.cell(row=row, column=1,
            value=(f"FMP analyst consensus: {est_years[0]}E – {est_years[-1]}E  "
                   f"({n_consensus} year{'s' if n_consensus != 1 else ''})   |   "
                   f"User estimates required from: "
                   f"{next((y for y in proj_years if y not in est_map), 'N/A')}E onwards"))
        flag.font      = fnt(size=9, bold=True, color="1A5276")
        flag.fill      = fll(C_BG)
        flag.border    = brd()
        flag.alignment = Alignment(horizontal="left", indent=2)
        ws.row_dimensions[row].height = 16
        row += 1
    row = blank(row)

    # ── SECTION 1: PROJECTION ASSUMPTIONS ────────────────────────────────────
    row = shdr(row,
        "SECTION 1 — PROJECTION ASSUMPTIONS  "
        "(darker blue = FMP consensus-derived  |  lighter blue = user input)", C_HD)

    # Helper: write one assumption row
    def assm_row(r, label, hist_vals, proj_fn, term_val, is_pct=True,
                 note_text="", term_color=C_BLUE):
        wcell(r, 1, label, bold=True, bg=C_ASSM, halign="left", indent=1)
        fmt = PCT if is_pct else NUM
        # Historical (greyed out — actuals, not assumptions)
        for i, c in enumerate(HIST_COLS):
            v = hist_vals[i] if hist_vals and i < len(hist_vals) else None
            cell = wcell(r, c, v, bg=C_HIST, color="999999", fmt=fmt)
            if v is None:
                cell.value = "—"; cell.number_format = "@"
        # Projection
        for i, c in enumerate(PROJ_COLS):
            yr = proj_years[i]
            val, color_, bg_ = proj_fn(i, yr, c)
            wcell(r, c, val, bg=bg_, color=color_, fmt=fmt)
        # Terminal
        wcell(r, TERM_COL, term_val, bg=C_ASSM, color=term_color, fmt=fmt)
        note(r, note_text, bg=C_ASSM)
        return r + 1

    # Extract historical actuals for back-reference
    hist_rev    = [(g(d, "revenue")        or 0) / 1e6 for d in is_data]
    hist_ebitda = [(g(d, "ebitda")         or 0) / 1e6 for d in is_data]
    hist_da     = [(g(d, "depreciationAndAmortization") or 0) / 1e6 for d in is_data]
    hist_capex  = [(abs(g(d, "capitalExpenditure") or 0)) / 1e6 for d in cf_data]
    hist_tax    = [(abs(g(d, "incomeTaxExpense") or 0) /
                    max(abs(g(d, "incomeBeforeTax") or 1), 1))
                   for d in is_data]

    # Prior-year revenue for growth calc (last historical as base)
    def prior_rev(i):
        """Return $mm revenue for the year before projection index i."""
        if i == 0:
            return hist_rev[-1] if hist_rev[-1] else 1
        # Use the consensus/projected revenue for the prior projection year
        prev_yr = proj_years[i - 1]
        if prev_yr in est_map:
            return est_map[prev_yr]["rev_avg"]
        return None   # can't compute; formula will handle

    # Revenue growth %
    def rev_growth_fn(i, yr, c):
        if yr in est_map:
            e = est_map[yr]
            prev = prior_rev(i)
            val  = round(e["rev_avg"] / prev - 1, 4) if prev else None
            return val, "1A5276", C_BG    # darker blue = consensus-derived
        return 0.08, C_BLUE, C_ALT       # lighter blue = user input default

    rev_growth_defaults = [rev_growth_fn(i, yr, c)[0]
                           for i, (yr, c) in enumerate(zip(proj_years, PROJ_COLS))]

    row = assm_row(row, "Revenue Growth %",
        [None,
         round(hist_rev[1]/hist_rev[0]-1, 4) if hist_rev[0] else None,
         round(hist_rev[2]/hist_rev[1]-1, 4) if len(hist_rev) > 2 and hist_rev[1] else None,
         round(hist_rev[3]/hist_rev[2]-1, 4) if len(hist_rev) > 3 and hist_rev[2] else None,
         round(hist_rev[4]/hist_rev[3]-1, 4) if len(hist_rev) > 4 and hist_rev[3] else None,
        ][:n_hist],
        rev_growth_fn, 0.03,
        note_text=("FMP consensus years: implied growth rate back-calculated from analyst "
                   "revenue estimates.  User input years: edit freely.  "
                   "Terminal = long-run growth rate (typically 2-4%)."),
        term_color=C_BLUE)
    rev_growth_row = row - 1

    # EBITDA margin %
    def ebitda_margin_fn(i, yr, c):
        if yr in est_map:
            e = est_map[yr]
            rev = e["rev_avg"] or 1
            val = round(e["ebitda_avg"] / rev, 4) if rev else None
            return val, "1A5276", C_BG
        # Default: step from last consensus or last historical
        last_known = (est_map[est_years[-1]]["ebitda_avg"] /
                      max(est_map[est_years[-1]]["rev_avg"], 1)
                      if estimates else
                      (hist_ebitda[-1] / max(hist_rev[-1], 1) if hist_rev[-1] else 0.50))
        return round(last_known, 4), C_BLUE, C_ALT

    hist_ebitda_margins = [
        round(hist_ebitda[i] / max(hist_rev[i], 1), 4) if hist_rev[i] else None
        for i in range(n_hist)
    ]
    row = assm_row(row, "EBITDA Margin %",
        hist_ebitda_margins, ebitda_margin_fn, None,
        note_text=("FMP consensus years: margin back-calculated from analyst EBITDA / Revenue.  "
                   "Terminal: enter your long-run normalised EBITDA margin."),
        term_color=C_BLUE)
    ebitda_margin_row = row - 1
    # Terminal EBITDA margin — manual blue input
    wcell(ebitda_margin_row, TERM_COL,
          hist_ebitda_margins[-1] if hist_ebitda_margins else 0.50,
          bg=C_ASSM, color=C_BLUE, fmt=PCT)

    # D&A % revenue — always user input (FMP doesn't estimate)
    hist_da_pct = [round(hist_da[i] / max(hist_rev[i], 1), 4)
                   if hist_rev[i] else None for i in range(n_hist)]
    last_da_pct = next((v for v in reversed(hist_da_pct) if v), 0.02)
    def da_pct_fn(i, yr, c):
        return round(last_da_pct, 4), C_BLUE, C_ALT

    row = assm_row(row, "D&A as % of Revenue  (user input — FMP does not estimate)",
        hist_da_pct, da_pct_fn, round(last_da_pct, 4),
        note_text="Historical from 3-statement model.  FMP provides no forward D&A estimate — enter your own.")
    da_pct_row = row - 1

    # CapEx % revenue — user input
    hist_capex_pct = [round(hist_capex[i] / max(hist_rev[i], 1), 4)
                      if hist_rev[i] else None for i in range(n_hist)]
    last_capex_pct = next((v for v in reversed(hist_capex_pct) if v), 0.02)
    def capex_pct_fn(i, yr, c):
        return round(last_capex_pct, 4), C_BLUE, C_ALT

    row = assm_row(row, "CapEx as % of Revenue  (user input)",
        hist_capex_pct, capex_pct_fn, round(last_capex_pct, 4),
        note_text="Historical from cash flow statement.  Adjust for known capex programmes.")
    capex_pct_row = row - 1

    # NWC change % revenue — user input
    def nwc_pct_fn(i, yr, c):
        return 0.01, C_BLUE, C_ALT

    row = assm_row(row, "Change in NWC as % of Revenue  (user input)",
        [0.01] * n_hist, nwc_pct_fn, 0.01,
        note_text="Cross-ref historical NWC change from Ratios & FCF tab.")
    nwc_pct_row = row - 1

    # Tax rate — show effective; user overrides for normalised
    last_tax = round(hist_tax[-1], 4) if hist_tax else 0.15
    def tax_fn(i, yr, c):
        return round(last_tax, 4), C_BLUE, C_ALT

    row = assm_row(row, "Effective Tax Rate  (user input — adjust to normalised if needed)",
        [round(t, 4) if t else None for t in hist_tax],
        tax_fn, round(last_tax, 4),
        note_text="Historical effective rate from income statement.  Consider using statutory rate for terminal year.")
    tax_row_dcf = row - 1

    row = blank(row)

    # ── SECTION 2: REVENUE & EBITDA ───────────────────────────────────────────
    row = shdr(row, "SECTION 2 — REVENUE & EBITDA  ($mm)", C_SECT)

    # Revenue
    wcell(row, 1, "Revenue", bold=True, bg=C_BG, halign="left", indent=1)
    for i, c in enumerate(HIST_COLS):
        wcell(row, c, round(hist_rev[i], 1) if hist_rev[i] else None,
              bg=C_HIST, color=C_BLUE, fmt=NUM)
    for i, c in enumerate(PROJ_COLS):
        yr = proj_years[i]
        if yr in est_map:
            e = est_map[yr]
            wcell(row, c, round(e["rev_avg"], 1), bg=C_BG, color="1A5276", fmt=NUM)
        else:
            prior_c = HIST_COLS[-1] if i == 0 else PROJ_COLS[i - 1]
            wcell(row, c, f"={cl(prior_c)}{row}*(1+{cl(c)}{rev_growth_row})",
                  bg=C_ALT, fmt=NUM)
    prior_proj = PROJ_COLS[-1]
    wcell(row, TERM_COL,
          f"={cl(prior_proj)}{row}*(1+{cl(TERM_COL)}{rev_growth_row})",
          bg=C_ASSM, fmt=NUM)
    note(row, ("FMP consensus years: analyst revenue average ($mm).  "
               "Range: Low–High shown below.  User years: prior × (1+growth)."))
    rev_row = row; row += 1

    # Revenue low / high (consensus years only)
    wcell(row, 1, "  Analyst Range:  Low — High  ($mm)",
          italic=True, halign="left", indent=2)
    for i, c in enumerate(PROJ_COLS):
        yr = proj_years[i]
        if yr in est_map:
            e = est_map[yr]
            wcell(row, c,
                  f"{e['rev_low']:,.0f} — {e['rev_high']:,.0f}",
                  italic=True, color="555555", bg=C_BG)
        else:
            wcell(row, c, "—", italic=True, color="999999")
    wcell(row, TERM_COL, "—", italic=True, color="999999", bg=C_ASSM)
    note(row, "Sell-side analyst low / high revenue estimates for consensus years")
    row += 1

    # Analyst count
    wcell(row, 1, "  Number of Analysts", italic=True, halign="left", indent=2)
    for i, c in enumerate(PROJ_COLS):
        yr = proj_years[i]
        if yr in est_map:
            n_ = est_map[yr]["n_rev"]
            color_ = ("B71C1C" if n_ < 5 else
                      "E65100" if n_ < 10 else C_BLACK)
            wcell(row, c, n_, italic=True, color=color_,
                  fmt='#,##0', bg=C_BG)
        else:
            wcell(row, c, "—", italic=True, color="999999")
    wcell(row, TERM_COL, "—", italic=True, color="999999", bg=C_ASSM)
    note(row, "Red < 5 analysts — treat estimate with caution.  Orange < 10.")
    row += 1

    # Revenue growth % (display row)
    wcell(row, 1, "  YoY Revenue Growth %", italic=True, halign="left", indent=2)
    for i, c in enumerate(HIST_COLS):
        if i > 0:
            f = f"=IFERROR({cl(c)}{rev_row}/{cl(HIST_COLS[i-1])}{rev_row}-1,\"\")"
            cell = wcell(row, c, f, italic=True, bg=C_HIST, fmt=PCT)
            cell.font = fnt(italic=True, color=C_BLACK)
        else:
            wcell(row, c, "—", italic=True, color="999999", bg=C_HIST)
    for i, c in enumerate(PROJ_COLS):
        prior_c = HIST_COLS[-1] if i == 0 else PROJ_COLS[i - 1]
        cell = wcell(row, c,
                     f"=IFERROR({cl(c)}{rev_row}/{cl(prior_c)}{rev_row}-1,\"\")",
                     italic=True, bg=C_BG if proj_years[i] in est_map else C_ALT, fmt=PCT)
        cell.font = fnt(italic=True, color=C_BLACK)
    wcell(row, TERM_COL,
          f"={cl(TERM_COL)}{rev_growth_row}",
          italic=True, bg=C_ASSM, fmt=PCT)
    note(row, "= this year / prior year − 1  (formula cross-check on assumptions)")
    row += 1
    row = blank(row)

    # EBITDA
    wcell(row, 1, "EBITDA", bold=True, bg=C_BG, halign="left", indent=1)
    for i, c in enumerate(HIST_COLS):
        wcell(row, c, round(hist_ebitda[i], 1) if hist_ebitda[i] else None,
              bg=C_HIST, color=C_BLUE, fmt=NUM)
    for i, c in enumerate(PROJ_COLS):
        yr = proj_years[i]
        if yr in est_map:
            e = est_map[yr]
            wcell(row, c, round(e["ebitda_avg"], 1), bg=C_BG, color="1A5276", fmt=NUM)
        else:
            wcell(row, c, f"={cl(c)}{rev_row}*{cl(c)}{ebitda_margin_row}",
                  bg=C_ALT, fmt=NUM)
    wcell(row, TERM_COL,
          f"={cl(TERM_COL)}{rev_row}*{cl(TERM_COL)}{ebitda_margin_row}",
          bg=C_ASSM, fmt=NUM)
    note(row, ("FMP consensus years: analyst EBITDA average.  "
               "User years: Revenue × EBITDA margin assumption."))
    ebitda_row = row; row += 1

    # EBITDA range
    wcell(row, 1, "  Analyst Range:  Low — High  ($mm)",
          italic=True, halign="left", indent=2)
    for i, c in enumerate(PROJ_COLS):
        yr = proj_years[i]
        if yr in est_map:
            e = est_map[yr]
            wcell(row, c,
                  f"{e['ebitda_low']:,.0f} — {e['ebitda_high']:,.0f}",
                  italic=True, color="555555", bg=C_BG)
        else:
            wcell(row, c, "—", italic=True, color="999999")
    wcell(row, TERM_COL, "—", italic=True, color="999999", bg=C_ASSM)
    note(row, "Sell-side analyst low / high EBITDA estimates for consensus years")
    row += 1

    # EBITDA margin display
    wcell(row, 1, "  EBITDA Margin %", italic=True, halign="left", indent=2)
    for i, c in enumerate(HIST_COLS):
        cell = wcell(row, c,
                     f"=IFERROR({cl(c)}{ebitda_row}/{cl(c)}{rev_row},\"\")",
                     italic=True, bg=C_HIST, fmt=PCT)
        cell.font = fnt(italic=True, color=C_BLACK)
    for i, c in enumerate(PROJ_COLS):
        cell = wcell(row, c,
                     f"=IFERROR({cl(c)}{ebitda_row}/{cl(c)}{rev_row},\"\")",
                     italic=True,
                     bg=C_BG if proj_years[i] in est_map else C_ALT, fmt=PCT)
        cell.font = fnt(italic=True, color=C_BLACK)
    cell = wcell(row, TERM_COL,
                 f"=IFERROR({cl(TERM_COL)}{ebitda_row}/{cl(TERM_COL)}{rev_row},\"\")",
                 italic=True, bg=C_ASSM, fmt=PCT)
    cell.font = fnt(italic=True, color=C_BLACK)
    note(row, "= EBITDA / Revenue  (formula — cross-check on margin assumption)")
    row += 1
    row = blank(row)

    # ── SECTION 3: FCF BUILD ──────────────────────────────────────────────────
    row = shdr(row, "SECTION 3 — UNLEVERED FREE CASH FLOW BUILD  ($mm)", C_SECT)

    # D&A
    wcell(row, 1, "  Less: D&A", halign="left", indent=2)
    for i, c in enumerate(HIST_COLS):
        wcell(row, c, -round(hist_da[i], 1) if hist_da[i] else None,
              bg=C_HIST, color=C_BLUE, fmt=NUM)
    for c in PROJ_COLS + [TERM_COL]:
        bg_ = C_BG if (proj_years[PROJ_COLS.index(c)] in est_map
                       if c in PROJ_COLS else False) else C_ALT
        if c == TERM_COL: bg_ = C_ASSM
        wcell(row, c, f"=-{cl(c)}{rev_row}*{cl(c)}{da_pct_row}",
              bg=bg_, fmt=NUM)
    note(row, "= Revenue × D&A % assumption  (negative = P&L charge)")
    da_row = row; row += 1

    # EBIT
    wcell(row, 1, "EBIT  (Operating Profit)", bold=True, bg=C_SUB, halign="left", indent=1)
    for i, c in enumerate(HIST_COLS):
        ebit_h = (hist_ebitda[i] or 0) - (hist_da[i] or 0)
        wcell(row, c, round(ebit_h, 1), bold=True, bg=C_HIST, fmt=NUM)
        ws.cell(row=row, column=c).font = fnt(bold=True, color=C_BLACK)
    for c in PROJ_COLS + [TERM_COL]:
        wcell(row, c, f"={cl(c)}{ebitda_row}+{cl(c)}{da_row}",
              bold=True, bg=C_SUB, fmt=NUM)
    note(row, "= EBITDA + D&A  (GAAP operating income, EBIT)")
    ebit_row = row; row += 1

    # Tax on EBIT
    wcell(row, 1, "  Less: Tax on EBIT  (unlevered)", halign="left", indent=2)
    for i, c in enumerate(HIST_COLS):
        ebit_h = (hist_ebitda[i] or 0) - (hist_da[i] or 0)
        wcell(row, c, -round(ebit_h * (hist_tax[i] or 0), 1),
              bg=C_HIST, fmt=NUM)
        ws.cell(row=row, column=c).font = fnt(color=C_BLACK)
    for c in PROJ_COLS + [TERM_COL]:
        wcell(row, c, f"=-{cl(c)}{ebit_row}*{cl(c)}{tax_row_dcf}",
              fmt=NUM)
    note(row, "= EBIT × tax rate  (no interest tax shield — UFCF is pre-debt)")
    tax_ebit_row = row; row += 1

    # NOPAT
    wcell(row, 1, "NOPAT", bold=True, bg=C_BG, halign="left", indent=1)
    for i, c in enumerate(HIST_COLS):
        ebit_h = (hist_ebitda[i] or 0) - (hist_da[i] or 0)
        wcell(row, c, round(ebit_h * (1 - (hist_tax[i] or 0)), 1),
              bold=True, bg=C_HIST, fmt=NUM)
        ws.cell(row=row, column=c).font = fnt(bold=True, color=C_BLACK)
    for c in PROJ_COLS + [TERM_COL]:
        wcell(row, c, f"={cl(c)}{ebit_row}+{cl(c)}{tax_ebit_row}",
              bold=True, bg=C_BG, fmt=NUM)
    note(row, "= EBIT × (1 − tax rate)")
    nopat_row = row; row += 1

    # D&A add-back
    wcell(row, 1, "  (+) D&A add-back  (non-cash)", halign="left", indent=2)
    for i, c in enumerate(HIST_COLS):
        wcell(row, c, round(hist_da[i], 1) if hist_da[i] else None,
              bg=C_HIST, fmt=NUM)
        ws.cell(row=row, column=c).font = fnt(color=C_BLACK)
    for c in PROJ_COLS + [TERM_COL]:
        wcell(row, c, f"=-{cl(c)}{da_row}", fmt=NUM)
    note(row, "= D&A added back (non-cash charge — converts NOPAT to cash basis)")
    da_back_row = row; row += 1

    # CapEx
    wcell(row, 1, "  (−) Capital Expenditures", halign="left", indent=2)
    for i, c in enumerate(HIST_COLS):
        wcell(row, c, -round(hist_capex[i], 1) if hist_capex[i] else None,
              bg=C_HIST, color=C_BLUE, fmt=NUM)
    for c in PROJ_COLS + [TERM_COL]:
        wcell(row, c, f"=-{cl(c)}{rev_row}*{cl(c)}{capex_pct_row}", fmt=NUM)
    note(row, "= Revenue × CapEx % assumption  (negative = cash outflow)")
    capex_row = row; row += 1

    # NWC
    wcell(row, 1, "  (−) Increase in Net Working Capital", halign="left", indent=2)
    for i, c in enumerate(HIST_COLS):
        wcell(row, c, None, bg=C_HIST, fmt=NUM)
    for c in PROJ_COLS + [TERM_COL]:
        wcell(row, c, f"=-{cl(c)}{rev_row}*{cl(c)}{nwc_pct_row}", fmt=NUM)
    note(row, "= Revenue × NWC % (positive revenue growth ties up cash in working capital)")
    nwc_row = row; row += 1

    # UFCF
    wcell(row, 1, "UNLEVERED FREE CASH FLOW  (UFCF)", bold=True,
          bg=C_BG, halign="left", indent=1)
    for i, c in enumerate(HIST_COLS):
        h = ((hist_ebitda[i] or 0) - (hist_da[i] or 0)) * (1 - (hist_tax[i] or 0)) + \
            (hist_da[i] or 0) - (hist_capex[i] or 0)
        wcell(row, c, round(h, 1), bold=True, bg=C_HIST, fmt=NUM)
        ws.cell(row=row, column=c).font = fnt(bold=True, color=C_BLACK)
    for c in PROJ_COLS + [TERM_COL]:
        wcell(row, c,
              f"={cl(c)}{nopat_row}+{cl(c)}{da_back_row}+{cl(c)}{capex_row}+{cl(c)}{nwc_row}",
              bold=True, bg=C_BG, fmt=NUM)
    note(row, "= NOPAT + D&A − CapEx − ΔNWC  (pre-debt, pre-interest free cash flow)")
    ufcf_row = row; row += 1

    wcell(row, 1, "  UFCF Margin %", italic=True, halign="left", indent=2)
    for i, c in enumerate(HIST_COLS):
        h = ((hist_ebitda[i] or 0) - (hist_da[i] or 0)) * (1 - (hist_tax[i] or 0)) + \
            (hist_da[i] or 0) - (hist_capex[i] or 0)
        cell = wcell(row, c, round(h / max(hist_rev[i], 1), 4) if hist_rev[i] else None,
                     italic=True, bg=C_HIST, fmt=PCT)
        cell.font = fnt(italic=True, color=C_BLACK)
    for c in PROJ_COLS + [TERM_COL]:
        cell = wcell(row, c,
                     f"=IFERROR({cl(c)}{ufcf_row}/{cl(c)}{rev_row},\"\")",
                     italic=True, fmt=PCT)
        cell.font = fnt(italic=True, color=C_BLACK)
    note(row, "Sense check — should approximate EBITDA margin less capex intensity")
    row += 1
    row = blank(row)

    # ── SECTION 4: TERMINAL VALUE ─────────────────────────────────────────────
    row = shdr(row, "SECTION 4 — TERMINAL VALUE  ($mm)", C_HD)

    # WACC ref
    wacc_ref = (f"=WACC!B{wacc_refs['wacc_row']}"
                if wacc_refs else None)
    wcell(row, 1, "WACC  (from WACC tab)", bold=True, bg=C_ASSM, halign="left", indent=1)
    wv = wcell(row, 2, wacc_ref or 0.12, bold=True, bg=C_ASSM,
               color=C_GREEN if wacc_ref else C_BLUE, fmt=PCT2)
    for c in range(3, NC + 1): wcell(row, c, None, bg=C_ASSM)
    note(row, ("Auto-linked from WACC tab selected output.  "
               "Override manually if needed."), bg=C_ASSM)
    wacc_dcf_row = row; row += 1

    wcell(row, 1, "Terminal Growth Rate  (g)", bold=True, bg=C_ASSM, halign="left", indent=1)
    wcell(row, 2, 0.03, bold=True, bg=C_ASSM, color=C_BLUE, fmt=PCT2)
    for c in range(3, NC + 1): wcell(row, c, None, bg=C_ASSM)
    note(row, "Long-run nominal GDP growth rate — typically 2-4% for US.  Keep below WACC.",
         bg=C_ASSM)
    tg_row = row; row += 1

    wcell(row, 1, "Terminal EV/EBITDA Multiple  (exit multiple method)",
          bold=True, bg=C_ASSM, halign="left", indent=1)
    wcell(row, 2, 20.0, bold=True, bg=C_ASSM, color=C_BLUE, fmt='0.0x')
    for c in range(3, NC + 1): wcell(row, c, None, bg=C_ASSM)
    note(row, "Use current or peer NTM EV/EBITDA.  Cross-check vs Gordon Growth TV.",
         bg=C_ASSM)
    tev_row = row; row += 1
    row = blank(row)

    wcell(row, 1, "Terminal Year UFCF  (grown by g)", halign="left", indent=1)
    wcell(row, 2, f"={cl(TERM_COL)}{ufcf_row}",
          color=C_GREEN, fmt=NUM)
    ws.cell(row=row, column=2).fill = fll(C_WHITE)
    ws.cell(row=row, column=2).border = brd()
    ws.cell(row=row, column=2).alignment = Alignment(horizontal="right")
    for c in range(3, NC + 1): wcell(row, c, None)
    note(row, "Cross-ref from Section 3 — terminal year UFCF (already grown by g in assumptions)")
    tv_ufcf_row = row; row += 1

    wcell(row, 1, "Terminal Value  [Gordon Growth:  UFCF / (WACC − g)]",
          bold=True, bg=C_BG, halign="left", indent=1)
    tv_gg = wcell(row, 2, f"=B{tv_ufcf_row}/(B{wacc_dcf_row}-B{tg_row})",
                  bold=True, bg=C_BG, fmt=NUM)
    tv_gg.font = fnt(bold=True, color=C_BLACK)
    for c in range(3, NC + 1): wcell(row, c, None, bg=C_BG)
    note(row, "Sensitive to g — always cross-check vs Exit Multiple below")
    tv_gg_row = row; row += 1

    wcell(row, 1, "Terminal Value  [Exit Multiple:  Terminal EBITDA × Multiple]",
          bold=True, bg=C_BG, halign="left", indent=1)
    tv_em = wcell(row, 2, f"={cl(TERM_COL)}{ebitda_row}*B{tev_row}",
                  bold=True, bg=C_BG, fmt=NUM)
    tv_em.font = fnt(bold=True, color=C_BLACK)
    for c in range(3, NC + 1): wcell(row, c, None, bg=C_BG)
    note(row, "Anchored to observable market multiples — less model-sensitive than Gordon Growth")
    tv_em_row = row; row += 1
    row = blank(row)

    # ── SECTION 5: DISCOUNTING ────────────────────────────────────────────────
    row = shdr(row, "SECTION 5 — DISCOUNTING  &  ENTERPRISE VALUE  ($mm)", C_SECT)

    wcell(row, 1, "Discount Period  (mid-year convention)", halign="left", indent=1)
    for i, c in enumerate(PROJ_COLS):
        wcell(row, c, i + 0.5, fmt='0.0')
    wcell(row, TERM_COL, len(proj_years), fmt='0.0', bg=C_ASSM)
    note(row, "Mid-year: 0.5, 1.5, 2.5...  assumes FCF received evenly through each year")
    disc_period_row = row; row += 1

    wcell(row, 1, "Discount Factor  =  1 / (1 + WACC) ^ period", halign="left", indent=1)
    for c in PROJ_COLS + [TERM_COL]:
        bg_ = C_ASSM if c == TERM_COL else C_WHITE
        wcell(row, c, f"=1/(1+B{wacc_dcf_row})^{cl(c)}{disc_period_row}",
              bg=bg_, fmt='0.000')
    note(row, "= 1 / (1 + WACC) ^ discount period")
    disc_factor_row = row; row += 1

    wcell(row, 1, "PV of UFCF", bold=True, bg=C_BG, halign="left", indent=1)
    for c in PROJ_COLS:
        wcell(row, c, f"={cl(c)}{ufcf_row}*{cl(c)}{disc_factor_row}",
              bold=True, bg=C_BG, fmt=NUM)
    wcell(row, TERM_COL, None, bg=C_BG)
    note(row, "= UFCF × discount factor")
    pv_ufcf_row = row; row += 1

    sum_f = "+".join(f"{cl(c)}{pv_ufcf_row}" for c in PROJ_COLS)
    wcell(row, 1, "Sum of PV of FCFs  (explicit period)", bold=True,
          bg=C_SUB, halign="left", indent=1)
    wcell(row, 2, f"={sum_f}", bold=True, bg=C_SUB, fmt=NUM)
    for c in range(3, NC + 1): wcell(row, c, None, bg=C_SUB)
    note(row, f"Sum of {len(proj_years)} discounted annual FCFs")
    sum_pv_row = row; row += 1

    for label, tv_r, bg_ in [("PV of Terminal Value  [Gordon Growth]",  tv_gg_row, C_SUB),
                              ("PV of Terminal Value  [Exit Multiple]",  tv_em_row, C_SUB)]:
        wcell(row, 1, label, bold=True, bg=bg_, halign="left", indent=1)
        wcell(row, 2, f"=B{tv_r}*{cl(TERM_COL)}{disc_factor_row}",
              bold=True, bg=bg_, fmt=NUM)
        for c in range(3, NC + 1): wcell(row, c, None, bg=bg_)
        note(row, "Terminal value × terminal-year discount factor")
        if "Gordon" in label: pvtv_gg_row = row
        else:                 pvtv_em_row = row
        row += 1
    row = blank(row)

    # ── SECTION 6: EQUITY BRIDGE ──────────────────────────────────────────────
    row = shdr(row, "SECTION 6 — EQUITY VALUE BRIDGE  &  IMPLIED SHARE PRICE", C_HD)

    ev_rows = {}
    for label, pv_tv in [("Gordon Growth", pvtv_gg_row),
                          ("Exit Multiple",  pvtv_em_row)]:
        wcell(row, 1, f"Enterprise Value  [{label}]",
              bold=True, bg=C_BG, halign="left", indent=1)
        wcell(row, 2, f"=B{sum_pv_row}+B{pv_tv}",
              bold=True, bg=C_BG, fmt=NUM)
        for c in range(3, NC + 1): wcell(row, c, None, bg=C_BG)
        ev_rows[label] = row; row += 1

    # Net debt from balance sheet
    bs0     = bs_data[-1]
    debt    = ((bs0.get("shortTermDebt") or 0) + (bs0.get("longTermDebt") or 0)) / 1e6
    cash    = (bs0.get("cashAndCashEquivalents") or 0) / 1e6
    net_debt = debt - cash

    wcell(row, 1, "  Less: Net Debt  (Debt − Cash)", halign="left", indent=2)
    wcell(row, 2, round(net_debt, 1), color=C_GREEN, fmt=NUM)
    ws.cell(row=row, column=2).fill = fll(C_WHITE); ws.cell(row=row, column=2).border = brd()
    ws.cell(row=row, column=2).alignment = Alignment(horizontal="right")
    for c in range(3, NC + 1): wcell(row, c, None)
    note(row, (f"Auto-linked from Balance Sheet: "
               f"Debt ${debt:,.0f}mm − Cash ${cash:,.0f}mm = ${net_debt:,.0f}mm  "
               f"(negative = net cash)"))
    nd_row = row; row += 1

    wcell(row, 1, "  Less: Minority Interest", halign="left", indent=2)
    mi = (bs0.get("minorityInterest") or 0) / 1e6
    wcell(row, 2, round(mi, 1), color=C_GREEN, fmt=NUM)
    ws.cell(row=row, column=2).fill = fll(C_WHITE); ws.cell(row=row, column=2).border = brd()
    ws.cell(row=row, column=2).alignment = Alignment(horizontal="right")
    for c in range(3, NC + 1): wcell(row, c, None)
    note(row, "Auto-linked from Balance Sheet: minorityInterest")
    mi_row = row; row += 1

    shares = (bs0.get("commonStockSharesOutstanding") or
              is_data[-1].get("weightedAverageShsOutDil") or 0) / 1e6
    wcell(row, 1, "  Shares Outstanding — Diluted  (mm)", halign="left", indent=2)
    wcell(row, 2, round(shares, 1), color=C_GREEN, fmt=NUM)
    ws.cell(row=row, column=2).fill = fll(C_WHITE); ws.cell(row=row, column=2).border = brd()
    ws.cell(row=row, column=2).alignment = Alignment(horizontal="right")
    for c in range(3, NC + 1): wcell(row, c, None)
    note(row, "Auto-linked: weightedAverageShsOutDil from income statement (diluted)")
    sh_row = row; row += 1
    row = blank(row)

    if current_price:
        price = float(current_price)
    else:
        price = float(is_data[-1].get("price") or 0)
        try:
            prof = requests.get(
                f"https://financialmodelingprep.com/stable/profile"
                f"?symbol={ticker}&apikey={API_KEY}", timeout=8
            ).json()
            price = float((prof[0] if isinstance(prof, list) else prof).get("price") or price)
        except Exception:
            pass

    for label in ["Gordon Growth", "Exit Multiple"]:
        ev_r = ev_rows[label]
        wcell(row, 1, f"Implied Share Price  [{label}]  ($)",
              bold=True, bg=C_BG, halign="left", indent=1)
        ip = wcell(row, 2,
                   f"=IFERROR((B{ev_r}-B{nd_row}-B{mi_row})/B{sh_row},\"\")",
                   bold=True, bg=C_BG, fmt=DOLS)
        ip.font = fnt(bold=True, color=C_BLACK)
        for c in range(3, NC + 1): wcell(row, c, None, bg=C_BG)
        if label == "Gordon Growth": ip_gg_row = row
        else:                        ip_em_row = row
        row += 1

    wcell(row, 1, "Current Market Price  ($)", halign="left", indent=1)
    cp = wcell(row, 2, round(price, 2) if price else None, color=C_GREEN, fmt=DOLS)
    cp.fill = fll(C_WHITE); cp.border = brd()
    cp.alignment = Alignment(horizontal="right")
    for c in range(3, NC + 1): wcell(row, c, None)
    note(row, "Auto-linked from FMP company profile — price")
    cp_row = row; row += 1

    for label, ip_r in [("Gordon Growth", ip_gg_row),
                         ("Exit Multiple",  ip_em_row)]:
        wcell(row, 1, f"Upside / (Downside)  [{label}]",
              bold=True, bg=C_SUB, halign="left", indent=1)
        cell = wcell(row, 2,
                     f"=IFERROR(B{ip_r}/B{cp_row}-1,\"\")",
                     bold=True, bg=C_SUB, fmt=PCT)
        cell.font = fnt(bold=True, color=C_BLACK)
        for c in range(3, NC + 1): wcell(row, c, None, bg=C_SUB)
        row += 1

    return {
        "ufcf_row": ufcf_row, "rev_row": rev_row,
        "ebitda_row": ebitda_row, "wacc_dcf_row": wacc_dcf_row,
    }

# ═══════════════════════════════════════════════════════════════════════════════
# SCORECARD
# ═══════════════════════════════════════════════════════════════════════════════
def build_scorecard(wb, ticker, is_data, bs_data, cf_data, years):
    """
    JS Scorecard tab — auto-scores 11 of 13 criteria.
    Quantitative: Revenue CAGR, FCF/NI, Capital Returns, ROIC, D/EBITDA, EBIT/Int
    Proxy-based:  Moat Profile, Management, Execution Risk, P/E vs Median, P/FCF vs Median
    Manual only:  Business Clarity (needs segment data), Long-Term Potential
    Scoring engine follows Master Prompt v2 thresholds.
    """
    ws = wb.create_sheet("Scorecard")
    NC = 8   # columns A–H

    # ── Column widths ─────────────────────────────────────────────────────────
    for col, w in zip("ABCDEFGH", [44, 7, 9, 26, 13, 8, 12, 52]):
        ws.column_dimensions[col].width = w

    # ── Pre-calculate quantitative metrics ───────────────────────────────────

    # 1. Revenue 3yr CAGR
    rev_cagr = None
    if len(is_data) >= 4:
        r_now = is_data[-1].get("revenue") or 0
        r_3ya = is_data[-4].get("revenue") or 0
        if r_now and r_3ya > 0:
            rev_cagr = (r_now / r_3ya) ** (1 / 3) - 1

    # 2. FCF/NI series
    def _fcf(cf):
        v = cf.get("freeCashFlow")
        if v:
            return v
        ocf = cf.get("operatingCashFlow") or 0
        cap = abs(cf.get("capitalExpenditure") or 0)
        return ocf - cap

    fcf_ni_series = []
    for i in range(min(len(is_data), len(cf_data))):
        ni = is_data[i].get("netIncome") or 0
        fcf_ni_series.append(_fcf(cf_data[i]) / ni if ni else None)

    fcf_ni_latest = fcf_ni_series[-1] if fcf_ni_series else None
    fcf_ni_3ya    = fcf_ni_series[-4] if len(fcf_ni_series) >= 4 else None
    fcf_ni_trend  = (fcf_ni_latest is not None and fcf_ni_3ya is not None
                     and (fcf_ni_3ya - fcf_ni_latest) > 0.15)

    # 3. ROIC series
    def _roic(is_, bs_):
        ebit    = abs(is_.get("operatingIncome") or 0)
        tax_e   = abs(is_.get("incomeTaxExpense") or 0)
        pretax  = abs(is_.get("incomeBeforeTax") or 1e-9)
        nopat   = ebit * (1 - min(tax_e / pretax, 0.50))
        equity  = bs_.get("totalStockholdersEquity") or 0
        debt    = (bs_.get("shortTermDebt") or 0) + (bs_.get("longTermDebt") or 0)
        cash    = bs_.get("cashAndCashEquivalents") or 0
        ic      = equity + debt - cash
        return (nopat / ic) if ic > 1 else None

    roic_series = [_roic(is_data[i], bs_data[i])
                   for i in range(min(len(is_data), len(bs_data)))]
    roic_latest = roic_series[-1] if roic_series else None
    roic_3ya    = roic_series[-4] if len(roic_series) >= 4 else None
    roic_trend  = (roic_latest is not None and roic_3ya is not None
                   and (roic_3ya - roic_latest) > 0.05)

    # 4. D/EBITDA and EBIT/Interest
    bs0 = bs_data[-1]; is0 = is_data[-1]; cf0 = cf_data[-1]
    total_debt  = (bs0.get("shortTermDebt") or 0) + (bs0.get("longTermDebt") or 0)
    cash0       = bs0.get("cashAndCashEquivalents") or 0
    net_cash_v  = cash0 - total_debt

    ebitda0 = is0.get("ebitda") or 0
    if not ebitda0:
        da = abs(is0.get("depreciationAndAmortization") or
                 cf0.get("depreciationAndAmortization") or 0)
        ebitda0 = (is0.get("operatingIncome") or 0) + da
    d_ebitda = total_debt / ebitda0 if ebitda0 > 0 else None

    ebit0   = abs(is0.get("operatingIncome") or 0)
    int_exp = abs(is0.get("interestExpense") or 0)
    ebit_int = ebit0 / int_exp if int_exp > 0 else None

    # 5. Capital Returns
    def _ret(cf):
        return (abs(cf.get("commonStockRepurchased") or
                    cf.get("stockRepurchase") or 0) +
                abs(cf.get("dividendsPaid") or 0))

    tot_ret       = _ret(cf0)
    ret_yrs_cnt   = sum(1 for cf_ in cf_data if _ret(cf_) > 0)
    debt_prior    = ((bs_data[-2].get("shortTermDebt") or 0) +
                     (bs_data[-2].get("longTermDebt") or 0)) if len(bs_data) >= 2 else total_debt
    debt_funded   = total_debt > debt_prior * 1.05 and tot_ret > 0

    # 6. Gross / operating margin series (used for moat + management proxies)
    rev_series = [is_.get("revenue") or 0 for is_ in is_data]
    gm_series  = [((is_.get("grossProfit") or 0) / rev if rev else None)
                  for is_, rev in zip(is_data, rev_series)]
    om_series  = [(abs(is_.get("operatingIncome") or 0) / rev if rev else None)
                  for is_, rev in zip(is_data, rev_series)]

    gm_latest    = gm_series[-1]
    gm_3yr_delta = ((gm_series[-1] - gm_series[-4])
                    if len(gm_series) >= 4 and gm_series[-1] and gm_series[-4]
                    else None)
    om_latest    = om_series[-1]
    om_3yr_delta = ((om_series[-1] - om_series[-4])
                    if len(om_series) >= 4 and om_series[-1] and om_series[-4]
                    else None)

    # Revenue growth std dev (for moat consistency + execution risk)
    rev_growths = [(rev_series[i] / rev_series[i - 1] - 1)
                   for i in range(1, len(rev_series))
                   if rev_series[i - 1] and rev_series[i]]
    if len(rev_growths) > 1:
        mu_rg  = sum(rev_growths) / len(rev_growths)
        rev_std = (sum((g - mu_rg) ** 2 for g in rev_growths) / len(rev_growths)) ** 0.5
    else:
        rev_std = 0.0

    # Op margin std dev (for execution risk)
    om_valid = [o for o in om_series if o is not None]
    if len(om_valid) > 1:
        mu_om  = sum(om_valid) / len(om_valid)
        om_std = (sum((o - mu_om) ** 2 for o in om_valid) / len(om_valid)) ** 0.5
    else:
        om_std = 0.0

    # ── FMP ratios: P/E, P/FCF — current + 5yr historical average ───────────
    yf_info          = {}   # kept for beta fallback in moat section
    trailing_pe      = None
    forward_pe       = None
    trailing_pfcf    = None
    pe_5yr_avg       = None
    pfcf_5yr_avg     = None
    sector_pe_med    = None
    sector_pfcf_med  = None

    try:
        rat_url = (f"https://financialmodelingprep.com/stable/ratios"
                   f"?symbol={ticker}&limit=5&apikey={API_KEY}")
        rat_data = requests.get(rat_url, timeout=10).json()
        if isinstance(rat_data, list) and rat_data:
            pe_vals   = [r["priceToEarningsRatio"]    for r in rat_data
                         if r.get("priceToEarningsRatio")    and r["priceToEarningsRatio"]    > 0]
            pfcf_vals = [r["priceToFreeCashFlowRatio"] for r in rat_data
                         if r.get("priceToFreeCashFlowRatio") and r["priceToFreeCashFlowRatio"] > 0]
            trailing_pe   = round(rat_data[0].get("priceToEarningsRatio")    or 0, 1) or None
            trailing_pfcf = round(rat_data[0].get("priceToFreeCashFlowRatio") or 0, 1) or None
            pe_5yr_avg    = round(sum(pe_vals)   / len(pe_vals),   1) if len(pe_vals)   > 1 else None
            pfcf_5yr_avg  = round(sum(pfcf_vals) / len(pfcf_vals), 1) if len(pfcf_vals) > 1 else None
            print(f"  FMP ratios: P/E={trailing_pe}  5yr avg={pe_5yr_avg}  "
                  f"P/FCF={trailing_pfcf}  5yr avg={pfcf_5yr_avg}")
    except Exception as e_rat:
        print(f"  FMP ratios fetch failed: {e_rat}")

    # Sector peer P/E and P/FCF medians (FMP ratios, latest year only)
    try:
        # Get sector from FMP profile (1 call — also used for beta below)
        prof_sc = {}
        try:
            p_sc = requests.get(
                f"https://financialmodelingprep.com/stable/profile"
                f"?symbol={ticker}&apikey={API_KEY}", timeout=8
            ).json()
            prof_sc = (p_sc[0] if isinstance(p_sc, list) and p_sc
                       else p_sc if isinstance(p_sc, dict) else {})
        except Exception:
            pass
        sector_str   = prof_sc.get("industry") or prof_sc.get("sector") or ""
        peer_list_sc = []
        for key, peers in SECTOR_PEERS.items():
            if key.lower() in sector_str.lower() or sector_str.lower() in key.lower():
                peer_list_sc = [p for p in peers if p != ticker]
                break
        peer_pes = []; peer_pfcfs = []
        for peer in peer_list_sc[:4]:
            try:
                pr = requests.get(
                    f"https://financialmodelingprep.com/stable/ratios"
                    f"?symbol={peer}&limit=1&apikey={API_KEY}", timeout=8
                ).json()
                if isinstance(pr, list) and pr:
                    pe = pr[0].get("priceToEarningsRatio")
                    pf = pr[0].get("priceToFreeCashFlowRatio")
                    if pe and 0 < pe < 300: peer_pes.append(pe)
                    if pf and 0 < pf < 300: peer_pfcfs.append(pf)
            except Exception:
                pass
        if peer_pes:
            sector_pe_med   = round(sorted(peer_pes)[len(peer_pes) // 2], 1)
        if peer_pfcfs:
            sector_pfcf_med = round(sorted(peer_pfcfs)[len(peer_pfcfs) // 2], 1)
        print(f"  Sector peer P/E median={sector_pe_med}  P/FCF median={sector_pfcf_med}")
    except Exception as e_peers:
        print(f"  Sector peer fetch failed: {e_peers}")

    # yfinance: beta only (for rough WACC in moat proxy) — graceful fallback
    try:
        import yfinance as yf
        yf_info = yf.Ticker(ticker).info or {}
    except Exception:
        yf_info = {}

    # ── Scoring helpers ───────────────────────────────────────────────────────
    TIER_ORDER = ["LOW", "MOD-LOW", "MOD-HIGH", "HIGH"]
    TIER_SCORE = {"HIGH": 10, "MOD-HIGH": 7, "MOD-LOW": 3, "LOW": 0}

    def down_tier(t):
        i = TIER_ORDER.index(t)
        return TIER_ORDER[max(i - 1, 0)]

    def _t_rev(v):
        if v is None:
            return None, "N/A — insufficient data"
        t = ("HIGH"     if v > 0.12 else
             "MOD-HIGH" if v > 0.08 else
             "MOD-LOW"  if v > 0.05 else "LOW")
        return t, f"{v:.1%}"

    def _t_fcf(v, pen):
        if v is None:
            return None, "N/A — insufficient data"
        v2 = abs(v)
        t  = ("HIGH"     if v2 > 0.85 else
              "MOD-HIGH" if v2 > 0.65 else
              "MOD-LOW"  if v2 > 0.50 else "LOW")
        s  = f"{v:.0%}"
        if pen:
            t = down_tier(t)
            s += "  [trend penalty: declined >15pp vs 3yr ago]"
        return t, s

    def _t_ret(tot, yrs, df):
        if tot == 0:
            return "LOW", "No capital returns in latest year"
        s = f"${tot / 1e6:,.0f}mm latest FY"
        if yrs < 3 or df:
            r = "debt-funded" if df else f"only {yrs}/{len(cf_data)}yr history"
            return "MOD-LOW", f"{s} — {r}"
        if yrs < 5:
            return "MOD-HIGH", f"{s} — {yrs}yr equity-funded"
        return "HIGH", f"{s} — {yrs}yr+ consistent equity-funded"

    def _t_roic(v, pen):
        if v is None:
            return None, "N/A — insufficient data"
        t = ("HIGH"     if v > 0.25 else
             "MOD-HIGH" if v > 0.15 else
             "MOD-LOW"  if v > 0.08 else "LOW")
        s = f"{v:.1%}"
        if pen:
            t = down_tier(t)
            s += "  [trend penalty: declined >5pp vs 3yr ago]"
        return t, s

    def _t_de(de, nc):
        if nc > 0:
            return "HIGH", f"Net cash ${nc / 1e6:,.0f}mm — no net leverage"
        if de is None:
            return None, "N/A"
        t = ("LOW"      if de > 4.0 else
             "MOD-LOW"  if de > 2.5 else
             "MOD-HIGH" if de > 1.0 else "HIGH")
        return t, f"{de:.1f}x"

    def _t_ei(v):
        if v is None:
            return "HIGH", "No interest expense — debt-free"
        t = ("HIGH"     if v > 10.0 else
             "MOD-HIGH" if v > 4.0  else
             "MOD-LOW"  if v > 2.0  else "LOW")
        return t, f"{v:.1f}x"

    tier_rev_cagr,  note_rev_cagr  = _t_rev(rev_cagr)
    tier_fcf_ni,    note_fcf_ni    = _t_fcf(fcf_ni_latest, fcf_ni_trend)
    tier_cap_ret,   note_cap_ret   = _t_ret(tot_ret, ret_yrs_cnt, debt_funded)
    tier_roic,      note_roic      = _t_roic(roic_latest, roic_trend)
    tier_d_ebitda,  note_d_ebitda  = _t_de(d_ebitda, net_cash_v)
    tier_ebit_int,  note_ebit_int  = _t_ei(ebit_int)

    # ── Moat proxy (4 indicators → tier) ─────────────────────────────────────
    # Rough WACC for ROIC spread — use FMP profile beta (already fetched above)
    beta_yf   = float(prof_sc.get("beta") or 1.0) or 1.0
    avg_erp   = (DAMODARAN_ERP_IMPLIED + DAMODARAN_ERP_HIST_AVG) / 2
    rough_re  = 0.043 + beta_yf * avg_erp
    avg_debt  = (total_debt + debt_prior) / 2 if debt_prior else total_debt
    tax_r_sc  = min(abs(is0.get("incomeTaxExpense") or 0) /
                    abs(is0.get("incomeBeforeTax") or 1), 0.50)
    rough_rd  = int_exp / avg_debt if avg_debt > 0 else 0.05
    mktcap_sc = yf_info.get("marketCap") or 0
    E_sc = mktcap_sc / 1e6; D_sc = total_debt / 1e6; V_sc = E_sc + D_sc
    w_e_sc = E_sc / V_sc if V_sc > 0 else 0.8
    w_d_sc = D_sc / V_sc if V_sc > 0 else 0.2
    rough_wacc = w_e_sc * rough_re + w_d_sc * rough_rd * (1 - tax_r_sc)

    moat_ind = []; moat_parts = []
    if gm_latest is not None:
        ok = gm_latest > 0.40
        if ok: moat_ind.append(True)
        moat_parts.append(f"GM {gm_latest:.1%} {'✓' if ok else '✗'} (>40%)")
    if gm_3yr_delta is not None:
        ok = gm_3yr_delta > 0.01
        if ok: moat_ind.append(True)
        moat_parts.append(f"GM trend {gm_3yr_delta:+.1%} {'✓' if ok else '✗'} (>+1pp)")
    if roic_latest is not None:
        spread = roic_latest - rough_wacc
        ok = spread > 0.05
        if ok: moat_ind.append(True)
        moat_parts.append(f"ROIC-WACC {spread:+.1%} {'✓' if ok else '✗'} (>+5pp)")
    ok_std = rev_std < 0.08
    if ok_std: moat_ind.append(True)
    moat_parts.append(f"Rev consistency σ={rev_std:.1%} {'✓' if ok_std else '✗'} (<8%)")
    n_moat = len(moat_ind)
    tier_moat = ("HIGH" if n_moat >= 4 else "MOD-HIGH" if n_moat == 3
                 else "MOD-LOW" if n_moat == 2 else "LOW")
    note_moat = "  |  ".join(moat_parts) + f"  [{n_moat}/4 indicators positive — proxy score]"

    # ── Management proxy (4 indicators → tier) ───────────────────────────────
    mgmt_ind = []; mgmt_parts = []
    if roic_latest is not None and roic_3ya is not None:
        chg = roic_latest - roic_3ya
        ok  = chg >= -0.02
        if ok: mgmt_ind.append(True)
        mgmt_parts.append(f"ROIC trend {chg:+.1%} {'✓' if ok else '✗'} (≥-2pp)")
    elif roic_latest is not None:
        mgmt_parts.append(f"ROIC {roic_latest:.1%} (no trend data)")
    if gm_3yr_delta is not None:
        ok = gm_3yr_delta >= -0.01
        if ok: mgmt_ind.append(True)
        mgmt_parts.append(f"GM maintained {gm_3yr_delta:+.1%} {'✓' if ok else '✗'}")
    if om_3yr_delta is not None:
        ok = om_3yr_delta >= -0.02
        if ok: mgmt_ind.append(True)
        mgmt_parts.append(f"Op margin {om_3yr_delta:+.1%} {'✓' if ok else '✗'} (≥-2pp)")
    ok_ret = tier_cap_ret in ("HIGH", "MOD-HIGH")
    if ok_ret: mgmt_ind.append(True)
    mgmt_parts.append(f"Capital returns {tier_cap_ret or 'N/A'} {'✓' if ok_ret else '✗'}")
    n_mgmt = len(mgmt_ind)
    tier_mgmt = ("HIGH" if n_mgmt >= 4 else "MOD-HIGH" if n_mgmt == 3
                 else "MOD-LOW" if n_mgmt == 2 else "LOW")
    note_mgmt = "  |  ".join(mgmt_parts) + f"  [{n_mgmt}/4 indicators positive — proxy score]"

    # ── Execution Risk proxy (rev + margin volatility → tier) ────────────────
    rev_risk_idx = (3 if rev_std < 0.05 else 2 if rev_std < 0.10
                    else 1 if rev_std < 0.18 else 0)
    om_risk_idx  = (3 if om_std < 0.02 else 2 if om_std < 0.04
                    else 1 if om_std < 0.08 else 0)
    exec_idx  = (rev_risk_idx + om_risk_idx) // 2
    tier_exec = TIER_ORDER[exec_idx]
    note_exec = (f"Rev growth σ={rev_std:.1%}  |  Op margin σ={om_std:.1%}"
                 f"  [proxy — lower σ = lower risk = higher score]")

    # ── Valuation: P/E and P/FCF vs 5yr historical average ───────────────────
    def _t_val(current, hist_avg, sect_med, label, roic_v, cagr_v):
        if not current:
            return None, f"N/A — {label} not available from yfinance"
        premium_ok = (roic_v is not None and roic_v > 0.25 and
                      cagr_v is not None and cagr_v > 0.15)
        benchmark  = hist_avg or sect_med
        parts_v    = [f"Current {current:.1f}x"]
        if hist_avg:   parts_v.append(f"5yr avg {hist_avg:.1f}x")
        if sect_med:   parts_v.append(f"Sector median {sect_med:.1f}x")
        note_v = "  |  ".join(parts_v)
        if not benchmark:
            return None, note_v + "  [no benchmark — review manually]"
        delta = (current - benchmark) / benchmark
        note_v += f"  [{delta:+.0%} vs benchmark"
        if delta > 0.25:
            tier_v = "MOD-LOW" if premium_ok else "LOW"
            if premium_ok:
                note_v += " — premium partly justified (ROIC>25% & fwd CAGR>15%)"
        elif delta > 0.10:
            tier_v = "MOD-LOW"
        elif delta >= -0.10:
            tier_v = "MOD-HIGH"
        else:
            tier_v = "HIGH"
        note_v += "]"
        return tier_v, note_v

    pe_current    = forward_pe or trailing_pe
    tier_pe,   note_pe   = _t_val(pe_current,    pe_5yr_avg,   sector_pe_med,
                                   "P/E",   roic_latest, rev_cagr)
    tier_pfcf, note_pfcf = _t_val(trailing_pfcf, pfcf_5yr_avg, sector_pfcf_med,
                                   "P/FCF", roic_latest, rev_cagr)

    # ── Hard floor gates ──────────────────────────────────────────────────────
    gate1 = d_ebitda is not None and d_ebitda > 4.0 and net_cash_v <= 0
    gate2 = ebit_int is not None and ebit_int < 2.0
    floor_cap = (59 if gate1 and gate2 else
                 64 if gate1 or gate2 else None)

    # ── Criteria table definition ─────────────────────────────────────────────
    # (part, label, weight, auto_tier, note, is_auto)
    # Business Clarity and Long-Term Potential remain manual (needs segment data / narrative)
    CRITERIA = [
        ("P1", "Business Clarity",                  2.5,  None,          "Segment data not on current FMP plan — assign manually after reviewing 10-K", False),
        ("P1", "Moat Profile",                       10.0, tier_moat,     note_moat,      True),
        ("P1", "Long-Term Potential",                10.0, None,          "Structural/TAM outlook — assign manually (genuinely qualitative)",            False),
        ("P1", "Management",                          7.5, tier_mgmt,     note_mgmt,      True),
        ("P2", "Revenue 3yr CAGR",                  10.0, tier_rev_cagr, note_rev_cagr,  True),
        ("P2", "Cash Quality  (FCF / Net Income)",  10.0, tier_fcf_ni,   note_fcf_ni,    True),
        ("P2", "Capital Returns",                    5.0,  tier_cap_ret,  note_cap_ret,   True),
        ("P2", "ROIC",                               7.5,  tier_roic,     note_roic,      True),
        ("P3", "Credit Risk  (D / EBITDA)",          5.0,  tier_d_ebitda, note_d_ebitda,  True),
        ("P3", "Interest Cover  (EBIT / Interest)",  7.5,  tier_ebit_int, note_ebit_int,  True),
        ("P3", "Execution Risk",                     5.0,  tier_exec,     note_exec,      True),
        ("P4", "Valuation vs Median  (P/E)",        10.0, tier_pe,       note_pe,        tier_pe   is not None),
        ("P4", "Valuation vs Median  (P/FCF)",      10.0, tier_pfcf,     note_pfcf,      tier_pfcf is not None),
    ]

    # ── Cell writing helpers ──────────────────────────────────────────────────
    def wcell(r, col, val="", bold=False, color=C_BLACK, bg=C_WHITE,
              halign="left", indent=0, italic=False, fmt=None, wrap=False):
        c = ws.cell(row=r, column=col, value=val)
        c.font      = fnt(bold=bold, color=color, italic=italic)
        c.fill      = fll(bg)
        c.border    = brd()
        c.alignment = Alignment(horizontal=halign, vertical="center",
                                indent=indent, wrap_text=wrap)
        if fmt:
            c.number_format = fmt
        return c

    def merge_row(r, val, bold=False, color=C_WHITE, bg=C_SECTION,
                  halign="left", size=10, indent=1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        c = ws.cell(row=r, column=1, value=val)
        c.font      = Font(name="Arial", bold=bold, color=color, size=size)
        c.fill      = fll(bg)
        c.alignment = Alignment(horizontal=halign, vertical="center",
                                indent=indent, wrap_text=True)
        ws.row_dimensions[r].height = 18
        return r + 1

    def blank_row(r):
        for col in range(1, NC + 1):
            ws.cell(row=r, column=col).fill = fll(C_WHITE)
            ws.cell(row=r, column=col).border = brd()
        return r + 1

    def tier_bg(t):
        return {"HIGH": "C8E6C9", "MOD-HIGH": "BBDEFB",
                "MOD-LOW": "FFE0B2", "LOW": "FFCDD2"}.get(t, C_WHITE)

    def tier_fg(t):
        return {"HIGH": "1B5E20", "MOD-HIGH": "1565C0",
                "MOD-LOW": "E65100", "LOW": "B71C1C"}.get(t, C_BLACK)

    SCORE_FORMULA = (
        '=IF({e}="HIGH",10,'
        'IF({e}="MOD-HIGH",7,'
        'IF({e}="MOD-LOW",3,'
        'IF({e}="LOW",0,""))))'
    )

    # ── Dropdown validation for column E ─────────────────────────────────────
    dv = DataValidation(type="list",
                        formula1='"HIGH,MOD-HIGH,MOD-LOW,LOW"',
                        allow_blank=True,
                        showDropDown=False)
    dv.sqref = "E9:E50"
    ws.add_data_validation(dv)

    # ════════════════════════════════════════════════════════════════════════
    # WRITE SHEET
    # ════════════════════════════════════════════════════════════════════════
    row = 1

    # Title
    ws.row_dimensions[row].height = 24
    row = merge_row(row,
                    f"JS SCORECARD — {ticker}  |  Master Prompt v2  |  "
                    f"{datetime.date.today():%d %b %Y}",
                    bold=True, size=12, bg=C_TITLE)

    # Subtitle / instructions
    ws.row_dimensions[row].height = 28
    row = merge_row(
        row,
        "Blue rows = auto-scored (FMP data + yfinance proxies).  "
        "Yellow rows = manual — select tier from dropdown in column E.  "
        "Only 2 criteria require manual input: Business Clarity + Long-Term Potential.  "
        "Score: HIGH=10 | MOD-HIGH=7 | MOD-LOW=3 | LOW=0",
        bold=False, color=C_BLACK, bg="EAF2FB", size=9
    )

    row = blank_row(row)

    # Gate status
    if gate1 or gate2:
        msgs = []
        if gate1:
            msgs.append(f"LEVERAGE GATE: D/EBITDA {d_ebitda:.1f}x > 4.0x")
        if gate2:
            msgs.append(f"COVERAGE GATE: EBIT/Interest {ebit_int:.1f}x < 2.0x")
        ws.row_dimensions[row].height = 20
        row = merge_row(
            row,
            "⚠  HARD FLOOR GATE(S) TRIGGERED — " + "  |  ".join(msgs) +
            f"  →  Overall score capped at {floor_cap}",
            bold=True, color=C_WHITE, bg="B71C1C", size=10
        )
    else:
        ws.row_dimensions[row].height = 16
        row = merge_row(
            row,
            "✓  No hard floor gates triggered  (D/EBITDA and EBIT/Interest within safe thresholds)",
            bold=False, color="1B5E20", bg="C8E6C9", size=9
        )

    row = blank_row(row)

    # Column headers
    ws.row_dimensions[row].height = 20
    for col, (txt, halign) in enumerate([
        ("CRITERION",        "left"),
        ("PART",             "center"),
        ("WEIGHT %",         "center"),
        ("CALCULATED VALUE", "center"),
        ("TIER  ▼",          "center"),
        ("SCORE",            "center"),
        ("WTD SCORE",        "center"),
        ("NOTES / COMMENTARY (editable)", "left"),
    ], start=1):
        c = ws.cell(row=row, column=col, value=txt)
        c.font      = fnt(bold=True, color=C_WHITE, size=9)
        c.fill      = fll(C_DETAIL_HD)
        c.border    = brd()
        c.alignment = Alignment(horizontal=halign, vertical="center", indent=1)
    row += 1

    hdr_row = row  # first criteria row index
    crit_rows = []  # track for SUM formula

    current_part = None
    for part, label, weight, auto_tier, note, is_auto in CRITERIA:
        # Part separator header
        if part != current_part:
            current_part = part
            part_labels = {
                "P1": "PART 1 — BUSINESS QUALITY  (qualitative)",
                "P2": "PART 2 — FINANCIAL PERFORMANCE  (quantitative)",
                "P3": "PART 3 — RISK PROFILE  (quantitative / qualitative)",
                "P4": "PART 4 — VALUATION  (qualitative / user-supplied market data)",
            }
            ws.row_dimensions[row].height = 16
            row = write_section_hdr(ws, row, part_labels[part], NC, C_SECTION)

        # Row background
        row_bg = C_ASSM if is_auto else C_AI_BG   # blue=auto, yellow=qualitative

        ws.row_dimensions[row].height = 18

        # A: Criterion name
        wcell(row, 1, f"  {label}", bold=is_auto, bg=row_bg, halign="left")

        # B: Part
        wcell(row, 2, part, bold=False, bg=row_bg, halign="center", color="555555")

        # C: Weight
        c_wt = wcell(row, 3, weight / 100, bold=False, bg=row_bg, halign="center")
        c_wt.number_format = "0.0%"

        # D: Calculated value
        if is_auto and note:
            wcell(row, 4, note, bold=False, bg=row_bg, halign="left", italic=True,
                  color="1A3A5C")
        else:
            wcell(row, 4, "— user input required", italic=True,
                  color="999999", bg=row_bg)

        # E: Tier (pre-filled for auto; blank for qualitative)
        e_addr = f"E{row}"
        if auto_tier:
            c_tier = ws.cell(row=row, column=5, value=auto_tier)
            c_tier.font      = Font(name="Arial", bold=True, size=10,
                                    color=tier_fg(auto_tier))
            c_tier.fill      = fll(tier_bg(auto_tier))
        else:
            c_tier = ws.cell(row=row, column=5, value=None)
            c_tier.fill = fll(C_WHITE)
        c_tier.border    = brd()
        c_tier.alignment = Alignment(horizontal="center", vertical="center")

        # F: Score (formula)
        score_f = SCORE_FORMULA.replace("{e}", e_addr)
        c_score = wcell(row, 6, score_f, bold=True, bg=row_bg,
                        halign="center", fmt='0;(0);"-"')
        c_score.font = fnt(bold=True, color=C_BLACK)

        # G: Weighted score (formula)
        wt_col   = get_column_letter(3)
        scr_col  = get_column_letter(6)
        c_wscore = wcell(row, 7,
                         f'=IF({scr_col}{row}="","",{wt_col}{row}*{scr_col}{row})',
                         bold=False, bg=row_bg, halign="center", fmt='0.00;(0.00);"-"')

        # H: Notes
        wcell(row, 8, note if is_auto else "", italic=True, color="555555",
              bg=row_bg, halign="left", wrap=True)

        crit_rows.append(row)
        row += 1

    # ── Total section ─────────────────────────────────────────────────────────
    row = blank_row(row)

    ws.row_dimensions[row].height = 20
    # Total row
    for col in range(1, NC + 1):
        ws.cell(row=row, column=col).fill = fll(C_SUBTOTAL)
        ws.cell(row=row, column=col).border = brd()
    c_tot_lbl = ws.cell(row=row, column=1, value="TOTAL SCORE  (max = 100.0)")
    c_tot_lbl.font      = fnt(bold=True, color=C_BLACK, size=11)
    c_tot_lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Weighted total formula — sum of all G cells in crit_rows
    g_refs = "+".join(f"G{r}" for r in crit_rows)
    c_tot = ws.cell(row=row, column=7, value=f"={g_refs}")
    c_tot.font         = fnt(bold=True, size=11)
    c_tot.number_format = "0.00"
    c_tot.alignment    = Alignment(horizontal="center", vertical="center")
    c_tot.border        = brd()

    # Note in H
    cap_txt = f"Floor cap {floor_cap} applies — see gate warning above." if floor_cap else "No floor cap."
    c_cap   = ws.cell(row=row, column=8, value=cap_txt)
    c_cap.font      = fnt(bold=(floor_cap is not None), color=("B71C1C" if floor_cap else "1B5E20"))
    c_cap.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c_cap.border    = brd()
    tot_row = row; row += 1

    # Floor-adjusted row (only shown when cap applies)
    if floor_cap is not None:
        ws.row_dimensions[row].height = 18
        for col in range(1, NC + 1):
            ws.cell(row=row, column=col).fill = fll(C_FLAG_BG)
            ws.cell(row=row, column=col).border = brd()
        c_fl = ws.cell(row=row, column=1,
                       value=f"FLOOR-ADJUSTED SCORE  (capped at {floor_cap})")
        c_fl.font      = fnt(bold=True, color="B71C1C", size=11)
        c_fl.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c_adj = ws.cell(row=row, column=7,
                        value=f"=MIN(G{tot_row},{floor_cap})")
        c_adj.font         = fnt(bold=True, color="B71C1C", size=11)
        c_adj.number_format = "0.00"
        c_adj.alignment    = Alignment(horizontal="center", vertical="center")
        c_adj.border        = brd()
        adj_row = row; row += 1
    else:
        adj_row = tot_row

    # ── Verdict row ───────────────────────────────────────────────────────────
    row = blank_row(row)

    ws.row_dimensions[row].height = 20
    # Verdict uses an Excel formula referencing the appropriate total cell
    score_ref = f"G{adj_row}"
    verdict_f = (
        f'=IF({score_ref}="","Score incomplete — fill qualitative tiers above",'
        f'IF({score_ref}>=80,"STRONG BUY",'
        f'IF({score_ref}>=65,"BUY",'
        f'IF({score_ref}>=50,"HOLD",'
        f'IF({score_ref}>=35,"REDUCE","SELL")))))'
    )
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NC)
    c_v = ws.cell(row=row, column=1, value=verdict_f)
    c_v.font      = Font(name="Arial", bold=True, size=12, color=C_WHITE)
    c_v.fill      = fll(C_SUMMARY_HD)
    c_v.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 24
    row += 1

    # Scoring legend
    row = blank_row(row)
    ws.row_dimensions[row].height = 14
    row = merge_row(
        row,
        "SCORING GUIDE:  ≥80 STRONG BUY  |  65–79 BUY  |  50–64 HOLD  |  35–49 REDUCE  |  <35 SELL  "
        "  ||  Floor gates: D/EBITDA >4x OR EBIT/Int <2x → cap 64;  both → cap 59",
        bold=False, color="444444", bg="F4F8FB", size=8, indent=2
    )

    # ── Metrics dict for portfolio heatmap ────────────────────────────────────
    # auto_score = raw points out of 100 earned from the 11 auto-scored criteria.
    # Max = 87.5  (Business Clarity 2.5 + Long-Term Potential 10.0 = 12.5 manual)
    _TIER_VAL = {"HIGH": 10, "MOD-HIGH": 7, "MOD-LOW": 3, "LOW": 0}
    _auto_criteria = [
        (tier_moat,     10.0),
        (tier_mgmt,      7.5),
        (tier_rev_cagr, 10.0),
        (tier_fcf_ni,   10.0),
        (tier_cap_ret,   5.0),
        (tier_roic,      7.5),
        (tier_d_ebitda,  5.0),
        (tier_ebit_int,  7.5),
        (tier_exec,      5.0),
        (tier_pe,       10.0),
        (tier_pfcf,     10.0),
    ]
    _scored = [(t, w) for t, w in _auto_criteria if t in _TIER_VAL]
    if _scored:
        _auto_score = round(sum((_TIER_VAL[t] / 10) * w for t, w in _scored), 1)
        if floor_cap is not None:
            _auto_score = min(_auto_score, floor_cap)
    else:
        _auto_score = None

    metrics = {
        "roic":         roic_latest,
        "rev_cagr":     rev_cagr,
        "fcf_ni":       fcf_ni_latest,
        "d_ebitda":     d_ebitda,
        "auto_score":   _auto_score,
        "floor_cap":    floor_cap,
        "pe_current":   pe_current,
        "pe_5yr_avg":   pe_5yr_avg,
        "pfcf_current": trailing_pfcf,
        "pfcf_5yr_avg": pfcf_5yr_avg,
    }

    print("  Scorecard tab built.")
    return ws, metrics


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    ticker = input("Enter ticker symbol (e.g. AAPL, MSFT, NVDA): ").strip().upper()
    manual_rating_raw = input(
        "Enter S&P / Moody's credit rating (optional — press Enter to skip): "
    ).strip()
    # Normalise Moody's to S&P if needed; blank → None
    if manual_rating_raw:
        tok = manual_rating_raw.strip().split()[0].strip(".,;:()")
        manual_rating = MOODY_TO_SP.get(tok) or (tok.upper() if tok.upper() in VALID_SP_RATINGS else None)
        if not manual_rating:
            print(f"  Warning: '{manual_rating_raw}' not recognised — ignoring manual rating.")
    else:
        manual_rating = None
    print(f"\nFetching data for {ticker}...")

    try:
        is_data = fetch("income-statement",       ticker)[:YEARS][::-1]
        bs_data = fetch("balance-sheet-statement",ticker)[:YEARS][::-1]
        cf_data = fetch("cash-flow-statement",    ticker)[:YEARS][::-1]
    except ValueError as e:
        print(f"\nERROR: {e}")
        print("\nCommon fixes:")
        print("  1. Check API_KEY is correct at top of script")
        print("  2. Verify ticker — try AAPL to confirm API working")
        print("  3. Under Armour = UAA not UA")
        input("\nPress Enter to exit...")
        return

    years = [d.get("fiscalYear") or d.get("calendarYear") or d["date"][:4] for d in is_data]
    print(f"  Years: {years}")

    wb = Workbook()
    build_cover(wb, ticker, years, is_data)
    pl_refs = build_pl(wb, is_data, years, ticker)
    bs_refs = build_bs(wb, bs_data, years, ticker)
    cf_refs = build_cf(wb, cf_data, years, ticker)
    build_ratios(wb, is_data, bs_data, cf_data, years, ticker, pl_refs, bs_refs, cf_refs)
    build_segments(wb, ticker, years)
    wacc_refs = build_wacc(wb, ticker, is_data, bs_data, manual_rating)
    dcf_refs  = build_dcf(wb, ticker, is_data, bs_data, cf_data, years, pl_refs, bs_refs, wacc_refs)
    build_scorecard(wb, ticker, is_data, bs_data, cf_data, years)

    base  = f"{ticker}_FinancialModel_{years[-1]}"
    fname = f"{base}.xlsx"
    fpath = os.path.join(SCRIPT_DIR, fname)
    counter = 1
    while os.path.exists(fpath):
        fname = f"{base}_v{counter}.xlsx"
        fpath = os.path.join(SCRIPT_DIR, fname)
        counter += 1
    wb.save(fpath)
    print(f"\n  Saved: {fpath}")
    print("  Tabs: Cover | P&L | Balance Sheet | Cash Flow | Ratios & FCF | Segments | WACC | DCF | Scorecard")
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
