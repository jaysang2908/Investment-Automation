"""
csv_schema.py — single source of truth for outputs.csv column order.

To add a new column: append it to COLUMNS. All existing rows will automatically
get an empty value for it on the next read/write cycle — no manual reset needed.
"""

COLUMNS = [
    "Ticker",
    "Price",
    "MktCap_B",
    "GG_Price",
    "GG_Upside",
    "EM_Price",
    "EM_Upside",
    "PE_Current",
    "PE_5yr",
    "PFCF_Current",
    "PFCF_5yr",
    "ROIC",
    "Rev_CAGR",
    "FCF_NI",
    "D_EBITDA",
    "Auto_Score",
    "Floor_Cap",
    "Manual_Clarity",
    "Manual_LTP",
    "Date",
]

HEADER = ",".join(COLUMNS) + "\n"


def migrate(content: str) -> str:
    """
    Read a CSV string with ANY old schema and return a CSV string
    matching the current COLUMNS, filling missing columns with "".
    Duplicate tickers are deduplicated — latest Date wins.
    """
    import csv
    from io import StringIO

    reader = csv.DictReader(StringIO(content))
    rows = list(reader)
    if not rows:
        return HEADER

    # Deduplicate: keep latest row per ticker (by Date string, lexicographic is fine for ISO dates)
    seen = {}
    for row in rows:
        t = row.get("Ticker", "").strip()
        if not t:
            continue
        existing_date = seen.get(t, {}).get("Date", "")
        if row.get("Date", "") >= existing_date:
            seen[t] = row

    lines = [HEADER.rstrip()]
    for row in seen.values():
        lines.append(",".join(str(row.get(col, "")) for col in COLUMNS))
    return "\n".join(lines) + "\n"
