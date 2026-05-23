import fmp_3statementv6 as mdl

tickers = ["NVDA", "NFLX", "META", "MSFT", "AAPL", "ADBE"]
print(f"  {'Ticker':<6}  {'SBC_B':>7}  {'FCF_B':>7}  {'SBC/FCF':>8}  {'Material?'}")
print(f"  {'-'*55}")
for t in tickers:
    try:
        cf = mdl.fetch("cash-flow-statement", t)[:5][::-1]
        if not cf:
            continue
        last = cf[-1]
        sbc = (last.get("stockBasedCompensation") or 0) / 1e9
        fcf_raw = last.get("freeCashFlow")
        if fcf_raw is None:
            fcf_raw = (last.get("operatingCashFlow") or 0) + (last.get("capitalExpenditure") or 0)
        fcf = fcf_raw / 1e9
        pct = sbc / fcf if fcf else 0
        flag = "YES" if pct > 0.05 else ""
        print(f"  {t:<6}  {sbc:>7.2f}  {fcf:>7.2f}  {pct:>8.1%}  {flag}")
    except Exception as e:
        print(f"  {t:<6}  ERROR: {e}")
