"""Smoke test: verify SBC metrics flow from build_scorecard into metrics dict."""
import fmp_3statementv6 as mdl
from openpyxl import Workbook

tickers = ["META", "NFLX", "NVDA"]

for ticker in tickers:
    print(f"\n{'='*50}")
    print(f"  {ticker}")
    print(f"{'='*50}")
    try:
        is_data = mdl.fetch("income-statement",        ticker)[:mdl.YEARS][::-1]
        bs_data = mdl.fetch("balance-sheet-statement", ticker)[:mdl.YEARS][::-1]
        cf_data = mdl.fetch("cash-flow-statement",     ticker)[:mdl.YEARS][::-1]

        wb = Workbook()
        mdl.build_pl(wb, is_data, [d["date"][:4] for d in is_data], ticker)
        mdl.build_cover(wb, ticker, [d["date"][:4] for d in is_data], is_data)
        bs_refs = mdl.build_bs(wb, bs_data, [d["date"][:4] for d in is_data], ticker)
        cf_refs = mdl.build_cf(wb, cf_data, [d["date"][:4] for d in is_data], ticker)

        _, metrics = mdl.build_scorecard(
            wb, ticker, is_data, bs_data, cf_data,
            [d["date"][:4] for d in is_data],
        )

        sbc_b     = metrics.get("sbc_trailing_b")
        fcf_ex    = metrics.get("fcf_ex_sbc_b")
        sbc_pct   = metrics.get("sbc_pct_fcf")
        pfcf_adj  = metrics.get("pfcf_adj")
        pfcf_rep  = metrics.get("pfcf_current")

        print(f"  SBC trailing:   ${sbc_b:.2f}B" if sbc_b else "  SBC trailing:   N/A")
        print(f"  FCF ex-SBC:     ${fcf_ex:.2f}B" if fcf_ex else "  FCF ex-SBC:     N/A")
        print(f"  SBC/FCF:        {sbc_pct:.1%}" if sbc_pct else "  SBC/FCF:        N/A")
        print(f"  P/FCF reported: {pfcf_rep:.1f}x" if pfcf_rep else "  P/FCF reported: N/A")
        print(f"  P/FCF adj:      {pfcf_adj:.1f}x" if pfcf_adj else "  P/FCF adj:      N/A")
        print(f"  Callout shown:  {'YES' if sbc_pct and sbc_pct > 0.05 else 'NO (below 5% threshold)'}")

    except Exception as e:
        import traceback
        print(f"  ERROR: {e}")
        traceback.print_exc()
