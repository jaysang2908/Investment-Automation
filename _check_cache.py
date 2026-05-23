import json
tickers = ['NFLX', 'NVDA', 'ADBE', 'COST', 'META', 'JNJ', 'JPM', 'BAC', 'AAPL', 'AMD', 'V']
for t in tickers:
    try:
        with open(f'static/data/{t}_data.json') as f:
            d = json.load(f)
        sm = d.get('scorecard_metrics', {})
        dp = d.get('dcf_prices', {})
        w = d.get('wacc_val')
        print(f"{t:6}  wacc={w and round(w,4)}  pe={sm.get('pe_current')}  auto={sm.get('auto_score')}  gg_price={dp.get('gg_price') and round(dp['gg_price'],1)}  fetched={d.get('fetched')}")
    except Exception as e:
        print(f"{t:6}  ERROR: {e}")
