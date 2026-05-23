"""Preview the 3yr adjusted-close chart for NVDA via the new helper."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from report_bridge import _fetch_price_history

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime as _dt

TICKER = "NVDA"
labels, values, stats = _fetch_price_history(TICKER, years=3)

print(f"\n=== {TICKER} 3yr price history ===")
print(f"  N points after sampling: {len(values)}")
print(f"  Full series points: {stats.get('n_points')}")
print(f"  First date: {stats.get('first_date')}  @  ${stats.get('first'):.2f}")
print(f"  Last date:  {stats.get('last_date')}  @  ${stats.get('last'):.2f}")
print(f"  3yr High:   ${stats.get('high'):.2f}  on  {stats.get('high_date')}")
print(f"  3yr Low:    ${stats.get('low'):.2f}  on  {stats.get('low_date')}")
print(f"  3yr return: {stats.get('return_pct')*100:+.1f}%")

# Render PNG
dates = [_dt.fromisoformat(d) for d in labels]
fig, ax = plt.subplots(figsize=(10, 4.5), dpi=130)
ax.plot(dates, values, color="#0f6b4e", linewidth=1.6)
ax.fill_between(dates, values, color="#0f6b4e", alpha=0.08)
ax.set_title(f"{TICKER} — 3 Year History (split/dividend-adjusted)",
             fontsize=12, color="#1a1a1a", pad=10, loc="left")
ax.set_ylabel("Price ($)", fontsize=10, color="#666")
ax.tick_params(labelsize=9, colors="#666")
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=4))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %y"))
ax.grid(True, alpha=0.25, color="#e8e6e0")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.spines["left"].set_color("#ddd")
ax.spines["bottom"].set_color("#ddd")

# High/Low annotations
high_dt = _dt.fromisoformat(stats["high_date"])
low_dt  = _dt.fromisoformat(stats["low_date"])
ax.scatter([high_dt], [stats["high"]], color="#0f6b4e", zorder=5, s=30)
ax.annotate(f'High ${stats["high"]:.0f}\n{stats["high_date"]}',
            xy=(high_dt, stats["high"]), xytext=(5, -25), textcoords="offset points",
            fontsize=8, color="#0f6b4e")
ax.scatter([low_dt], [stats["low"]], color="#b54a3a", zorder=5, s=30)
ax.annotate(f'Low ${stats["low"]:.0f}\n{stats["low_date"]}',
            xy=(low_dt, stats["low"]), xytext=(5, 10), textcoords="offset points",
            fontsize=8, color="#b54a3a")

# Stats footer
footer = (f"3yr return: {stats['return_pct']*100:+.1f}%   "
          f"Range: ${stats['low']:.0f} – ${stats['high']:.0f}   "
          f"From: {stats['first_date']}   To: {stats['last_date']}   "
          f"({len(values)} sampled / {stats['n_points']} raw)")
fig.text(0.5, 0.02, footer, ha="center", fontsize=8.5, color="#666",
         fontfamily="monospace")

plt.tight_layout(rect=[0, 0.04, 1, 1])
out = os.path.join(os.path.dirname(__file__), "_preview_NVDA_3yr.png")
plt.savefig(out, facecolor="white", bbox_inches="tight")
print(f"\n  Saved: {out}")
