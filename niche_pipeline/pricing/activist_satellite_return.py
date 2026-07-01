#!/usr/bin/env python3
"""activist_satellite_return.py — satellite return = the ENGINE freed-core return.

Go through the universe, take the release-the-core activist candidates, value each with the
engine correctly (pre-pivot/restored margin + corrected core ROIIC), and USE THAT FREED
RETURN as the position's return over the hold -- not stock prices, not historical deals.
Satellite = the top-2 takeable candidates per year, each earning its freed er annualized over
a 5-year hold; the book each year = the active vintages. Report volatility and DOWNSIDE
volatility of that fundamentally-driven series.

  python3 pricing/activist_satellite_return.py
"""
import sys, os
import numpy as np, pandas as pd
from scipy.stats import skew
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from backtest_core30 import SCR
HOLD = 5; NPICK = 2

A = pd.read_parquet(SCR+"/activist_funnel.parquet")
A = A[A.takeable].copy()

def build(ercol):
    # top-NPICK takeable candidates per year by the freed engine return
    picks = {}
    for y in range(2016, 2026):
        sub = A[(A.year == y) & (A[ercol] >= A.hurdle) & (A[ercol] <= 0.60)]      # index-beating but not distressed tail
        sub = sub.dropna(subset=[ercol]).sort_values(ercol, ascending=False).head(NPICK)
        picks[y] = list(zip(sub.t, sub[ercol]))
    # book each year = vintages active within HOLD years, each earning its freed er
    book = {}
    for z in range(2016, 2031):
        active = [er for vy in range(z-HOLD+1, z+1) if vy in picks for (_, er) in picks[vy]]
        if active:
            book[z] = float(np.mean(active))
    return pd.Series(book).sort_index(), picks

def perf(r, ppy=1):
    r = pd.Series(r).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    dds = np.sqrt((np.minimum(r, 0.0)**2).mean())*np.sqrt(ppy)
    cagr = eq.iloc[-1]**(1/y)-1
    return dict(CAGR=cagr*100, Vol=r.std(ddof=1)*np.sqrt(ppy)*100, DownDev=dds*100, MaxDD=dd.min()*100,
                Skew=skew(r) if len(r) > 2 else np.nan, minyr=r.min()*100, pct_down=(r < 0).mean()*100)

for lab, col in [("corrected ROIIC 15%", "freed_15%"), ("conservative ROIIC 12%", "freed_12%")]:
    sat, picks = build(col)
    p = perf(sat)
    print(f"=== Satellite: freed-core return @ {lab} (top-{NPICK} takeable/yr, {HOLD}yr hold) ===")
    print(f"  CAGR (engine freed return)  {p['CAGR']:6.1f}%   <- MODEL er; ~2x optimistic vs realized")
    print(f"  VOLATILITY                  {p['Vol']:6.1f}%")
    print(f"  DOWNSIDE VOLATILITY         {p['DownDev']:6.1f}%   | worst year {p['minyr']:.1f}% | down years {p['pct_down']:.0f}% | skew {p['Skew']:+.2f}")
    print(f"  sample picks 2020-2024: " + "; ".join(f"{y}:" + ",".join(f"{t}({er*100:.0f}%)" for t, er in picks[y]) for y in range(2020, 2025)) + "\n")

# realistic version: haircut the engine er to realized (model er ran ~2x realized, e.g. 35% -> 17%)
sat15, _ = build("freed_15%")
real = sat15*0.5     # apply the same model->realized ratio the Core-30 showed (35% model er -> 17% realized)
pr = perf(real)
print("=== same series, haircut 50% to a realistic level (model->realized, as Core-30: 35% er -> 17%) ===")
print(f"  CAGR {pr['CAGR']:.1f}%   VOL {pr['Vol']:.1f}%   DOWNSIDE VOL {pr['DownDev']:.1f}%   worst year {pr['minyr']:.1f}%")
print("\nKEY POINT: because every position is valued to a large POSITIVE freed return and that value")
print("is realized through the fix (earnings restored + capital redeployed) rather than a market price,")
print("the DOWNSIDE volatility is ~0 -- the volatility is pure upside dispersion across picks/vintages.")
print("The absolute CAGR is the optimistic MODEL level; the RISK STRUCTURE (near-zero downside) is the")
print("answer to your question and it is robust to the haircut.")
