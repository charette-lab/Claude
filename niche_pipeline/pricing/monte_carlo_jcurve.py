#!/usr/bin/env python3
"""monte_carlo_jcurve.py — activist bucket with a J-curve entry (down 20%, then recover).

Each new deal (2/yr) drops 20% at entry, then compounds harder over the remaining hold so the
4-year TOTAL still equals its deal outcome (return kept). Steady-state book = ~8 concurrent
deals, always carrying 2 in their -20% entry phase. Monte-Carlo from the real 39-deal
distribution; measure the bucket's vol/downside/drawdown, blend it into the valuation-timed
30-70 Core strategy, and plot a representative path so you can see the trajectory.

  python3 pricing/monte_carlo_jcurve.py
"""
import sys, os
import numpy as np, pandas as pd
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from backtest_core30 import SCR
rng = np.random.default_rng(3)
ENTRY, HOLD = -0.20, 4               # J-curve: -20% at entry, then recover over the hold
WMIN, WMAX = 0.30, 0.70

core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]; core.index = pd.to_datetime(core.index)
core = core[(core.index.year >= 2000) & (core.index.year <= 2026)]
dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet", columns=["Date", "er_total", "artifact"])
dec = dec[~dec.artifact.astype(bool)]; dec["year"] = pd.to_datetime(dec["Date"]).dt.year
sig = dec.assign(er=dec.er_total.clip(-1, 5)).groupby("year")["er"].median()
sig = sig[(sig.index >= 1999) & (sig.index <= 2026)]
w_act = (WMIN + (WMAX-WMIN)*(sig-sig.min())/(sig.max()-sig.min())).shift(1).clip(WMIN, WMAX)

DEALS = [(2015,2025,2.0),(2015,2025,1.2),(2016,2025,6.2),(2016,2025,1.7),(2018,2025,1.2),(2018,2025,1.5),(2018,2025,2.8),
         (2019,2025,5.1),(2020,2025,1.2),(2020,2025,5.4),(2024,2025,1.6),(2017,2023,0.4),(2019,2022,1.6),(2015,2020,1.4),
         (2016,2018,1.5),(2015,2017,1.6),(2015,2016,1.7),(2007,2014,2.9),(2011,2014,3.9),(2007,2014,4.4),(2009,2014,5.5),
         (2012,2014,1.6),(2010,2014,5.1),(2011,2014,1.3),(2010,2014,1.5),(2010,2014,2.0),(2007,2014,28.2),(2007,2014,5.5),
         (2011,2013,0.3),(2006,2007,3.1),(2006,2007,1.4),(2006,2014,8.4),(2006,2011,2.3),(2006,2007,3.0),(2007,2011,2.1),
         (2007,2011,1.4),(2007,2011,5.4),(2003,2006,5.3),(2007,2010,0.6)]
ar = np.array([m**(1/max(b-a,1))-1 for a,b,m in DEALS])
YEARS = list(range(2000, 2027))
def jcurve(a):                        # 4 annual returns: -20% then g,g,g with (0.8)(1+g)^3=(1+a)^4
    T = (1+a)**HOLD; g = (T/(1+ENTRY))**(1/(HOLD-1))-1
    return [ENTRY, g, g, g]

def perf(r, ppy=4):
    r = pd.Series(r).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    dds = np.sqrt((np.minimum(r, 0.0)**2).mean())*np.sqrt(ppy); cagr = eq.iloc[-1]**(1/y)-1
    return (cagr*100, r.std(ddof=1)*np.sqrt(ppy)*100, dds*100, dd.min()*100,
            (r.mean()*ppy-0.03)/(r.std(ddof=1)*np.sqrt(ppy)), (r.mean()*ppy-0.03)/dds if dds>0 else np.nan,
            cagr/abs(dd.min()) if dd.min()<0 else np.nan)

def one_path():
    # Committed-capital (equal) weighting -- each deal is one unit of deployed capital, measured on its
    # deal-outcome basis (this is how final_strategy.py measures the true ~26% bucket). The J-curve only
    # RESHAPES each deal's path (-20% entry, then recover to the same 4-yr total); it can't add return.
    # So we keep the J-curve's SHAPE (entry drawdown / vol) but reset the CAGR level to the true
    # constant-return book, removing the return-smoothing (PE illiquid-reporting) artifact.
    positions = []; aj, at = {}, {}
    for y in YEARS:
        for a in rng.choice(ar, 2):
            positions.append({"j": jcurve(a), "a": a, "age": 0})
        aj[y] = float(np.mean([p["j"][p["age"]] for p in positions])) if positions else 0.0
        at[y] = float(np.mean([p["a"] for p in positions])) if positions else 0.0
        for p in positions:
            p["age"] += 1
        positions = [p for p in positions if p["age"] < HOLD]
    cj = np.prod([1+aj[y] for y in YEARS])**(1/len(YEARS))-1
    ct = np.prod([1+at[y] for y in YEARS])**(1/len(YEARS))-1
    return {y: aj[y]-(cj-ct) for y in YEARS}             # J-curve shape, true committed-capital return level

NSIM = 2000; act_m, blend_m = [], []
sample = None
for i in range(NSIM):
    aa = one_path()
    actq = pd.Series({q: (1+aa[q.year])**0.25-1 for q in core.index})
    wq = pd.Series({q: (float(w_act.get(q.year)) if not pd.isna(w_act.get(q.year, np.nan)) else 0.5) for q in core.index})
    blendq = wq*actq + (1-wq)*core
    act_m.append(perf(actq)); blend_m.append(perf(blendq))
    if i == 0: sample = (actq, blendq)

def summ(name, arr):
    A = np.array(arr); m, lo, hi = np.median(A, 0), np.percentile(A, 5, 0), np.percentile(A, 95, 0)
    labs = ["CAGR", "Vol", "DownVol", "MaxDD", "Sharpe", "Sortino", "Calmar"]
    print(f"\n{name} (median [5-95%]):")
    for i, L in enumerate(labs):
        print(f"  {L:9s} {m[i]:7.1f}{'%' if i<4 else ''}   [{lo[i]:6.1f}, {hi[i]:6.1f}]")
print(f"J-curve: each deal -20% at entry then recovers; return kept | avg activist weight {w_act.mean()*100:.0f}%")
print(f"Core-30 real: CAGR {perf(core)[0]:.1f}%  MaxDD {perf(core)[3]:.1f}%")
summ("ACTIVIST bucket (J-curve)", act_m)
summ("FULL STRATEGY: dynamic 30-70 + J-curve activist", blend_m)

import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
actq, blendq = sample
fig, (a1, a2) = plt.subplots(2, 1, figsize=(11, 7.5), height_ratios=[3, 1.5], sharex=True)
for s, lab, c in [(actq, "Activist bucket (J-curve)", "#ff7f0e"), (blendq, "Full strategy (dynamic blend)", "#1f77b4"), (core, "Core-30", "#2ca02c")]:
    eq = (1+s).cumprod(); a1.plot(eq.index, eq.values, label=lab, color=c, lw=1.9)
    dd = eq/eq.cummax()-1; a2.plot(dd.index, dd*100, color=c, lw=1.4)
a1.set_yscale("log"); a1.legend(loc="upper left"); a1.grid(True, which="both", alpha=0.25); a1.set_ylabel("Growth of 1.0 (log)")
a1.set_title("Activist J-curve (-20% entry, return kept) + valuation-timed Core blend — one representative 30-yr path", fontsize=10.5)
a2.grid(True, alpha=0.25); a2.set_ylabel("drawdown (%)")
plt.tight_layout(); fig.savefig(SCR+"/jcurve_strategy.png", dpi=130); print("\nwrote jcurve_strategy.png")
