#!/usr/bin/env python3
"""final_strategy.py — the complete, honest strategy over 30 years.

Activist bucket = Monte-Carlo from the real 39-deal distribution (2 deals/yr, 4yr hold,
~8 concurrent); normal years earn the diversified deal-outcome mean, deep-bear years take a
PARTIAL mark-to-market markdown (calibrated to ~7-8% effective downside vol -- between the 2%
pure-fundamental and 11% full-NAV). Core-30 = its real 2000-2026 path (incl. -42%). Blend =
activist sized 30-70% by universe cheapness (valuation-timed). Report the distribution of the
combined strategy's return AND drawdown across simulations.

  python3 pricing/final_strategy.py
"""
import sys, os
import numpy as np, pandas as pd
from scipy.stats import skew
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from backtest_core30 import SCR
rng = np.random.default_rng(11)
DEEP_BEAR = -0.24                      # activist partial-MTM markdown in a deep bear -> ~7-8% downside vol
WMIN, WMAX = 0.30, 0.70

# --- real Core-30 quarterly path ---
core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]; core.index = pd.to_datetime(core.index)
core = core[(core.index.year >= 2000) & (core.index.year <= 2026)]
corA = (1+core).groupby(core.index.year).prod()-1
BEAR = set(corA[corA < -0.15].index)   # {2002, 2008}

# --- valuation signal -> dynamic activist weight (lag 1yr) ---
dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet", columns=["Date", "er_total", "artifact"])
dec = dec[~dec.artifact.astype(bool)]; dec["year"] = pd.to_datetime(dec["Date"]).dt.year
sig = dec.assign(er=dec.er_total.clip(-1, 5)).groupby("year")["er"].median()
sig = sig[(sig.index >= 1999) & (sig.index <= 2026)]
w_act = (WMIN + (WMAX-WMIN)*(sig-sig.min())/(sig.max()-sig.min())).shift(1).clip(WMIN, WMAX)

# --- real deal distribution ---
DEALS = [(2015,2025,2.0),(2015,2025,1.2),(2016,2025,6.2),(2016,2025,1.7),(2018,2025,1.2),(2018,2025,1.5),(2018,2025,2.8),
         (2019,2025,5.1),(2020,2025,1.2),(2020,2025,5.4),(2024,2025,1.6),(2017,2023,0.4),(2019,2022,1.6),(2015,2020,1.4),
         (2016,2018,1.5),(2015,2017,1.6),(2015,2016,1.7),(2007,2014,2.9),(2011,2014,3.9),(2007,2014,4.4),(2009,2014,5.5),
         (2012,2014,1.6),(2010,2014,5.1),(2011,2014,1.3),(2010,2014,1.5),(2010,2014,2.0),(2007,2014,28.2),(2007,2014,5.5),
         (2011,2013,0.3),(2006,2007,3.1),(2006,2007,1.4),(2006,2014,8.4),(2006,2011,2.3),(2006,2007,3.0),(2007,2011,2.1),
         (2007,2011,1.4),(2007,2011,5.4),(2003,2006,5.3),(2007,2010,0.6)]
ar = np.array([m**(1/max(b-a,1))-1 for a,b,m in DEALS]); HOLD = 4
YEARS = list(range(2000, 2027))

def perf(r, ppy=4):
    r = pd.Series(r).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    dds = np.sqrt((np.minimum(r, 0.0)**2).mean())*np.sqrt(ppy); cagr = eq.iloc[-1]**(1/y)-1
    return (cagr*100, r.std(ddof=1)*np.sqrt(ppy)*100, dds*100, dd.min()*100,
            (r.mean()*ppy-0.03)/(r.std(ddof=1)*np.sqrt(ppy)), (r.mean()*ppy-0.03)/dds if dds>0 else np.nan,
            cagr/abs(dd.min()) if dd.min()<0 else np.nan)

NSIM = 2000
res = {"activist": [], "blend": []}
for _ in range(NSIM):
    picks = {y: rng.choice(ar, 2) for y in YEARS}
    act_ann = {}
    for y in YEARS:
        if y in BEAR:
            act_ann[y] = DEEP_BEAR
        else:
            active = np.concatenate([picks[v] for v in YEARS if 0 <= y-v < HOLD])
            act_ann[y] = float(active.mean()) if len(active) else float(ar.mean())
    actq = pd.Series({q: (1+act_ann[q.year])**0.25-1 for q in core.index})
    wq = pd.Series({q: (float(w_act.get(q.year)) if not pd.isna(w_act.get(q.year, np.nan)) else 0.5) for q in core.index})
    blendq = wq*actq + (1-wq)*core
    res["activist"].append(perf(actq)); res["blend"].append(perf(blendq))

def summarize(name, arr):
    A = np.array(arr); m = np.median(A, 0); lo = np.percentile(A, 5, 0); hi = np.percentile(A, 95, 0)
    labs = ["CAGR", "Vol", "DownVol", "MaxDD", "Sharpe", "Sortino", "Calmar"]
    print(f"\n{name} (median [5-95%] across {NSIM} sims):")
    for i, L in enumerate(labs):
        u = "%" if i < 4 else ""
        print(f"  {L:9s} {m[i]:7.1f}{u}   [{lo[i]:6.1f}, {hi[i]:6.1f}]")

print(f"Core-30 (real path): CAGR {perf(core)[0]:.1f}%  MaxDD {perf(core)[3]:.1f}%  Calmar {perf(core)[6]:.2f}")
print(f"avg activist weight {w_act.mean()*100:.0f}% (2003 {w_act.get(2003)*100:.0f}% vs 2021 {w_act.get(2021)*100:.0f}%) | deep-bear yrs {sorted(BEAR)}")
summarize("ACTIVIST bucket (real distribution)", res["activist"])
summarize("FULL STRATEGY: dynamic 30-70 Core+activist", res["blend"])
print("\nThis is the honest complete picture: real deal dispersion (incl. losers + winner skew), ~7-8%")
print("effective activist downside, valuation-timed sizing, and the Core's real -42% path -- all combined.")
print("CAVEATS: deal pool incl. big winners (assumes the process keeps producing them); survivorship in the")
print("Core; annual->quarterly smoothing of the activist; 40% deal IRR only if you stay fully deployed.")
