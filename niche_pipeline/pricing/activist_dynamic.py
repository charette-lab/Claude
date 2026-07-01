#!/usr/bin/env python3
"""activist_dynamic.py — valuation-timed Core + activist bucket (30-70%), full 30yr.

Per the manager: the activist bucket earns its DEAL IRR (~40%/yr ex-best deal), with an
~8% loss roughly once a decade (deep-bear years); the Core-30 is the flexing base; and the
activist bucket is sized 30-70% by how CHEAP the universe is -- high universe median engine
ER (cheap) -> size UP to 70%; low (expensive) -> down to 30%. Core marked to market (its
real -42% path); activist on its deal-outcome basis (+40% / -8%). Full history 2000-2026.

  python3 pricing/activist_dynamic.py
"""
import sys, os
import numpy as np, pandas as pd
from scipy.stats import skew
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from backtest_core30 import SCR
ACT_RET, LOSS = 0.40, -0.08                 # deal IRR (ex-best) ; decade loss
WMIN, WMAX = 0.30, 0.70

core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]; core.index = pd.to_datetime(core.index)
core = core[core.index.year <= 2026]
corA = (1+core).groupby(core.index.year).prod()-1
loss_years = set(corA[corA < -0.15].index)  # deep-bear years => the ~decade 8% loss (data-driven)
print(f"deep-bear (activist loss) years: {sorted(loss_years)}  -> {len(loss_years)} in {len(corA)} yrs (~1/decade)")

# universe cheapness signal: cross-sectional MEDIAN engine ER per year (winsorized)
dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet", columns=["Date", "er_total", "artifact"])
dec = dec[~dec.artifact.astype(bool)]; dec["year"] = pd.to_datetime(dec["Date"]).dt.year
dec["er"] = dec["er_total"].clip(-1, 5)
sig = dec.groupby("year")["er"].median()
sig = sig[(sig.index >= 2000) & (sig.index <= 2026)]
norm = (sig - sig.min())/(sig.max()-sig.min())
w_act = (WMIN + (WMAX-WMIN)*norm).shift(1).clip(WMIN, WMAX)   # lag 1yr: size this year on last year's valuation
print(f"universe median ER range {sig.min()*100:.0f}%-{sig.max()*100:.0f}% -> activist weight {w_act.min()*100:.0f}%-{w_act.max()*100:.0f}%")

act_annual = {y: (LOSS if y in loss_years else ACT_RET) for y in range(2000, 2027)}
def act_q(q): return (1+act_annual.get(q.year, ACT_RET))**0.25-1
def wt(q): return float(w_act.get(q.year, 0.5)) if not pd.isna(w_act.get(q.year, np.nan)) else 0.5

C = core[core.index.year >= 2000]
books = {}
books["Core-30 only"] = C
books["Activist only (fundamental)"] = pd.Series({q: act_q(q) for q in C.index})
books["Static 50/50"] = pd.Series({q: 0.5*act_q(q)+0.5*C[q] for q in C.index})
books["Static 70/30 (max activist)"] = pd.Series({q: 0.7*act_q(q)+0.3*C[q] for q in C.index})
books["DYNAMIC 30-70 (valuation-timed)"] = pd.Series({q: wt(q)*act_q(q)+(1-wt(q))*C[q] for q in C.index})

def perf(r, ppy=4):
    r = pd.Series(r).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    dds = np.sqrt((np.minimum(r, 0.0)**2).mean())*np.sqrt(ppy); cagr = eq.iloc[-1]**(1/y)-1
    return dict(CAGR=cagr*100, Vol=r.std(ddof=1)*np.sqrt(ppy)*100, DownVol=dds*100, MaxDD=dd.min()*100,
                Sharpe=(r.mean()*ppy-0.03)/(r.std(ddof=1)*np.sqrt(ppy)), Sortino=(r.mean()*ppy-0.03)/dds if dds > 0 else np.nan,
                Calmar=cagr/abs(dd.min()) if dd.min() < 0 else np.nan)
T = pd.DataFrame({k: perf(v) for k, v in books.items()}).T
print(f"\nperiod {C.index.min():%Y}-{C.index.max():%Y} | activist {ACT_RET*100:.0f}%/yr, {LOSS*100:.0f}% in loss years, weight {WMIN*100:.0f}-{WMAX*100:.0f}% by valuation\n")
print(T[["CAGR", "Vol", "DownVol", "MaxDD", "Sharpe", "Sortino", "Calmar"]].round(2).to_string())
# average activist weight, and the weight in a couple of key years
print(f"\n  avg activist weight {w_act.mean()*100:.0f}% | e.g. 2003 (post-crash) {w_act.get(2003,np.nan)*100:.0f}% vs 2021 (peak) {w_act.get(2021,np.nan)*100:.0f}%")
print("\nCAVEAT: 40% is the DEAL IRR (return on fully-deployed capital over a deal's life). It equals a 40%")
print("annual BUCKET return only if you stay fully deployed -- your realized fund did ~18% gross partly")
print("from cash drag. So this is the 'fully-deployed' upper case; the deep-bear -42% Core path is still MTM real.")
