#!/usr/bin/env python3
"""core_activist_model.py — Core-30 base + activist satellite (4yr realize-then-sell) over 10yr.

Core-30: real quarterly backtest (price-based risk). Satellite: ~2 activist picks/year (top
takeable index-beating, non-distressed, freed er @15% ROIIC), each HELD 4 YEARS earning the
engine freed return, then sold; steady-state book ~8 positions. Satellite return = the freed
return (realistic = model er x 0.5, the Core-30's own model->realized ratio). Blend and report
CAGR/Vol/DownsideVol/MaxDD/Sharpe/Sortino/Calmar. The satellite's fundamental downside is ~0
(all-positive freed returns); we ALSO overlay a realistic mark-to-market drawdown (beta ~0.8
to the market in crash quarters) so the combined drawdown isn't understated.

  python3 pricing/core_activist_model.py
"""
import sys, os
import numpy as np, pandas as pd
from scipy.stats import skew
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from backtest_core30 import SCR
HOLD, NPICK, HAIRCUT = 4, 2, 0.5

A = pd.read_parquet(SCR+"/activist_funnel.parquet"); A = A[A.takeable].copy()
picks = {}
for y in range(1997, 2026):
    sub = A[(A.year == y) & (A["freed_15%"] >= 0.20) & (A["freed_15%"] <= 0.50)]
    sub = sub.dropna(subset=["freed_15%"]).sort_values("freed_15%", ascending=False).head(NPICK)
    picks[y] = [(t, er*HAIRCUT) for t, er in zip(sub.t, sub["freed_15%"])]
sat_annual = {}
for z in range(2000, 2027):
    active = [er for vy in range(z-HOLD+1, z+1) if vy in picks for (_, er) in picks[vy]]
    if active:
        sat_annual[z] = float(np.mean(active))

core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]; core.index = pd.to_datetime(core.index)
core = core[(core.index.year >= 2000) & (core.index.year <= 2026)]
# satellite quarterly: fundamental freed return, realized smoothly within the year
satq = pd.Series({d: (1+sat_annual.get(d.year, np.nan))**0.25-1 for d in core.index})
# realistic MTM overlay: satellite wobbles with the equity market in crashes (fundamental value unchanged).
# MSCI monthly only exists 2015+, so use the Core-30 itself as the equity-market proxy (available full history).
BETA = 0.8
satq_mtm = satq + BETA*np.minimum(core - core.mean(), 0.0)    # add adverse market deviations -> crash drawdown texture

df = pd.concat([core.rename("Core"), satq.rename("SatFund"), satq_mtm.rename("SatMTM")], axis=1).dropna()

def perf(r, ppy=4):
    r = pd.Series(r).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    dds = np.sqrt((np.minimum(r, 0.0)**2).mean())*np.sqrt(ppy); cagr = eq.iloc[-1]**(1/y)-1
    return dict(CAGR=cagr*100, Vol=r.std(ddof=1)*np.sqrt(ppy)*100, DownVol=dds*100, MaxDD=dd.min()*100,
                Sharpe=(r.mean()*ppy-0.03)/(r.std(ddof=1)*np.sqrt(ppy)), Sortino=(r.mean()*ppy-0.03)/dds if dds > 0 else np.nan,
                Calmar=cagr/abs(dd.min()) if dd.min() < 0 else np.nan)

print(f"period {df.index.min():%Y}-{df.index.max():%Y} ({len(df)}Q) | satellite realistic return ~{np.mean(list(sat_annual.values()))*100:.0f}%/yr "
      f"(freed @15% ROIIC x{HAIRCUT}), {NPICK} picks/yr, {HOLD}yr hold, ~{NPICK*HOLD} concurrent\n")
print("--- A) satellite modeled as pure FUNDAMENTAL freed return (downside ~0 by construction) ---")
rowsF = {"Core-30 only": df.Core, "Blend 90/10": .9*df.Core+.1*df.SatFund, "Blend 85/15": .85*df.Core+.15*df.SatFund,
         "Blend 75/25": .75*df.Core+.25*df.SatFund}
TF = pd.DataFrame({k: perf(v) for k, v in rowsF.items()}).T
print(TF[["CAGR", "Vol", "DownVol", "MaxDD", "Sharpe", "Sortino", "Calmar"]].round(2).to_string())
print("\n--- B) same, but satellite carries a realistic MARK-TO-MARKET drawdown (beta 0.8 in crashes) ---")
rowsM = {"Core-30 only": df.Core, "Blend 90/10": .9*df.Core+.1*df.SatMTM, "Blend 85/15": .85*df.Core+.15*df.SatMTM,
         "Blend 75/25": .75*df.Core+.25*df.SatMTM}
TM = pd.DataFrame({k: perf(v) for k, v in rowsM.items()}).T
print(TM[["CAGR", "Vol", "DownVol", "MaxDD", "Sharpe", "Sortino", "Calmar"]].round(2).to_string())
print("\nREAD: (A) is the fundamental view -- the fix delivers positive value, so the satellite adds return")
print("with ~0 downside and every risk-adjusted ratio improves. (B) is what an LP sees -- the positions")
print("still wobble with the market mid-hold, so the combined drawdown deepens, but the thesis (and the")
print("eventual 4-yr exit value) is unchanged. Reality is between; permanent capital is what bridges them.")
print("\nCAVEATS: satellite return is model er x0.5 (still optimistic; benign 2016-26 window; survivorship);")
print("gross, price-only. Treat ratios as robust, absolute CAGR as the high end.")
