#!/usr/bin/env python3
"""count_opportunities.py — how many investable names per year clear the framework AND
beat the Core Index expected return?

Goes through the daily full-engine decomposition (er_total per security-day, 2016-2026),
applies the framework eligibility (non-artifact, moat>=6.5, carries risk tags, no Extreme
severity), and counts DISTINCT securities per year whose expected return persistently
exceeds the Core Index's own expected return (opportunity cost of just buying the index).
'Persistently' = qualifying on >=60% of the year's trading days (not a one-day blip).

  python3 pricing/count_opportunities.py
"""
import sys, os, json
import numpy as np, pandas as pd
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from backtest_core30 import load_meta, SCR

MIN_MOAT = 6.5
PERSIST = 0.60            # must qualify on >=60% of the year's days
MIN_DAYS = 120           # and trade most of the year

tags, maxsev, moat, country = load_meta()
dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet", columns=["Instrument", "Date", "er_total", "artifact"])
dec["Instrument"] = dec["Instrument"].astype(str); dec["Date"] = pd.to_datetime(dec["Date"])
dec["year"] = dec["Date"].dt.year

# framework eligibility (static attributes)
elig = {r for r in set(dec.Instrument) if moat.get(r, 0) >= MIN_MOAT and r in tags and maxsev.get(r, 0) < 4}
dec = dec[dec.Instrument.isin(elig) & (~dec.artifact.astype(bool))]

# Core Index expected return per year = median er of the top-30 eligible names by er each day, averaged
hurdle = {}
for y, g in dec.groupby("year"):
    # per day, median er of that day's top-30 eligible names
    top = g.groupby("Date")["er_total"].apply(lambda s: s.nlargest(30).median())
    hurdle[y] = float(top.median())
print("Core Index expected-return hurdle by year (median er of the top-30 eligible names):")
print("  " + "  ".join(f"{y}:{hurdle[y]*100:.0f}%" for y in sorted(hurdle)))
avg_hurdle = np.mean(list(hurdle.values()))
print(f"  average hurdle ~{avg_hurdle*100:.0f}%   (realized Core-30 CAGR was ~16.7%)\n")

# per (year, security): fraction of days above each threshold
def count_year(thr_fn, label):
    out = {}
    for y, g in dec.groupby("year"):
        gg = g.copy(); gg["ok"] = gg["er_total"] > thr_fn(y)
        per = gg.groupby("Instrument").agg(frac=("ok", "mean"), n=("ok", "size"))
        avail = per[(per.frac >= PERSIST) & (per.n >= MIN_DAYS)]
        out[y] = len(avail)
    return out

g1 = count_year(lambda y: 0.12, "Gate1 12%")
bi = count_year(lambda y: hurdle[y], "beat Core Index")
h20 = count_year(lambda y: 0.20, ">=20%")
h25 = count_year(lambda y: 0.25, ">=25%")
uni = {y: g.Instrument.nunique() for y, g in dec.groupby("year")}

print(f"{'year':6}{'universe':>9}{'pass 12%':>10}{'BEAT INDEX':>12}{'>=20%':>8}{'>=25%':>8}")
yrs = [y for y in sorted(uni) if 2016 <= y <= 2025]
for y in yrs:
    print(f"{y:<6}{uni[y]:>9}{g1[y]:>10}{bi[y]:>12}{h20[y]:>8}{h25[y]:>8}")
mean = lambda d: np.mean([d[y] for y in yrs])
print(f"\n  per-year AVERAGE:  eligible universe {mean(uni):.0f} | pass 12% {mean(g1):.0f} | "
      f"BEAT CORE INDEX {mean(bi):.0f} | >=20% {mean(h20):.0f} | >=25% {mean(h25):.0f}")
print(f"\n  So the framework surfaces ~{mean(bi):.0f} names/yr that persistently beat the Core Index on")
print(f"  EXPECTED return. Note: that's expected-return dominance only; beating the index net of")
print(f"  concentration/idiosyncratic risk (and after your DD) narrows it further -- consistent with")
print(f"  doing 1-2 high-conviction deals from a much larger opportunity set.")
