#!/usr/bin/env python3
"""satellite_blend.py — Core base + activist Satellite: a realistic return with risk metrics.

Satellite = each year, the top-15 TAKEABLE activist names by freed-core er (from
activist_funnel), equal-weight, annual rebalance; quarterly returns from prices. That is
the PASSIVE capture (no intervention) -- the floor. Core = Resilient Core-30 (quarterly).
We report the real backtest for Core, Satellite(passive), and Core+Satellite blends, then a
REALISTIC-activist scenario: lift the satellite to the manager's demonstrated activist level
(~20% gross) by a constant per-quarter premium (keeps the risk shape), and re-blend. All
gross, price-only; metrics: CAGR, Vol, Sharpe, Sortino, Calmar, MaxDD.

  python3 pricing/satellite_blend.py
"""
import sys, os
import numpy as np, pandas as pd
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from backtest_core30 import SCR
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
PRICE = [U+"972f0581-daily_volume_price_0.parquet", U+"257124b3-daily_volume_price_1.parquet",
         U+"86e54ec3-daily_volume_price_2.parquet", U+"13f82c18-daily_volume_price_0630.parquet"]
N_SAT = 15
TARGET_ACTIVIST = 0.20        # realistic satellite CAGR anchored to the manager's demonstrated activist returns

px = pd.concat([pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"]) for f in PRICE])
px["Instrument"] = px["Instrument"].astype(str); px["Date"] = pd.to_datetime(px["Date"])
px["c"] = pd.to_numeric(px["Close Price"], errors="coerce")
W = px[px.c > 0].pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()
qends = pd.Series(W.index).groupby(pd.Series(W.index).dt.to_period("Q")).max().values
Wq = W.reindex(pd.DatetimeIndex(sorted(qends)), method="ffill")
Rq = Wq.pct_change()          # quarterly returns

A = pd.read_parquet(SCR+"/activist_funnel.parquet")
sat_rows = []
for y in range(2016, 2026):
    sub = A[(A.year == y) & (A.takeable)].dropna(subset=["freed_15%"])
    picks = sub.sort_values("freed_15%", ascending=False).head(N_SAT)["t"].tolist()
    picks = [p for p in picks if p in Rq.columns]
    if not picks:
        continue
    for q in Rq.index:
        if q.year == y:                                   # hold the year-y book through year y
            r = Rq.loc[q, picks].dropna()
            if len(r):
                sat_rows.append((q, float(r.mean())))
sat = pd.Series(dict(sat_rows)).sort_index(); sat.name = "Satellite_passive"

core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]
core.index = pd.to_datetime(core.index).to_period("Q").to_timestamp("Q")
sat.index = pd.to_datetime(sat.index).to_period("Q").to_timestamp("Q")
df = pd.concat([core.rename("Core"), sat], axis=1).dropna()

def perf(r, ppy=4):
    r = pd.Series(r).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    dn = r[r < 0]
    cagr = eq.iloc[-1]**(1/y)-1
    return dict(CAGR=cagr*100, Vol=r.std(ddof=1)*np.sqrt(ppy)*100, MaxDD=dd.min()*100,
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)),
                Sortino=r.mean()*ppy/(dn.std(ddof=1)*np.sqrt(ppy)) if len(dn) > 1 else np.nan,
                Calmar=cagr/abs(dd.min()))

# realistic activist satellite: constant per-quarter premium to hit TARGET (keeps risk shape)
base_cagr = (1+df.Satellite_passive).prod()**(4/len(df))-1
prem_q = (1+TARGET_ACTIVIST)**0.25 - (1+base_cagr)**0.25
sat_real = df.Satellite_passive + prem_q

books = {"Core-30 (base)": df.Core, "Satellite passive (no intervention)": df.Satellite_passive,
         "Satellite realistic (~20% activist)": sat_real,
         "Blend 90/10 (realistic)": 0.90*df.Core + 0.10*sat_real,
         "Blend 80/20 (realistic)": 0.80*df.Core + 0.20*sat_real,
         "Blend 75/25 (realistic)": 0.75*df.Core + 0.25*sat_real,
         "Blend 75/25 (passive floor)": 0.75*df.Core + 0.25*df.Satellite_passive}
T = pd.DataFrame({k: perf(v) for k, v in books.items()}).T
print(f"quarters: {len(df)} ({df.index.min():%Y}-{df.index.max():%Y}) | satellite = top-{N_SAT} takeable activist names/yr, EW\n")
print(f"  satellite PASSIVE realized CAGR: {base_cagr*100:.1f}%  (the no-intervention floor)")
print(f"  realistic activist target:       {TARGET_ACTIVIST*100:.0f}%  (anchored to your demonstrated activist returns)\n")
show = T[["CAGR", "Vol", "MaxDD", "Sharpe", "Sortino", "Calmar"]].round(2)
print(show.to_string())
print("\nCAVEATS: gross, price-only, local-ccy (absolute CAGRs optimistic ~by the dividend yield +);")
print("2016-26 benign window (no secular bear); survivorship; activist premium is an assumption, not")
print("a backtest -- realized activism carries execution/time risk. Treat RANKINGS as robust, LEVELS as high.")
