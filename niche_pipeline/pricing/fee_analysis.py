#!/usr/bin/env python3
"""fee_analysis.py — the high-water-mark cost, in cash.

Runs the manager's ACTUAL gross monthly returns (2006-2026) through the real
fee structure (1.75% mgmt, 20% performance, annual high-water-mark crystallization),
then does the same for the recommended structure — systematic Resilient Core-30 as
the spine + the manager's own book as a bounded satellite — over the window where the
systematic core exists (2016-2026). Identical fee engine on both, so the comparison
isolates what the trajectory costs through the HWM, in money rather than in Sharpe.

  python3 pricing/fee_analysis.py
"""
import numpy as np, pandas as pd
SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
N = None
# ---- manager's ACTUAL gross monthly returns, % (from the provided table) ----
G = {
 2006:[5.8,4.6,0.9,6.1,2.2,-5.6,2.3,5.4,9.4,8.6,11.1,0.4],
 2007:[3.3,0.3,4.3,4.0,1.7,-3.9,-2.0,-2.4,-4.8,-2.4,-4.6,-1.0],
 2008:[-10.8,7.3,-0.2,-4.6,2.2,-17.8,0.9,3.9,-13.8,-12.6,-0.3,-1.4],
 2009:[0.6,2.2,3.1,21.7,6.4,-1.2,8.3,8.7,3.6,4.8,2.9,2.3],
 2010:[2.0,-1.9,7.2,9.5,-8.6,-2.6,7.4,-5.5,13.3,-0.8,1.0,6.8],
 2011:[-1.8,-1.0,5.7,-0.2,-2.3,-4.1,-3.1,-4.0,-3.9,5.3,-3.6,14.3],
 2012:[6.0,11.2,-5.0,-0.7,-6.7,-1.7,-2.8,2.1,4.0,-3.4,1.9,1.7],
 2013:[7.9,6.7,1.8,-4.0,5.7,-2.8,8.5,-0.1,5.5,4.0,2.7,0.8],
 2014:[6.7,7.6,0.8,2.4,1.8,1.2,N,N,N,N,N,N],
 2015:[N,0.4,2.9,2.0,3.1,-3.4,2.3,-3.7,-0.6,1.2,6.4,2.6],
 2016:[-4.7,1.4,0.4,4.8,3.5,-2.9,4.2,9.5,3.9,1.5,5.2,5.1],
 2017:[6.9,5.9,0.8,10.6,-3.6,-1.1,-1.0,-4.8,2.9,-4.4,-0.6,-3.3],
 2018:[5.7,0.0,-1.9,1.4,4.7,-0.5,0.6,-0.9,-2.2,-4.6,-6.0,-7.5],
 2019:[5.9,1.6,-3.8,4.6,-7.0,0.4,-5.2,-9.6,1.8,-0.3,5.0,3.4],
 2020:[0.39,-8.27,-18.36,11.71,-0.05,3.20,2.75,10.17,2.77,-5.72,12.81,9.90],
 2021:[5.15,109.55,3.28,2.16,-0.35,0.63,-0.27,1.31,29.26,11.80,2.68,1.76],
 2022:[-1.07,-1.53,0.64,-2.76,-0.52,-3.72,0.54,-3.46,-5.14,6.43,14.78,-0.54],
 2023:[0.15,-6.37,-6.81,5.02,-0.02,6.57,5.90,-2.43,-1.59,-3.15,1.33,3.73],
 2024:[3.14,3.45,6.50,1.66,2.05,-0.79,7.49,-1.57,0.35,0.04,6.58,0.79],
 2025:[4.50,0.08,-2.07,-1.61,8.96,2.38,3.22,-1.65,-0.47,-3.48,-0.82,-1.63],
 2026:[-2.02,-0.28,-2.29,1.72,2.08,N,N,N,N,N,N,N],
}
MGMT, PERF = 0.0175, 0.20

def monthly_series():
    idx, val = [], []
    for y in sorted(G):
        for mo in range(12):
            r = G[y][mo]
            if r is None: continue
            idx.append(pd.Timestamp(y, mo+1, 1) + pd.offsets.MonthEnd(0)); val.append(r/100.0)
    return pd.Series(val, index=pd.DatetimeIndex(idx)).sort_index()

def to_quarterly(m):
    return (1+m).groupby(m.index.to_period("Q")).prod()-1

def to_annual(m):
    return (1+m).groupby(m.index.year).prod()-1

def fee_run(returns, periods_per_year, year_of):
    """Apply mgmt (accrued each period) + performance (crystallized at year-end,
    HWM). `returns` is an ordered iterable of (key, gross_return). Returns net
    series + fee ledger."""
    nav, hwm = 1.0, 1.0
    cum_m, cum_p = 0.0, 0.0
    net_r, keys = [], []
    yrs_perf, yrs_under = set(), set()
    items = list(returns)
    for i, (k, g) in enumerate(items):
        yr = year_of(k)
        beg = nav
        m = MGMT/periods_per_year*beg          # mgmt on beginning NAV
        nav = beg*(1+g) - m
        cum_m += m
        last_of_year = (i == len(items)-1) or (year_of(items[i+1][0]) != yr)
        if last_of_year:
            excess = nav - hwm
            if excess > 0:
                p = PERF*excess; nav -= p; cum_p += p; yrs_perf.add(yr)
            else:
                yrs_under.add(yr)
            hwm = max(hwm, nav)
        net_r.append(nav/beg-1); keys.append(k)
    return pd.Series(net_r, index=keys), dict(mgmt=cum_m, perf=cum_p,
            final=nav, yrs_perf=len(yrs_perf), yrs_under=len(yrs_under))

def stats(net, ppy):
    r = pd.Series(net).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy
    dd = eq/eq.cummax()-1; dn = r[r < 0]
    cagr = eq.iloc[-1]**(1/y)-1
    return dict(CAGR=cagr, Vol=r.std(ddof=1)*np.sqrt(ppy), MaxDD=dd.min(),
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)),
                Sortino=r.mean()*ppy/(dn.std(ddof=1)*np.sqrt(ppy)) if len(dn) else np.nan,
                Calmar=cagr/abs(dd.min()), final=eq.iloc[-1])

m = monthly_series()
ann = to_annual(m); q = to_quarterly(m)
q.index = q.index.to_timestamp("Q")   # period -> quarter-end timestamp

# ========== PART A: manager's actual book, full 2006-2025, annual crystallization ==========
print("="*84)
print("PART A  —  Athanase actual book, gross vs net, full record (annual crystallization)")
print("="*84)
grossA = (1+ann).prod()**(1/(len(ann)))-1
netA, ledgerA = fee_run(((y, ann[y]) for y in ann.index), 1, lambda k: k)
sA_g = stats(ann, 1); sA_n = stats(netA, 1)
print(f"  years in record: {ann.index.min()}-{ann.index.max()}  ({len(ann)} annual obs; 2014 H1 & 2026 partial)")
print(f"  GROSS  CAGR {sA_g['CAGR']*100:6.1f}%   MaxDD {sA_g['MaxDD']*100:6.1f}%   final x{sA_g['final']:6.1f}")
print(f"  NET    CAGR {sA_n['CAGR']*100:6.1f}%   MaxDD {sA_n['MaxDD']*100:6.1f}%   final x{sA_n['final']:6.1f}")
print(f"  fee drag: {(sA_g['CAGR']-sA_n['CAGR'])*100:.1f} pts/yr of CAGR")
print(f"  cumulative mgmt fees  : {ledgerA['mgmt']*100:6.1f}%  of starting capital")
print(f"  cumulative perf fees  : {ledgerA['perf']*100:6.1f}%  of starting capital")
print(f"  years a perf fee was EARNED (above HWM): {ledgerA['yrs_perf']} of {len(ann)}")
print(f"  years BELOW high-water mark (zero carry): {ledgerA['yrs_under']} of {len(ann)}")

# ========== PART B: recommended structure over the systematic-core window (2016Q2-2026Q2) ==========
core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]
core.index = pd.to_datetime(core.index)
# align quarters (match on year+quarter, tolerate day-of-quarter-end differences)
cp = core.copy(); cp.index = cp.index.to_period("Q")
tp = q.copy();   tp.index = q.index.to_period("Q")
both = pd.DataFrame({"core": cp, "book": tp}).dropna()
print("\n"+"="*84)
print(f"PART B  —  recommended structure vs the book, {both.index.min()}-{both.index.max()} "
      f"({len(both)} quarters, where the systematic core exists)")
print("="*84)
print("  Both series are GROSS; identical fee engine applied to each. Core is the systematic")
print("  Resilient Core-30 (quality floor + caps + inverse-vol); 'book' is the actual returns.")
mixes = {"Book only (100% actual)": (0.00, 1.00),
         "Blend 75/25 (Core spine + book satellite)": (0.75, 0.25),
         "Blend 90/10 (Core spine + small satellite)": (0.90, 0.10),
         "Core only (100% systematic)": (1.00, 0.00)}
rows = {}
navs = {}
yq = lambda p: p.year
for name, (wc, wb) in mixes.items():
    gross = wc*both["core"] + wb*both["book"]
    net, led = fee_run(((k, gross[k]) for k in gross.index), 4, yq)
    s = stats(net, 4)
    rows[name] = {**s, **{"mgmt%": led["mgmt"]*100, "perf%": led["perf"]*100,
                          "yrs_perf": led["yrs_perf"], "yrs_under": led["yrs_under"]}}
    navs[name] = (1+net).cumprod()
tb = pd.DataFrame(rows).T
disp = pd.DataFrame({
    "CAGR%": (tb["CAGR"]*100).round(1), "Vol%": (tb["Vol"]*100).round(1),
    "MaxDD%": (tb["MaxDD"]*100).round(1), "Sharpe": tb["Sharpe"].round(2),
    "Sortino": tb["Sortino"].round(2), "Calmar": tb["Calmar"].round(2),
    "final x": tb["final"].round(2), "perfFee%": tb["perf%"].round(0),
    "yrs<HWM": tb["yrs_under"].astype(int)})
print(disp.to_string())
nyears = len(both)/4
print(f"\n  (window = {nyears:.1f} years; 'perfFee%' = cumulative performance fee as % of starting capital;")
print(f"   'yrs<HWM' = calendar years the book was below its high-water mark and earned zero carry)")

# chart
import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
fig, (a1, a2) = plt.subplots(2, 1, figsize=(11, 8.5), height_ratios=[3, 1.5], sharex=True)
col = {"Book only (100% actual)": "#d62728", "Blend 75/25 (Core spine + book satellite)": "#1f77b4",
       "Blend 90/10 (Core spine + small satellite)": "#2ca02c", "Core only (100% systematic)": "#7f7f7f"}
x0 = both.index.min().to_timestamp("Q") - pd.offsets.QuarterEnd()
for name, nav in navs.items():
    idx = pd.DatetimeIndex([x0]).append(pd.PeriodIndex(nav.index).to_timestamp("Q"))
    y = np.r_[1.0, nav.values]
    a1.plot(idx, y, label=name, color=col[name], lw=2.4 if "Blend 75" in name or "Book" in name else 1.6)
    dd = y/np.maximum.accumulate(y)-1
    a2.plot(idx, dd*100, color=col[name], lw=1.7 if "Blend 75" in name or "Book" in name else 1.2)
a1.set_yscale("log"); a1.legend(loc="upper left", fontsize=9); a1.grid(True, which="both", alpha=0.25)
a1.set_ylabel("Growth of 1.0 NET of fees (log)")
a1.set_title("NET-of-fee outcomes under the actual 1.75% / 20% / high-water-mark structure\n"
             "actual book vs Core-spine + book-satellite blends (2016Q2-2026Q2, gross inputs, identical fee engine)", fontsize=10.5)
a2.grid(True, alpha=0.25); a2.set_ylabel("drawdown (%)"); a2.axhline(0, color="k", lw=0.5)
plt.tight_layout(); fig.savefig(SCR+"/fee_analysis.png", dpi=130)
print("\nwrote fee_analysis.png")
