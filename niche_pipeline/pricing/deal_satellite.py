#!/usr/bin/env python3
"""deal_satellite.py — satellite risk from the REAL deal outcomes, not stock prices.

The activist earns the deal outcome (earnings rise over the hold -> value realized), which
the passive stock-price path misrepresents. So build the satellite from the manager's actual
~2 deals/year: each deal earns its annualized return (MOIC^(1/years)-1) over its holding
period; the book each year = invested-weighted return of the deals active that year. Then
measure the volatility and DOWNSIDE volatility of that fundamentally-driven series, and blend
onto the Core-30.

  python3 pricing/deal_satellite.py
"""
import sys, os
import numpy as np, pandas as pd
from scipy.stats import skew
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from backtest_core30 import SCR
# (name, entry, exit, invested, returned, MOIC)
DEALS = [
 ("Bufab",2015,2025,54,107,2.0),("DistIt",2015,2025,77,94,1.2),("Ngex",2016,2025,8,51,6.2),("Alcadon",2016,2025,38,64,1.7),
 ("Tempest",2018,2025,20,23,1.2),("Zalaris",2018,2025,50,73,1.5),("Zutec",2018,2025,37,103,2.8),("Roko",2019,2025,26,134,5.1),
 ("Catella",2020,2025,71,87,1.2),("AI",2020,2025,113,612,5.4),("Renold",2024,2025,35,56,1.6),("Actic",2017,2023,81,34,0.4),
 ("Haldex1",2019,2022,22,36,1.6),("Robit",2015,2020,54,73,1.4),("Kitron",2016,2018,43,63,1.5),("Sandvik",2015,2017,17,27,1.6),
 ("Transcom1",2015,2016,21,36,1.7),
 ("Bilia",2007,2014,21,62,2.9),("Concentric",2011,2014,93,364,3.9),("eWork",2007,2014,30,133,4.4),("Haldex2",2009,2014,85,467,5.5),
 ("Lindab",2012,2014,448,702,1.6),("Note",2010,2014,37,186,5.1),("Transcom2",2011,2014,128,170,1.3),("Global",2010,2014,142,214,1.5),
 ("Ferronordic",2010,2014,50,102,2.0),("Klarna",2007,2014,30,836,28.2),("Acne",2007,2014,70,386,5.5),("Usports",2011,2013,96,30,0.3),
 ("Catena",2006,2007,88,270,3.1),("Tigerholm",2006,2007,115,157,1.4),("Avanza",2006,2014,313,2614,8.4),("Fabege",2006,2011,849,1942,2.3),
 ("JohnsonPump",2006,2007,71,213,3.0),("Klovern",2007,2011,110,227,2.1),("Nobia",2007,2011,490,668,1.4),("SkiStar",2007,2011,99,534,5.4),
 ("Brokk",2003,2006,120,630,5.3),("HQ",2007,2010,684,393,0.6)]

rows = []
for n, y0, y1, inv, ret, moic in DEALS:
    yrs = max(y1-y0, 1); ar = moic**(1/yrs)-1
    rows.append(dict(name=n, y0=y0, y1=y1, inv=inv, ar=ar))
D = pd.DataFrame(rows)

# invested-weighted annual book return: deals active in year Y each earn their annualized return
sat = {}
for Y in range(2006, 2026):
    act = D[(D.y0 <= Y) & (D.y1 > Y)]                       # active during year Y (entry..exit-1)
    if len(act) == 0:
        continue
    w = act.inv/act.inv.sum()
    sat[Y] = float((w*act.ar).sum())
sat = pd.Series(sat).sort_index(); sat.name = "DealSatellite"

def perf(r, ppy=1):
    r = pd.Series(r).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    dds = np.sqrt((np.minimum(r, 0.0)**2).mean())*np.sqrt(ppy)
    cagr = eq.iloc[-1]**(1/y)-1
    return dict(CAGR=cagr*100, Vol=r.std(ddof=1)*np.sqrt(ppy)*100, DownDev=dds*100, MaxDD=dd.min()*100,
                Sharpe=(r.mean()*ppy-0.03)/(r.std(ddof=1)*np.sqrt(ppy)),
                Sortino=(r.mean()*ppy-0.03)/dds if dds > 0 else np.nan,
                Calmar=cagr/abs(dd.min()) if dd.min() < 0 else np.nan, Skew=skew(r), pct_down=(r < 0).mean()*100)

p = perf(sat)
print("=== SATELLITE from real deal outcomes (~2 deals/yr, annualized MOIC, invested-weighted) ===")
print(f"  period 2006-2025 | active deals/yr avg {np.mean([len(D[(D.y0<=Y)&(D.y1>Y)]) for Y in sat.index]):.1f}\n")
print(f"  CAGR        {p['CAGR']:6.1f}%")
print(f"  VOLATILITY  {p['Vol']:6.1f}%   <-- vs stock-price satellite 26.8% and monthly NAV 31%")
print(f"  DOWNSIDE VOL{p['DownDev']:6.1f}%   <-- the number you asked about")
print(f"  worst year  {p['MaxDD']:6.1f}%   | down years {p['pct_down']:.0f}% | skew {p['Skew']:+.1f} | Sortino {p['Sortino']:.2f} | Calmar {p['Calmar']:.2f}")
print(f"\n  down years: " + ", ".join(f"{y}:{v*100:.0f}%" for y, v in sat.items() if v < 0) + "  (driven by the loser deals: Actic 0.4x, Usports 0.3x, HQ 0.6x)")

# blend with Core-30 (annual) over the overlap
core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]; core.index = pd.to_datetime(core.index)
corA = (1+core).groupby(core.index.year).prod()-1
df = pd.concat([corA.rename("Core"), sat.rename("Sat")], axis=1).dropna()
print(f"\n=== Core base + deal-outcome Satellite, {df.index.min()}-{df.index.max()} (annual, gross) ===")
for name, w in [("Core-30 (base)", None), ("Satellite (deals)", None), ("Blend 90/10", 0.90), ("Blend 85/15", 0.85), ("Blend 75/25", 0.75)]:
    r = df.Core if name == "Core-30 (base)" else df.Sat if name == "Satellite (deals)" else w*df.Core+(1-w)*df.Sat
    q = perf(r)
    print(f"  {name:20s} CAGR {q['CAGR']:5.1f}%  vol {q['Vol']:5.1f}%  downvol {q['DownDev']:5.1f}%  "
          f"MaxDD {q['MaxDD']:6.1f}%  Sortino {q['Sortino']:.2f}  Calmar {q['Calmar']:.2f}")
print("\nCAVEAT: annualized-MOIC smooths each deal to a constant yearly return, so this is the")
print("FUNDAMENTAL (deal-outcome) volatility -- true mark-to-market vol sits between this and the")
print("stock-price number. But for 'what the activist actually earns', this is the right lens.")
