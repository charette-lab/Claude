#!/usr/bin/env python3
"""monte_carlo_activist.py — activist bucket risk from the ACTUAL deal-return distribution.

Draw the bucket from the manager's real 39 deals (annualized return = MOIC^(1/years)-1),
2 new deals/year, held 4 years, ~8 concurrent, invested-weight-ish equal. Monte-Carlo many
26-year paths and measure the realized volatility and DOWNSIDE volatility -- the honest
version of the +40%/-8% smooth model. Two regimes: (A) independent deals (pure
diversification), (B) crash-correlated (deep-bear years haircut all live deals ~ the 8%
loss), which is what actually creates book-level down years.

  python3 pricing/monte_carlo_activist.py
"""
import numpy as np
from scipy.stats import skew
rng = np.random.default_rng(7)
# (entry, exit, MOIC) from the real deal tables
DEALS = [(2015,2025,2.0),(2015,2025,1.2),(2016,2025,6.2),(2016,2025,1.7),(2018,2025,1.2),(2018,2025,1.5),(2018,2025,2.8),
         (2019,2025,5.1),(2020,2025,1.2),(2020,2025,5.4),(2024,2025,1.6),(2017,2023,0.4),(2019,2022,1.6),(2015,2020,1.4),
         (2016,2018,1.5),(2015,2017,1.6),(2015,2016,1.7),(2007,2014,2.9),(2011,2014,3.9),(2007,2014,4.4),(2009,2014,5.5),
         (2012,2014,1.6),(2010,2014,5.1),(2011,2014,1.3),(2010,2014,1.5),(2010,2014,2.0),(2007,2014,28.2),(2007,2014,5.5),
         (2011,2013,0.3),(2006,2007,3.1),(2006,2007,1.4),(2006,2014,8.4),(2006,2011,2.3),(2006,2007,3.0),(2007,2011,2.1),
         (2007,2011,1.4),(2007,2011,5.4),(2003,2006,5.3),(2007,2010,0.6)]
ar = np.array([m**(1/max(b-a,1))-1 for a,b,m in DEALS])          # annualized deal returns
print(f"deal annualized-return distribution ({len(ar)} deals):")
print(f"  mean {ar.mean()*100:.0f}%  median {np.median(ar)*100:.0f}%  std {ar.std()*100:.0f}%  "
      f"min {ar.min()*100:.0f}%  max {ar.max()*100:.0f}%  | losers (<0): {(ar<0).mean()*100:.0f}%")
ar_exbest = np.sort(ar)[:-1]                                     # ex the single best (Klarna 28x)
print(f"  ex-best: mean {ar_exbest.mean()*100:.0f}%  median {np.median(ar_exbest)*100:.0f}%")

NSIM, YEARS, NPICK, HOLD = 4000, 26, 2, 4
CRASH_YEARS = {3, 9}                                            # ~year indices for the 2 deep bears (~1/decade)
def simulate(correlated, pool):
    vols, dvols, cagrs, minyrs = [], [], [], []
    for _ in range(NSIM):
        picks = {y: rng.choice(pool, NPICK) for y in range(YEARS)}
        book = []
        for z in range(HOLD, YEARS):
            active = np.concatenate([picks[v] for v in range(z-HOLD+1, z+1)])
            r = active.mean()
            if correlated and z in CRASH_YEARS:
                r = min(r, 0.0) - 0.08                          # deep-bear: all live deals haircut -> ~ -8% book yr
            book.append(r)
        b = np.array(book); eq = np.cumprod(1+b)
        vols.append(b.std(ddof=1)); dvols.append(np.sqrt((np.minimum(b,0)**2).mean()))
        cagrs.append(eq[-1]**(1/len(b))-1); minyrs.append(b.min())
    q = lambda a,p: np.percentile(a,p)
    return dict(vol=np.median(vols)*100, dvol=np.median(dvols)*100, cagr=np.median(cagrs)*100,
                dvol5=q(dvols,5)*100, dvol95=q(dvols,95)*100, minyr=np.median(minyrs)*100)

for lab, corr, pool in [("A) independent deals (pure diversification)", False, ar),
                        ("B) crash-correlated deep-bear years", True, ar),
                        ("B') crash-correlated, EX-BEST deal pool", True, ar_exbest)]:
    s = simulate(corr, pool)
    print(f"\n{lab}")
    print(f"   bucket CAGR {s['cagr']:5.1f}%  |  VOLATILITY {s['vol']:5.1f}%  |  DOWNSIDE VOL {s['dvol']:4.1f}% "
          f"(5-95%: {s['dvol5']:.1f}-{s['dvol95']:.1f})  |  worst yr {s['minyr']:.0f}%")
print("\nread: 'downside vol' is the annual downside deviation of the ~8-deal book. Pure diversification (A)")
print("makes it small; realistic crash-correlation (B) -- where a deep bear marks every live deal at once --")
print("is what lifts it toward your 7-8% intuition. Bucket CAGR reflects your real deal spread incl. big winners.")
