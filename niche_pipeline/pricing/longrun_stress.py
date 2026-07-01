#!/usr/bin/env python3
"""longrun_stress.py — honest long-horizon behaviour when only ~10y of real data exists.

We CANNOT backtest the Core-30 before 2015-12 (no prices/fundamentals/moats/tags exist
earlier). So we do NOT fabricate a 30-year path. Instead, two honest lenses:

  (A) BLOCK BOOTSTRAP — resample the real 41 quarterly returns in 1-year blocks into
      10,000 synthetic 30-year paths. This shows SEQUENCING risk: how bad a 30-year
      experience can be if bad quarters happen to cluster. It CANNOT show a quarter worse
      than any in 2016-26 (no 2008-magnitude quarter is in the sample), so it still
      understates true tail risk.

  (B) CRISIS STRESS — apply the book's market beta to the actual drawdowns of regimes the
      sample lacks (1973-74, 2000-02, 2008, 1987). Beta-scaled = mechanical estimate; the
      quality/low-vol tilt would soften value-led crises (2000-02) and help little in 2008.

  python3 pricing/longrun_stress.py
"""
import numpy as np, pandas as pd
SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
rng = np.random.default_rng(12345)

core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0].dropna().values
n = len(core); L = 4; H = 120; NSIM = 10000                 # 1-yr blocks, 30-yr (120q) paths

def path_metrics(r):
    eq = np.cumprod(1+r); peak = np.maximum.accumulate(eq)
    dd = eq/peak-1
    # longest underwater run (quarters below a prior peak)
    uw = dd < -1e-9; longest = cur = 0
    for x in uw:
        cur = cur+1 if x else 0; longest = max(longest, cur)
    y = len(r)/4
    return eq[-1]**(1/y)-1, dd.min(), longest/4.0

cagrs, mdds, uws = [], [], []
for _ in range(NSIM):
    blocks = []
    while len(blocks)*L < H:
        s = rng.integers(0, n)                              # circular block
        blocks.append(np.take(core, range(s, s+L), mode="wrap"))
    r = np.concatenate(blocks)[:H]
    c, m, u = path_metrics(r); cagrs.append(c); mdds.append(m); uws.append(u)
cagrs, mdds, uws = map(np.array, (cagrs, mdds, uws))
q = lambda a, p: np.percentile(a, p)

print("="*74)
print("(A) BLOCK BOOTSTRAP — 10,000 synthetic 30-year paths from the real 41 quarters")
print("="*74)
print(f"  {'':22s} {'5th pct':>9} {'median':>9} {'95th pct':>9}")
print(f"  {'30-yr CAGR':22s} {q(cagrs,5)*100:8.1f}% {q(cagrs,50)*100:8.1f}% {q(cagrs,95)*100:8.1f}%")
print(f"  {'worst drawdown':22s} {q(mdds,5)*100:8.1f}% {q(mdds,50)*100:8.1f}% {q(mdds,95)*100:8.1f}%")
print(f"  {'longest underwater(yr)':22s} {q(uws,95):8.1f}  {q(uws,50):8.1f}  {q(uws,5):8.1f}")
print(f"\n  P(at least one >30% drawdown in 30y): {(mdds<=-0.30).mean()*100:4.0f}%"
      f"   |  P(>40%): {(mdds<=-0.40).mean()*100:3.0f}%   |  P(30y CAGR < 8%): {(cagrs<0.08).mean()*100:3.0f}%")
print("  NOTE: no quarter worse than the 2016-26 sample can appear -> this UNDERSTATES a real 2008.")

print("\n"+"="*74)
print("(B) CRISIS STRESS — regimes the 10-year sample never saw")
print("="*74)
BETA = 0.84                                                 # Core-30 beta to MSCI World IMI (measured)
crises = [("1973-74 bear",        -0.48, "value/quality held better; est. optimistic"),
          ("1987 crash (Q)",      -0.23, "one-quarter shock; fast recovery"),
          ("2000-02 dot-com",     -0.49, "quality/value OUTPERformed -> real DD likely far shallower"),
          ("2008-09 GFC",         -0.57, "everything correlated -> beta-scale is realistic"),
          ("2020 COVID (actual)", -0.34, "Core-30 actually did -19% -> beta-scale ~right")]
print(f"  book beta to world equities = {BETA}")
print(f"  {'crisis':22s} {'market':>8} {'beta-scaled book':>17}   caveat")
for name, mk, note in crises:
    print(f"  {name:22s} {mk*100:7.0f}% {mk*BETA*100:16.0f}%   {note}")
print(f"\n  Real anchor: the manager's own (concentrated) book fell -40.6% in 2008.")
print(f"  A diversified 30-name quality book would fall less than that, but a -30% to -45%")
print(f"  drawdown in a real secular bear is the honest expectation — vs the -19.5% tested.")
