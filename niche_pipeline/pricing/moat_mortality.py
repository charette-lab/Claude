#!/usr/bin/env python3
"""moat_mortality.py — estimated death (bankruptcy/liquidation) hazard by moat band.

We CANNOT measure this from the panel: it is survivor-only (see survivorship.py -- the price
data and the screen are the same 3351 names, so companies that died before today were never
pulled). So this is a calibrated JUDGEMENT model, anchored to the empirical small-cap
delisting-FOR-CAUSE literature (CRSP delist codes 400s; Datastream dead lists), mapped onto the
AIP v3.2 moat scale where 6.5 is the quality floor, 7.8+ is wide, and <5 is a weak/narrow moat.

Death here = went out of business (liquidation/bankruptcy/wipeout), NOT acquired-exit.
Annual hazard rises steeply as the moat weakens and is scaled UP for small size (near the
USD 100M floor a weak franchise has thin liquidity, refinancing and single-customer risk).

  python3 pricing/moat_mortality.py
"""
import numpy as np

# central annual death hazard by moat band (large/mid). Judgement, anchored to empirical
# small-cap bankruptcy rates (~1-2%/yr broad market; wide-moat compounders far lower).
BANDS = [
    ("8.0+  wide",        0.002),
    ("7.0   strong",      0.005),
    ("6.5   floor",       0.008),
    ("6.0   narrow",      0.012),
    ("5.0   thin",        0.025),
    ("4.5   weak",        0.040),   # <- the question
    ("4.0   very weak",   0.055),
    ("3.0   no moat",     0.090),
]
# small-cap multiplier (near the $100M floor): weak franchise + small size compounds risk
SMALL_MULT = {"lo": 1.0, "central": 1.15, "hi": 1.4}
HORIZONS = [7, 14, 31]

def cum(h, n):
    return 1 - (1-h)**n

print("Estimated OUT-OF-BUSINESS probability by moat band (central hazard, mid/large size):")
print(f"{'moat band':>16} | {'ann.':>5} | " + " | ".join(f"{n}yr" for n in HORIZONS))
print("-"*46)
for lbl, h in BANDS:
    print(f"{lbl:>16} | {h*100:4.1f}% | " + " | ".join(f"{cum(h,n)*100:4.0f}%" for n in HORIZONS))

print("\n--- FOCUS: small company (near $100M) at MOAT 4.5 ---")
h0 = dict(BANDS)["4.5   weak"]
print(f"{'scenario':>10} | {'ann.hazard':>10} | " + " | ".join(f"{n}yr death':>0" and f'{n}yr' for n in HORIZONS))
print("-"*46)
for k, m in SMALL_MULT.items():
    h = h0*m
    lab = {"lo":"low","central":"CENTRAL","hi":"high"}[k]
    print(f"{lab:>10} | {h*100:8.1f}%  | " + " | ".join(f"{cum(h,n)*100:4.0f}%" for n in HORIZONS))

hc = h0*SMALL_MULT["central"]
print(f"\nGUESS: a small (~$100M) company at moat 4.5 has ~{hc*100:.0f}%/yr chance of going out of business")
print(f"       -> ~{cum(hc,7)*100:.0f}% over one 7-yr moat cycle, ~{cum(hc,14)*100:.0f}% over two, ~{cum(hc,31)*100:.0f}% over the full 1995-2026 sample.")
print("\nWhy this matters: your moats are sticky (<10%/7yr), so a name seen at 4.5 is a genuine")
print("tail -- >10% below the 6.0 floor, i.e. structurally weak, not a healthy name having a bad year.")
print("That 4.5 bucket is EXACTLY the population absent from the survivor-only panel, so the true")
print("backtest drawdowns/CAGR are flattered by roughly the death rate of the names you'd have")
print("touched at the weak end. The Core-30 (moat>=6.5 at selection) rarely holds a 4.5, so its")
print("direct exposure is small; the broader eligible/activist funnel is where the bias bites.")
