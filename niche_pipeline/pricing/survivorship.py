#!/usr/bin/env python3
"""survivorship.py — find companies that went out of business (delisted/died) in the sample.

Universe spec (per the manager): listed names in Europe / Americas / Korea / Japan, CompanyMoat
>= 6, market cap TODAY >= USD 100M. That "today" screen is exactly what creates survivorship
bias: a company that went bankrupt in 2009 has no market cap today, so it can never appear in
the screened list. To find the dead ones we must look at the PRICE PANEL, not the screen:

  1. Per instrument, compute first/last trade date, peak close, last close, and drawdown-at-death.
  2. A name whose series ENDS well before the panel's last date (2026) is delisted. Split it:
       - COLLAPSE  (last price <= 30% of peak)  -> went out of business / wiped out
       - EXIT      (last price  > 30% of peak)  -> acquired / merged / taken private (return realised)
  3. Cross-reference against the Scored screen: names in the price panel that are NOT in today's
     screen are the ones the survivor screen silently drops.

  python3 pricing/survivorship.py
"""
import sys, os, glob
import numpy as np, pandas as pd
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import backtest_core30 as B
import openpyxl

SCR = B.SCR
REGIONS = {  # Europe / Americas / Korea / Japan
    "Europe": {"United Kingdom","Germany","France","Switzerland","Sweden","Italy","Norway","Poland",
               "Finland","Greece","Spain","Denmark","Austria","Netherlands","Ireland","Croatia","Portugal",
               "Romania","Hungary","Luxembourg","Lithuania","Slovenia","Cyprus","Estonia","Bulgaria",
               "Czech Republic","Malta","Iceland","Republic of Serbia","Macedonia","Faroe Islands",
               "Isle of Man","Monaco","Russia"},
    "Americas": {"United States of America","Canada","Brazil","Chile","Mexico","Argentina","Peru",
                 "Jamaica","Bermuda","Trinidad and Tobago","Uruguay","Colombia","Panama","Cayman Islands",
                 "Virgin Islands; British"},
    "Korea": {"Korea; Republic (S. Korea)"},
    "Japan": {"Japan"},
}
REGION_OF = {c: reg for reg, cs in REGIONS.items() for c in cs}

# ---- scored screen (today's survivors) ----
ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rows = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rows[0])}
scored = {}
for r in rows[1:]:
    t = r[ci["Ticker"]]
    if not t:
        continue
    scored[str(t)] = dict(company=r[ci["Company"]], country=r[ci["Country"]],
                          moat=r[ci["CompanyMoat(v3.2)"]])
scored_set = set(scored)

# ---- price panel: per-instrument life span + collapse ----
fr = []
for f in B.PRICE_FILES:
    p = pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"])
    p["Instrument"] = p["Instrument"].astype(str)
    fr.append(p)
px = pd.concat(fr)
px["Date"] = pd.to_datetime(px["Date"]); px["c"] = pd.to_numeric(px["Close Price"], errors="coerce")
px = px[px["c"] > 0].dropna(subset=["c"]).drop_duplicates(["Instrument", "Date"])
PANEL_END = px["Date"].max()
DEAD_BEFORE = PANEL_END - pd.Timedelta(days=180)   # last trade > 6 months before panel end = delisted

g = px.sort_values("Date").groupby("Instrument")
life = pd.DataFrame({
    "first": g["Date"].first(), "last": g["Date"].last(),
    "peak": g["c"].max(), "lastpx": g["c"].last(),
})
life["dd_at_end"] = life["lastpx"]/life["peak"] - 1.0
life["delisted"] = life["last"] < DEAD_BEFORE
life["collapse"] = life["delisted"] & (life["lastpx"] <= 0.30*life["peak"])

panel_set = set(life.index)
print(f"panel: {len(panel_set)} instruments, {px['Date'].min():%Y-%m}..{PANEL_END:%Y-%m}")
print(f"scored screen (today's survivors): {len(scored_set)} tickers")
print(f"  in BOTH panel & screen: {len(panel_set & scored_set)}")
print(f"  in panel, NOT in screen: {len(panel_set - scored_set)}  <- candidates the survivor screen drops")
print(f"  in screen, NOT in panel: {len(scored_set - panel_set)}  (no price series available)\n")

# delisted names in the panel, classified
de = life[life["delisted"]].copy()
de["in_screen"] = [i in scored_set for i in de.index]
de["region"] = [REGION_OF.get(scored.get(i, {}).get("country"), "?") for i in de.index]
de["moat"] = [scored.get(i, {}).get("moat") for i in de.index]
de["company"] = [scored.get(i, {}).get("company", "") for i in de.index]
print(f"DELISTED (last trade > 6mo before {PANEL_END:%Y-%m}): {len(de)} instruments")
print(f"  of which COLLAPSE (<=30% of peak, 'out of business'): {int(de['collapse'].sum())}")
print(f"  of which EXIT (acquired/merged, price held up):        {int((~de['collapse']).sum())}\n")

col = de[de["collapse"]].sort_values("last")
print(f"--- {len(col)} COLLAPSE / OUT-OF-BUSINESS names ---")
print(f"{'Instrument':>14} {'last':>10} {'dd%':>6} {'moat':>4} {'reg':>8}  company")
for i, r in col.iterrows():
    m = f"{r['moat']:.1f}" if isinstance(r["moat"], (int, float)) else "  - "
    print(f"{i:>14} {r['last']:%Y-%m-%d} {r['dd_at_end']*100:6.0f} {m:>4} {str(r['region']):>8}  {str(r['company'])[:36]}")

# how many collapses would have been ELIGIBLE (moat>=6, region ok) at some point -> the true bias
elig_dead = col[(col["moat"].apply(lambda x: isinstance(x,(int,float)) and x>=6)) & (col["region"]!="?")]
print(f"\nCollapses that were in-screen with moat>=6 and in-region: {len(elig_dead)}")
print("(These are the names that WOULD bias a backtest if the strategy could have held them and they then died.)")

life.reset_index().to_parquet(SCR+"/survivorship_life.parquet")
print("\nwrote survivorship_life.parquet")
