#!/usr/bin/env python3
"""activist_funnel.py — how many names/year beat the Core Index on the ACTIVIST basis?

Same funnel as count_opportunities, but each name is valued as if the core is freed:
core earnings = pre-pivot NOI margin x sales, ROIIC set to the freed core's level (15% base,
12/20% sensitivity), reinvestment KEPT (redirected into the core). Uses the daily
decomposition for each name's representative market cap and reported er per year, then the
real engine er1 for the freed value. Counts elite-core (CoreMoat>=7.8) names whose FREED er
beats the Core Index's own (model) expected-return hurdle, and isolates the PURE activist
set (reported er below the hurdle, freed er above it -- reachable only via the intervention).

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/activist_funnel.py
"""
import sys, os, json
import numpy as np, pandas as pd, openpyxl
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30, aip
from backtest_core30 import load_meta, SCR
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
num = lambda v: v if isinstance(v, (int, float)) else None
pct = lambda x: (None if x is None else (x/100.0 if abs(x) > 1.5 else x))
import json  # HURDLE (Core Index expected return per year) computed from actual Core-30 holdings below
ROIC = {"12%": 0.12, "15%": 0.15, "20%": 0.20}

tags, maxsev, moat, country = load_meta()
ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
gc = lambda n: {str(r[ci["Ticker"]]): r[ci[n]] for r in rws[1:] if r[ci["Ticker"]] and n in ci and isinstance(r[ci[n]], (int, float))}
coremoat = gc("CoreMoat(v3.2)")
ownv = {str(r[ci["Ticker"]]): r[ci["OwnerVerdict"]] for r in rws[1:] if r[ci["Ticker"]]}
ind = {str(r[ci["Ticker"]]): r[ci["Industry"]] for r in rws[1:] if r[ci["Ticker"]]}
ctry = {str(r[ci["Ticker"]]): r[ci["Country"]] for r in rws[1:] if r[ci["Ticker"]]}

print("loading decomposition (reported er + representative market cap per year)...", flush=True)
dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet", columns=["Instrument", "Date", "er_total", "market_cap", "artifact"])
dec["Instrument"] = dec["Instrument"].astype(str); dec["year"] = pd.to_datetime(dec["Date"]).dt.year
dec = dec[(dec.year >= 2000) & (dec.year <= 2026)]
agg = dec.groupby(["Instrument", "year"]).agg(rep_er=("er_total", "median"), rep_mc=("market_cap", "median"),
                                              art=("artifact", "mean")).reset_index()
agg = agg[(agg.art < 0.5) & (agg.rep_mc > 0)]
# Core Index hurdle per year = median reported er of the actual Core-30 holdings that year
ermap = {(int(r.year), r.Instrument): r.rep_er for r in agg.itertuples()}
held = {}
for dstr, tks in json.load(open(SCR+"/bt_core30_holds_Q.json")).items():
    held.setdefault(pd.Timestamp(dstr).year, set()).update(str(x) for x in tks)
HURDLE = {}
for y, names in held.items():
    e = [ermap[(y, n)] for n in names if (y, n) in ermap]
    if e: HURDLE[y] = float(np.median(e))
HDEF = float(np.median(list(HURDLE.values()))) if HURDLE else 0.35
print(f"  hurdle by year (median): {min(HURDLE.values())*100:.0f}%-{max(HURDLE.values())*100:.0f}% | default {HDEF*100:.0f}%", flush=True)

print("loading panel (pre-pivot margin + fundamentals per year)...", flush=True)
by, idx = panel30.load(XLSB)
di = idx.get("Period_End_Date", idx.get("Date"))
def G(rows, i, *names):
    for k in names:
        if k in idx and idx[k] < len(rows[i]) and isinstance(rows[i][idx[k]], (int, float)):
            return rows[i][idx[k]]
    return None
FUND = {}      # (t, year) -> dict(pre, sales, rr, netdebt)
for t, rs in by.items():
    t = str(t)
    if coremoat.get(t, 0) < 7.8: continue
    marg = []
    for i in range(len(rs)):
        dval = rs[i][di] if (di is not None and di < len(rs[i])) else None
        d = pd.to_datetime(dval, errors="coerce"); noi, sal = num(G(rs, i, "New Operating Income")), num(G(rs, i, "Sales"))
        if pd.isna(d) or noi is None or sal is None or sal <= 0: continue
        marg.append((d.year, noi/sal))
        m = pd.Series({y: v for y, v in marg if abs(v) <= 0.60})
        if len(m) < 5: continue
        pre = m.rolling(3, min_periods=2).mean().max()
        if pre is None or not (0 < pre <= 0.50): continue
        rr = pct(G(rs, i, "average_C - 7 year")); ev, mc = num(G(rs, i, "EV")), num(G(rs, i, "Market Capitalization", "Market_Capitalization"))
        if rr is None or ev is None or mc is None: continue
        FUND[(t, d.year)] = dict(pre=pre, sales=sal, rr=rr, netdebt=ev-mc)

def er1(t, nopat, roic, rr, nd, mc):
    f = {"Company Name": t, "Instrument": t, "GICS Industry Group Name": ind.get(t), "Country of Headquarters": ctry.get(t),
         "New Operating Income": nopat, "ROICm 7": roic, "RR 7": rr, "Moat Score": coremoat.get(t), "Gross debt": 0.0,
         "Income Tax Rate - Instrument": 0.25, "Net debt": nd, "Sales": None, "Market Cap": mc}
    try:
        v = aip.value_and_return(f, re=0.07, re2=0.12); return v.get("er1") if v else None
    except Exception:
        return None

# for each (name, year): reported er, freed er at each ROIIC, at the representative market cap (trade-year Y uses FY Y-1)
rec = []
for _, r in agg.iterrows():
    t, y = r.Instrument, int(r.year)
    key = (t, y-1) if (t, y-1) in FUND else ((t, y) if (t, y) in FUND else None)   # 6-mo reporting lag
    if key is None: continue
    f = FUND[key]; freed = f["pre"]*f["sales"]
    row = dict(t=t, year=y, hurdle=HURDLE.get(y, HDEF), rep_er=float(r.rep_er), takeable=(ownv.get(t) != "HARD-BLOCK"))
    for lab, rv in ROIC.items():
        row[f"freed_{lab}"] = er1(t, freed, rv, f["rr"], f["netdebt"], float(r.rep_mc))
    rec.append(row)
D = pd.DataFrame(rec); D.to_parquet(SCR+"/activist_funnel.parquet")

def cnt(mask_fn):
    return {y: int(mask_fn(g).sum()) for y, g in D.groupby("year")}
yrs = list(range(2000, 2026)); mean = lambda d: np.mean([d.get(y, 0) for y in yrs])
passive = cnt(lambda g: g.rep_er > g.hurdle)
act15 = cnt(lambda g: g["freed_15%"] > g.hurdle)
pure15 = cnt(lambda g: (g.rep_er <= g.hurdle) & (g["freed_15%"] > g.hurdle))
pure15_tk = cnt(lambda g: (g.rep_er <= g.hurdle) & (g["freed_15%"] > g.hurdle) & g.takeable)
act12 = cnt(lambda g: g["freed_12%"] > g.hurdle); act20 = cnt(lambda g: g["freed_20%"] > g.hurdle)

print(f"\nelite-core name-years evaluated: {len(D)} | hurdle = Core Index model er (~35%/yr)\n")
print(f"{'year':6}{'passive':>9}{'activist@15%':>13}{'PURE act@15%':>13}{'takeable':>10}{'act@12%':>9}{'act@20%':>9}")
for y in yrs:
    print(f"{y:<6}{passive.get(y,0):>9}{act15.get(y,0):>13}{pure15.get(y,0):>13}{pure15_tk.get(y,0):>10}{act12.get(y,0):>9}{act20.get(y,0):>9}")
print(f"\n  per-year AVG:  passive {mean(passive):.0f} | activist@15% {mean(act15):.0f} | "
      f"PURE activist@15% {mean(pure15):.0f} | takeable {mean(pure15_tk):.0f} | @12% {mean(act12):.0f} | @20% {mean(act20):.0f}")
print(f"\n  PURE activist = beats the index ONLY after the freed-core fix (reachable via intervention).")
print(f"  Both sides are model er (same ~2x optimism), so it's apples-to-apples. ROIIC is the swing:")
print(f"  ~{mean(act12):.0f}/yr at a conservative 12% core ROIIC, ~{mean(act20):.0f}/yr at 20%.")
