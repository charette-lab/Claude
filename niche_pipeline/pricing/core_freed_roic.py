#!/usr/bin/env python3
"""core_freed_roic.py — corrected freed-core valuation: ROIIC must move to the CORE's level.

Prior error: I freed the margin (pre-pivot NOI margin x sales) but kept the consolidated
ROIIC (~2%), which is the subsidiary drag itself -- so the retained capital kept 'destroying'
value in the model. Once the core is freed, reinvestment goes into the moat-7.8 core at a
return ABOVE the cost of capital, so it ADDS value. This re-runs the freed core across a grid
of core ROIIC levels (keeping the actual reinvestment rate), using the real engine er1, to
show how the implied return responds when ROIIC reflects the new reality.

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/core_freed_roic.py
"""
import sys, warnings
import numpy as np, pandas as pd, openpyxl
warnings.filterwarnings("ignore")
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30, aip
SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
num = lambda v: v if isinstance(v, (int, float)) else None
pct = lambda x: (None if x is None else (x/100.0 if abs(x) > 1.5 else x))
GRID = [0.07, 0.10, 0.12, 0.15, 0.20, 0.25]

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
coremoat = {str(r[ci["Ticker"]]): r[ci["CoreMoat(v3.2)"]] for r in rws[1:]
            if r[ci["Ticker"]] and isinstance(r[ci["CoreMoat(v3.2)"]], (int, float))}
ind = {str(r[ci["Ticker"]]): r[ci["Industry"]] for r in rws[1:] if r[ci["Ticker"]]}
ctry = {str(r[ci["Ticker"]]): r[ci["Country"]] for r in rws[1:] if r[ci["Ticker"]]}

print("loading panel...", flush=True)
by, idx = panel30.load(XLSB)
di = idx.get("Period_End_Date", idx.get("Date"))
def G(rows, i, *names):
    for k in names:
        if k in idx and idx[k] < len(rows[i]) and isinstance(rows[i][idx[k]], (int, float)):
            return rows[i][idx[k]]
    return None
def fin(t, nopat, roic, rr, moat, nd, mc):
    return {"Company Name": t, "Instrument": t, "GICS Industry Group Name": ind.get(t),
            "Country of Headquarters": ctry.get(t), "New Operating Income": nopat, "ROICm 7": roic,
            "RR 7": rr, "Moat Score": moat, "Gross debt": 0.0, "Income Tax Rate - Instrument": 0.25,
            "Net debt": nd, "Sales": None, "Market Cap": mc}
def er1(f):
    try:
        v = aip.value_and_return(f, re=0.07, re2=0.12); return (v.get("er1"), v.get("wacc")) if v else (None, None)
    except Exception:
        return None, None

rows = []
for t, rs in by.items():
    t = str(t); cm = coremoat.get(t)
    if cm is None or cm < 7.8: continue
    ser = {}
    for i in range(len(rs)):
        dval = rs[i][di] if (di is not None and di < len(rs[i])) else None
        d = pd.to_datetime(dval, errors="coerce"); noi, sal = num(G(rs, i, "New Operating Income")), num(G(rs, i, "Sales"))
        if pd.isna(d) or noi is None or sal is None or sal <= 0: continue
        ser[d.year] = (i, noi/sal, sal, noi)
    if len(ser) < 6: continue
    yrs = sorted(ser); li = ser[yrs[-1]][0]
    mser = pd.Series({y: ser[y][1] for y in yrs if abs(ser[y][1]) <= 0.60})
    if len(mser) < 5: continue
    pre = mser.rolling(3, min_periods=2).mean().max()
    cur = pd.Series({y: ser[y][1] for y in yrs}).rolling(3, min_periods=1).mean().iloc[-1]
    if pre is None or not (0 < pre <= 0.50) or cur >= 0.80*pre: continue
    sal, rep_noi = ser[yrs[-1]][2], ser[yrs[-1]][3]
    roic, rr = pct(G(rs, li, "ROICm_total - 7 years")), pct(G(rs, li, "average_C - 7 year"))
    mc, ev = num(G(rs, li, "Market Capitalization", "Market_Capitalization")), num(G(rs, li, "EV"))
    if None in (roic, rr, mc, ev) or mc <= 0: continue
    nd = ev-mc; freed = pre*sal
    row = dict(t=t, rr=rr, roic_rep=roic)
    row["rep"], row["wacc"] = er1(fin(t, rep_noi, roic, rr, cm, nd, mc))         # reported
    row["free_oldROIC"], _ = er1(fin(t, freed, roic, rr, cm, nd, mc))            # my WRONG scenario
    for gx in GRID:
        row[f"free_{int(gx*100)}"], _ = er1(fin(t, freed, gx, rr, cm, nd, mc))   # corrected: core ROIIC
    rows.append(row)
D = pd.DataFrame(rows); D.to_parquet(SCR+"/core_freed_roic.parquet")
m = lambda s: pd.Series(s).dropna().median()
sh = lambda s, k: (pd.Series(s).dropna() >= k).mean()*100
print(f"\nwrecked elite-core names: {len(D)} | median WACC {m(D.wacc)*100:.1f}% | median reinvest rate {m(D.rr)*100:.0f}%\n")
print(f"  reported (as-is):                       median er1 {m(D.rep)*100:5.1f}%   >=12%: {sh(D.rep,0.12):3.0f}%")
print(f"  free margin, KEEP old ROIIC ({m(D.roic_rep)*100:.0f}%) [WRONG]: median er1 {m(D.free_oldROIC)*100:5.1f}%   >=12%: {sh(D.free_oldROIC,0.12):3.0f}%")
print("\n=== CORRECTED: free the core, ROIIC set to the core's level (reinvest rate kept) ===")
print(f"  {'core ROIIC':>12}{'median er1':>12}{'>=12%':>8}{'>=20%':>8}")
for gx in GRID:
    col = D[f"free_{int(gx*100)}"]
    tag = " (=WACC)" if abs(gx-m(D.wacc)) < 0.01 else ""
    print(f"  {gx*100:10.0f}%{m(col)*100:11.1f}%{sh(col,0.12):7.0f}%{sh(col,0.20):7.0f}%{tag}")
print(f"\n  reading: once ROIIC > WACC (~{m(D.wacc)*100:.0f}%), reinvestment ADDS value -> the return climbs with ROIIC,")
print("  exactly as you said. The consolidated ROICm can't supply the core ROIIC (it's the drag);")
print("  it comes from the moat / DD estimate of the freed core's returns on capital.")
