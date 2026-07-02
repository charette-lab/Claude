#!/usr/bin/env python3
"""value_bridge.py — per-name activist value bridge on the takeable-jewel list.

For each ranked jewel: status-quo engine er1 (reported) -> freed-core er1 across a grid of
core ROIIC assumptions (the DD sensitivity axis), with the reinvestment rate KEPT (redirect
the capital into the core, don't cut it). The margin uplift shows how much comes from
restoring the margin; the ROIIC columns show how much comes from redirecting the capital.
Analysts pick the ROIIC column matching their DD estimate and read the underwritten return.

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/value_bridge.py
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
GRID = [0.10, 0.12, 0.15, 0.20]

JW = pd.read_parquet(SCR+"/jewel_scorecard.parquet").set_index("t")     # ranked takeable jewels + ownership
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
for t in JW.index:
    t = str(t); cm = coremoat.get(t)
    if t not in by or cm is None: continue
    rs = by[t]; ser = {}
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
    if pre is None or not (0 < pre <= 0.50): continue
    sal, rep_noi = ser[yrs[-1]][2], ser[yrs[-1]][3]
    roic, rr = pct(G(rs, li, "ROICm_total - 7 years")), pct(G(rs, li, "average_C - 7 year"))
    mc, ev = num(G(rs, li, "Market Capitalization", "Market_Capitalization")), num(G(rs, li, "EV"))
    if None in (roic, rr, mc, ev) or mc <= 0: continue
    nd = ev-mc; freed = pre*sal
    rep_er, wacc = er1(fin(t, rep_noi, roic, rr, cm, nd, mc))
    grid_er = {int(g*100): er1(fin(t, freed, g, rr, cm, nd, mc))[0] for g in GRID}
    j = JW.loc[t]
    rows.append(dict(Ticker=t, Country=ctry.get(t), Sector=(str(ind.get(t))[:22] if ind.get(t) else None),
                     CoreMoat=cm, GroupMoat=round(float(j.companymoat), 1) if pd.notna(j.companymoat) else None,
                     OwnerVerdict=j.verdict, OwnerBloc=(round(float(j.bloc)) if pd.notna(j.bloc) else None),
                     RepMargin=round(cur*100), PreMargin=round(pre*100), MarginUplift=round(pre/cur, 2) if cur > 0 else None,
                     ReinvestRate=round(rr*100), WACC=round(wacc*100, 1) if wacc else None,
                     RepER1=round(rep_er*100, 1) if rep_er is not None else None,
                     **{f"FreedER1@{g}%": (round(grid_er[g]*100, 1) if grid_er[g] is not None else None) for g in [10, 12, 15, 20]},
                     Jewel=round(float(j.JEWEL), 3), OwnerNotes=str(j.notes)[:120]))
B = pd.DataFrame(rows).sort_values("Jewel", ascending=False)
B.to_excel(SCR+"/value_bridge.xlsx", index=False); B.to_parquet(SCR+"/value_bridge.parquet")
med = lambda c: pd.to_numeric(B[c], errors="coerce").median()
print(f"\nvalue bridge built for {len(B)} takeable jewels (RR kept = redirect capital, not cut)\n")
print(f"median: MarginUplift {med('MarginUplift'):.2f}x  ReinvestRate {med('ReinvestRate'):.0f}%  RepER1 {med('RepER1'):.1f}%")
print(f"        FreedER1  @10%={med('FreedER1@10%'):.0f}%  @12%={med('FreedER1@12%'):.0f}%  @15%={med('FreedER1@15%'):.0f}%  @20%={med('FreedER1@20%'):.0f}%\n")
print("TOP 18 (status-quo er1 -> freed-core er1 by assumed core ROIIC):")
cols = ["Ticker", "Country", "CoreMoat", "OwnerVerdict", "OwnerBloc", "RepMargin", "PreMargin", "ReinvestRate", "RepER1", "FreedER1@12%", "FreedER1@15%", "FreedER1@20%"]
print(B[cols].head(18).to_string(index=False))
print("\nwrote value_bridge.xlsx (full ranked list; pick the ROIIC column matching your DD estimate).")
