#!/usr/bin/env python3
"""core_freed.py — why doesn't doubling NOPAT produce huge returns? (full engine)

The engine scales value linearly with NOPAT, so freeing the core (~1.76x NOPAT) should
lift the implied return a lot -- unless something else caps it. Test, for each wrecked
elite-core name, the real engine er_adj under three scenarios at the CURRENT market cap:
  (1) REPORTED   : reported NOPAT, actual reinvestment (RR) and incremental ROIC
  (2) FREE MARGIN: pre-pivot-margin NOPAT, SAME RR/ROIC (fix only the margin)
  (3) ACTIVIST   : pre-pivot-margin NOPAT + fix capital allocation (RR->40%, ROIC->12%)
Also report the freed earnings yield (freed NOPAT / market cap) to see if they are simply
expensive. Answers: is the value hidden because the market is expensive, or because the
engine (correctly) assumes 80%+ reinvestment at ~0% keeps destroying it?

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/core_freed.py
"""
import sys, warnings
import numpy as np, pandas as pd, openpyxl
warnings.filterwarnings("ignore")
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30, aip, overearning, history
import frameworks as F

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
num = lambda v: v if isinstance(v, (int, float)) else None
pct = lambda x: (None if x is None else (x/100.0 if abs(x) > 1.5 else x))

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
cc = lambda n: {str(r[ci["Ticker"]]): r[ci[n]] for r in rws[1:] if r[ci["Ticker"]] and n in ci and isinstance(r[ci[n]], (int, float))}
coremoat = cc("CoreMoat(v3.2)")
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

def fin(t, nopat, roic, rr, moat, netdebt, mc):
    return {"Company Name": t, "Instrument": t, "GICS Industry Group Name": ind.get(t),
            "Country of Headquarters": ctry.get(t), "New Operating Income": nopat, "ROICm 7": roic,
            "RR 7": rr, "Moat Score": moat, "Gross debt": 0.0, "Income Tax Rate - Instrument": 0.25,
            "Net debt": netdebt, "Sales": None, "Market Cap": mc}

def er(f, sig):
    try:
        ts = overearning.two_stage_return(f, sig, re=0.07, re2=0.12)
        return ts.get("er_adj") if ts else None
    except Exception:
        return None

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
    if pre is None or not (0 < pre <= 0.50) or cur >= 0.80*pre: continue      # wrecked only
    sal, rep_noi = ser[yrs[-1]][2], ser[yrs[-1]][3]
    roic, rr = pct(G(rs, li, "ROICm_total - 7 years")), pct(G(rs, li, "average_C - 7 year"))
    mc, ev = num(G(rs, li, "Market Capitalization", "Market_Capitalization")), num(G(rs, li, "EV"))
    if None in (roic, rr, mc, ev) or mc <= 0: continue
    nd = ev-mc; freed = pre*sal
    try:
        sig = overearning.panel_signals(rs[:li+1], idx)
    except Exception:
        continue
    er_rep = er(fin(t, rep_noi, roic, rr, cm, nd, mc), sig)
    er_free = er(fin(t, freed, roic, rr, cm, nd, mc), sig)
    er_act = er(fin(t, freed, max(roic, 0.12), 0.40, cm, nd, mc), sig)
    rows.append(dict(t=t, cm=cm, rr=rr, roic=roic, uplift=freed/rep_noi if rep_noi > 0 else np.nan,
                     freed_yield=freed/mc, er_rep=er_rep, er_free=er_free, er_act=er_act))
D = pd.DataFrame(rows); D.to_parquet(SCR+"/core_freed.parquet")
m = lambda s: pd.Series(s).dropna().median()
sh = lambda s, k: (pd.Series(s).dropna() >= k).mean()*100
print(f"\nwrecked elite-core names valued (full engine): {len(D)}\n")
print(f"  reinvestment rate (RR) median:       {m(D.rr)*100:4.0f}%")
print(f"  incremental ROIC median:             {m(D.roic)*100:4.0f}%   <- capital plowed back at ~this return")
print(f"  NOPAT uplift if core freed (median):  {m(D.uplift):.2f}x")
print(f"  FREED earnings yield (freedNOPAT/mktcap) median: {m(D.freed_yield)*100:4.0f}%  <- NOT expensive if this is high\n")
print("=== full-engine implied return (er_adj), median, and % clearing hurdles ===")
print(f"  {'scenario':34s}{'median er':>10}{'>=12%':>8}{'>=20%':>8}")
print(f"  {'(1) reported earnings':34s}{m(D.er_rep)*100:9.1f}%{sh(D.er_rep,0.12):7.0f}%{sh(D.er_rep,0.20):7.0f}%")
print(f"  {'(2) free the margin only':34s}{m(D.er_free)*100:9.1f}%{sh(D.er_free,0.12):7.0f}%{sh(D.er_free,0.20):7.0f}%")
print(f"  {'(3) + fix capital allocation':34s}{m(D.er_act)*100:9.1f}%{sh(D.er_act,0.12):7.0f}%{sh(D.er_act,0.20):7.0f}%")
print(f"\n  lift from freeing margin only:      {(m(D.er_free)-m(D.er_rep))*100:+.1f} pts")
print(f"  lift from ALSO fixing reinvestment: {(m(D.er_act)-m(D.er_free))*100:+.1f} pts  <- the real prize")
