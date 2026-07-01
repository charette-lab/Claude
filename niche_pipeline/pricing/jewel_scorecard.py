#!/usr/bin/env python3
"""jewel_scorecard.py — rank elite cores as ACTIVIST targets (you create the catalyst).

Not 'what recovers passively' (that's beta) but 'where is the most locked value that
WON'T close without intervention, and is fixable by a board that stops the pivot'.
Snapshot per elite-core name (CoreMoat>=7.8), scored on four data-grounded activist axes:

  LOCKED VALUE   core earnings yield = pre-pivot NOI margin x sales / EV   (cheap on the core)
  MOAT GAP       CoreMoat - CompanyMoat                                    (quality to unlock)
  SELF-HELP      high reinvestment (average_C) into LOW incremental ROIC   (value being destroyed)
  BALANCE SHEET  net cash / market cap = (mktcap - EV)/mktcap              (capital to hand back)
  + gross margin still intact (the core product economics are fine)

Composite = mean percentile across the four axes. No catalyst required; no daily prices.

  python3 pricing/jewel_scorecard.py
"""
import sys, warnings
import numpy as np, pandas as pd, openpyxl
warnings.filterwarnings("ignore")
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
num = lambda v: v if isinstance(v, (int, float)) else None
pct = lambda x: (None if x is None else (x/100.0 if abs(x) > 1.5 else x))

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
c = lambda n: {str(r[ci["Ticker"]]): r[ci[n]] for r in rws[1:] if r[ci["Ticker"]] and n in ci and isinstance(r[ci[n]], (int, float))}
coremoat, companymoat = c("CoreMoat(v3.2)"), c("CompanyMoat(v3.2)")
ctry = {str(r[ci["Ticker"]]): r[ci["Country"]] for r in rws[1:] if r[ci["Ticker"]]}
sect = {str(r[ci["Ticker"]]): r[ci["Industry"]] for r in rws[1:] if r[ci["Ticker"]]}

print("loading panel...", flush=True)
by, idx = panel30.load(XLSB)
def G(rows, i, *names):
    for k in names:
        if k in idx and idx[k] < len(rows[i]) and isinstance(rows[i][idx[k]], (int, float)):
            return rows[i][idx[k]]
    return None

rows = []
for t, rs in by.items():
    t = str(t); cm = coremoat.get(t)
    if cm is None or cm < 7.8: continue
    di = idx.get("Period_End_Date", idx.get("Date"))
    ser = {}
    for i in range(len(rs)):
        dval = rs[i][di] if (di is not None and di < len(rs[i])) else None
        d = pd.to_datetime(dval, errors="coerce")
        noi, sal = num(G(rs, i, "New Operating Income")), num(G(rs, i, "Sales"))
        if pd.isna(d) or noi is None or sal is None or sal <= 0: continue
        ser[d.year] = dict(i=i, m=noi/sal, sal=sal)
    if len(ser) < 6: continue
    yrs = sorted(ser); li = ser[yrs[-1]]["i"]
    mser = pd.Series({y: ser[y]["m"] for y in yrs if abs(ser[y]["m"]) <= 0.60})
    if len(mser) < 5: continue
    pre = mser.rolling(3, min_periods=2).mean().max()
    if pre is None or not (0 < pre <= 0.50): continue
    ev, mc = num(G(rs, li, "EV")), num(G(rs, li, "Market Capitalization", "Market_Capitalization"))
    sal = ser[yrs[-1]]["sal"]
    roic, rr = pct(G(rs, li, "ROICm_total - 7 years")), pct(G(rs, li, "average_C - 7 year"))
    gm, gmn = pct(G(rs, li, "Gross Profit Margin")), pct(G(rs, li, "Gross_Average_Margin_Last_10_yr"))
    if None in (ev, mc, roic, rr) or ev <= 0 or mc <= 0: continue
    rows.append(dict(t=t, coremoat=cm, companymoat=companymoat.get(t, np.nan),
                     moatgap=cm-companymoat.get(t, cm), core_yield=pre*sal/ev,
                     rr=rr, roic=roic, netcash=(mc-ev)/mc, gross_ok=(gm is None or gmn is None or gm >= 0.9*gmn),
                     country=ctry.get(t), sector=sect.get(t)))
D = pd.DataFrame(rows)
# percentile scores on the four activist axes (higher = better target)
D["p_value"] = D.core_yield.rank(pct=True)
D["p_gap"] = D.moatgap.rank(pct=True)
D["p_selfhelp"] = (D.rr.rank(pct=True) + (1-D.roic.rank(pct=True)))/2      # high reinvest + low incremental ROIC
D["p_cash"] = D.netcash.rank(pct=True)
D["JEWEL"] = D[["p_value", "p_gap", "p_selfhelp", "p_cash"]].mean(axis=1)
J = D[D.gross_ok].sort_values("JEWEL", ascending=False)
J.to_parquet(SCR+"/jewel_scorecard.parquet")
print(f"\nelite-core names scored: {len(D)} | gross-intact: {len(J)}\n")
print("TOP 20 ACTIVIST TARGETS (jewel score = locked value + moat gap + self-help + balance-sheet slack):")
print(f"  {'ticker':10s}{'Core':>5}{'Grp':>5}{'gap':>5}{'coreYld':>8}{'RR':>6}{'ROIC':>6}{'netCash':>8}{'JEWEL':>7}")
for _, r in J.head(20).iterrows():
    print(f"  {r.t:10s}{r.coremoat:5.1f}{r.companymoat:5.1f}{r.moatgap:5.1f}{r.core_yield*100:7.0f}%{r.rr*100:5.0f}%{r.roic*100:5.0f}%{r.netcash*100:7.0f}%{r.JEWEL:7.2f}")
top = J.head(60)
print("\nwhere the top-60 targets cluster:")
print("  country:", ", ".join(f"{k}:{v}" for k, v in top.country.value_counts().head(6).items()))
print("  sector :", ", ".join(f"{k[:22]}:{v}" for k, v in top.sector.value_counts().head(6).items()))
print(f"\n  median of the top-60: coreYld {top.core_yield.median()*100:.0f}%  moatgap {top.moatgap.median():.1f}  "
      f"RR {top.rr.median()*100:.0f}%  ROIC {top.roic.median()*100:.0f}%  netCash/mc {top.netcash.median()*100:.0f}%")
print("\nMissing axis = ACTIONABILITY (free float / ownership / control). Add that file and this")
print("becomes a 'cheap AND takeable jewel' list; without it, some tops may be founder/state-locked.")
