#!/usr/bin/env python3
"""core_value.py — value the CORE (subs divested/closed) and check if it's cheap.

For every elite-core name (CoreMoat>=7.8) at its latest snapshot, isolate the core:
  core NOPAT   = normalized (10yr-avg) EBITA margin x sales x (1-tax)   [drag removed]
  core ROIIC   = long-run ROICm (21/14yr)                              [pre-pivot return]
  core RR      = long-run average_C (21/14yr)                          [pre-pivot reinvestment]
  rest (subs)  = divested/closed -> zero going forward (net debt stays with the core)
Then run the engine on CORE inputs vs GROUP (reported) inputs at the current market cap,
and compare the implied returns. Cross-check core growth vs the empirical moat-tier growth.

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/core_value.py
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
EMPIRICAL_G = {7.8: 0.076}      # median NOPAT CAGR for CoreMoat>=7.8 (our full-window table)

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
col = lambda n: {str(r[ci["Ticker"]]): r[ci[n]] for r in rws[1:] if r[ci["Ticker"]] and n in ci and isinstance(r[ci[n]], (int, float))}
coremoat, companymoat, moatgap = col("CoreMoat(v3.2)"), col("CompanyMoat(v3.2)"), col("MoatGap")
ind = {str(r[ci["Ticker"]]): r[ci["Industry"]] for r in rws[1:] if r[ci["Ticker"]]}
ctry = {str(r[ci["Ticker"]]): r[ci["Country"]] for r in rws[1:] if r[ci["Ticker"]]}

print("loading deep panel...", flush=True)
by, idx = panel30.load(XLSB)
def G(r, *names):
    for k in names:
        if k in idx and idx[k] < len(r) and isinstance(r[idx[k]], (int, float)):
            return r[idx[k]]
    return None

def er(nopat, roic, rr, moat, netdebt, mc, t):
    if nopat is None or roic is None or rr is None or not mc or nopat <= 0:
        return None
    fin = {"Company Name": t, "Instrument": t, "GICS Industry Group Name": ind.get(t),
           "Country of Headquarters": ctry.get(t), "New Operating Income": nopat, "ROICm 7": roic,
           "RR 7": max(min(rr, 0.95), 0.0), "Moat Score": moat, "Gross debt": 0.0,
           "Income Tax Rate - Instrument": 0.25, "Net debt": netdebt, "Sales": None, "Market Cap": mc}
    try:
        v = aip.value_and_return(fin, re=0.07, re2=0.12)
        return float(v["er1"]) if v and v.get("er1") is not None else None
    except Exception:
        return None

rows = []
for t, rs in by.items():
    t = str(t); cm = coremoat.get(t)
    if cm is None or cm < 7.8:
        continue
    latest = None
    for r in rs:
        d = pd.to_datetime(G(r, "Period_End_Date") or (r[idx["Date"]] if "Date" in idx else None), errors="coerce")
        if not pd.isna(d) and (latest is None or d > latest[0]):
            latest = (d, r)
    if latest is None:
        continue
    r = latest[1]
    sales, tax = num(G(r, "Sales")), pct(G(r, "Income_Tax_Rate___Instrument", "Income Tax Rate - Instrument"))
    em, em10 = pct(G(r, "EBITA_Margin")), pct(G(r, "EBITA_Average_Margin 10 yr", "EBITA_Average_Margin_Last_10_yr"))
    nopat_g = num(G(r, "New Operating Income"))
    roic_g, rr_g = pct(G(r, "ROICm_total - 7 years", "ROICm_total___7_years")), pct(G(r, "average_C - 7 year", "average_C___7_year"))
    roic_c = pct(G(r, "ROICm_total - 21 years", "ROICm_total___21_years", "ROICm_total - 14 years", "ROICm_total___14_years"))
    rr_c = pct(G(r, "average_C - 21_year", "average_C___21_year", "average_C - 14 year", "average_C___14_year"))
    mc, ev = num(G(r, "Market Capitalization", "Market_Capitalization")), num(G(r, "EV"))
    if None in (sales, em, em10, mc, ev) or tax is None:
        continue
    netdebt = ev - mc
    nopat_c = em10 * sales * (1 - tax)                      # normalized-margin core earnings
    wrecked = em <= 0.80*em10                               # group currently depressed vs own norm
    # two core-growth estimates
    g_termstruct = (roic_c*rr_c) if (roic_c and rr_c) else None
    g_emp = EMPIRICAL_G[7.8]
    er_g = er(nopat_g, roic_g, rr_g, companymoat.get(t), netdebt, mc, t)
    er_c = er(nopat_c, roic_c, rr_c, cm, netdebt, mc, t)
    rows.append(dict(t=t, coremoat=cm, companymoat=companymoat.get(t), moatgap=moatgap.get(t),
                     wrecked=wrecked, roic_g=roic_g, roic_c=roic_c, rr_g=rr_g, rr_c=rr_c,
                     g_ts=g_termstruct, g_emp=g_emp, nopat_g=nopat_g, nopat_c=nopat_c,
                     uplift=(nopat_c/nopat_g if (nopat_g and nopat_g > 0) else np.nan), er_g=er_g, er_c=er_c))

D = pd.DataFrame(rows)
D.to_parquet(SCR+"/core_value.parquet", index=False)
W = D[D.wrecked].copy()
def m(s): s = pd.Series(s).dropna(); return s.median() if len(s) else np.nan
print(f"\nelite-core names valued (CoreMoat>=7.8): {len(D)}  | currently group-wrecked (thesis set): {len(W)}\n")

print("=== your RR/ROIIC hypothesis, tested on the wrecked names ===")
print(f"  core ROIIC (21yr) median   {m(W.roic_c)*100:5.1f}%   vs   group ROIIC (7yr) median {m(W.roic_g)*100:5.1f}%   -> core higher? {m(W.roic_c) > m(W.roic_g)}")
print(f"  core RR   (21yr) median    {m(W.rr_c)*100:5.1f}%   vs   group RR   (7yr) median  {m(W.rr_g)*100:5.1f}%   -> core lower?  {m(W.rr_c) < m(W.rr_g)}")
print(f"  core growth: term-structure {m(W.g_ts)*100:4.1f}%  |  empirical moat-tier {EMPIRICAL_G[7.8]*100:.1f}%   (triangulation)")
print(f"  normalized core NOPAT / reported group NOPAT (earnings uplift): {m(W.uplift):.2f}x")

print("\n=== are the core returns cheap? implied return on CORE vs GROUP ===")
print(f"  GROUP implied return (reported): median {m(W.er_g)*100:5.1f}%   | >=12%: {(W.er_g>=0.12).mean()*100:3.0f}%")
print(f"  CORE  implied return (isolated): median {m(W.er_c)*100:5.1f}%   | >=12%: {(W.er_c>=0.12).mean()*100:3.0f}%")
both = W.dropna(subset=["er_g", "er_c"])
print(f"  core > group for {(both.er_c>both.er_g).mean()*100:3.0f}% of names   | median uplift in implied return: {m(both.er_c-both.er_g)*100:+.1f} pts")
print(f"  CHEAP on the core (er_core>=12% AND core>group): {((both.er_c>=0.12)&(both.er_c>both.er_g)).sum()} of {len(both)} names")
print("\nCAVEATS: today's moat snapshot; period-end market cap (not daily); subs valued at zero")
print("(conservative); normalized-margin core is a modeled reconstruction, not observed segments.")
