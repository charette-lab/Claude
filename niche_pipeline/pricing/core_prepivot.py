#!/usr/bin/env python3
"""core_prepivot.py — value the core at its PRE-PIVOT New Operating Income margin.

No segment data, no ROICm. Use the company's own history: the NOI margin it earned
before management pivoted into the value-destroying subsidiaries is the cleanest read
on core earnings power. Method per elite-core name (CoreMoat>=7.8):

  NOI margin series  = New Operating Income / Sales, per year
  pre-pivot margin   = peak sustained (3yr-avg) NOI margin in its history  [the clean core]
  current margin     = latest 3yr-avg NOI margin
  normalized core NOPAT = pre-pivot margin x CURRENT sales                 [core earnings today]
  grow at            = empirical moat-tier NOPAT growth (~7.6% for CoreMoat>=7.8)
  implied core return = normalized core NOPAT / EV  +  moat-tier growth    [Gordon/Bogle]

Compare to the same on the CURRENT (post-pivot) margin to see the uplift, and judge cheap
vs the 12% hurdle. 'Wrecked' = current margin well below the pre-pivot peak.

  python3 pricing/core_prepivot.py
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
G_BY_MOAT = {9.0: 0.084, 7.8: 0.076}      # empirical full-window NOPAT CAGR by moat tier

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
col = lambda n: {str(r[ci["Ticker"]]): r[ci[n]] for r in rws[1:] if r[ci["Ticker"]] and n in ci and isinstance(r[ci[n]], (int, float))}
coremoat, companymoat = col("CoreMoat(v3.2)"), col("CompanyMoat(v3.2)")

print("loading deep panel...", flush=True)
by, idx = panel30.load(XLSB)
def GET(r, *names):
    for k in names:
        if k in idx and idx[k] < len(r) and isinstance(r[idx[k]], (int, float)):
            return r[idx[k]]
    return None

rows = []
for t, rs in by.items():
    t = str(t); cm = coremoat.get(t)
    if cm is None or cm < 7.8:
        continue
    ser = {}
    for r in rs:
        d = pd.to_datetime(GET(r, "Period_End_Date") or (r[idx["Date"]] if "Date" in idx else None), errors="coerce")
        noi, sal = num(GET(r, "New Operating Income")), num(GET(r, "Sales"))
        ev, mc = num(GET(r, "EV")), num(GET(r, "Market Capitalization", "Market_Capitalization"))
        if pd.isna(d) or noi is None or sal is None or sal <= 0:
            continue
        ser[d.year] = dict(m=noi/sal, sal=sal, ev=ev, mc=mc, y=d.year)
    if len(ser) < 6:
        continue
    yrs = sorted(ser)
    sals = pd.Series({y: ser[y]["sal"] for y in yrs})
    # peak margin only from years with a PLAUSIBLE NOI margin (drop |margin|>60% tiny-sales artifacts)
    marg = pd.Series({y: ser[y]["m"] for y in yrs if abs(ser[y]["m"]) <= 0.60})
    if len(marg) < 5:
        continue
    roll3 = marg.rolling(3, min_periods=2).mean()
    pre = roll3.max()                                   # peak sustained (pre-pivot) core margin
    pre_yr = int(roll3.idxmax())
    cur = pd.Series({y: ser[y]["m"] for y in yrs}).rolling(3, min_periods=1).mean().loc[yrs[-1]]
    last = ser[yrs[-1]]
    ev, mc, sal = last["ev"], last["mc"], last["sal"]
    if pre is None or not (0 < pre <= 0.50) or ev is None or ev <= 0 or mc is None or mc <= 0 or sal <= 0:
        continue
    # the pre-pivot peak must come from a period with real sales (>=25% of current) — kill tiny-denominator flukes
    peak_sales = sals.loc[max(pre_yr-2, yrs[0]):pre_yr].mean()
    if pd.isna(peak_sales) or peak_sales < 0.25*sal:
        continue
    g = G_BY_MOAT[9.0] if cm >= 9.0 else G_BY_MOAT[7.8]
    nopat_pre = pre * sal                               # core earnings at pre-pivot margin, current sales
    nopat_now = cur * sal
    er_core = nopat_pre/ev + g                          # Gordon: normalized yield + growth
    er_now = nopat_now/ev + g
    rows.append(dict(t=t, coremoat=cm, companymoat=companymoat.get(t), pre_margin=pre, cur_margin=cur,
                     pre_yr=pre_yr, margin_destroyed=1-cur/pre if pre > 0 else np.nan,
                     nopat_uplift=nopat_pre/nopat_now if nopat_now > 0 else np.nan,
                     ev=ev, mc=mc, er_core=er_core, er_now=er_now, g=g,
                     wrecked=(cur < 0.80*pre)))
D = pd.DataFrame(rows); D.to_parquet(SCR+"/core_prepivot.parquet", index=False)
W = D[D.wrecked].copy()
mm = lambda s: pd.Series(s).dropna().median()
print(f"\nelite-core names (CoreMoat>=7.8) with usable history: {len(D)}")
print(f"currently PIVOTED (latest margin <80% of pre-pivot peak): {len(W)}\n")

print("=== the pivot's damage (wrecked names) ===")
print(f"  pre-pivot NOI margin (peak 3yr) median: {mm(W.pre_margin)*100:5.1f}%   (median peak year {int(mm(W.pre_yr))})")
print(f"  current   NOI margin        median:     {mm(W.cur_margin)*100:5.1f}%")
print(f"  margin destroyed by the pivot median:   {mm(W.margin_destroyed)*100:5.0f}%")
print(f"  core earnings uplift (pre/current NOPAT) median: {mm(W.nopat_uplift):.2f}x")

print("\n=== are the core returns cheap? (implied return = normalized yield + moat growth) ===")
print(f"  on CURRENT (post-pivot) earnings: median {mm(W.er_now)*100:5.1f}%   | >=12%: {(W.er_now>=0.12).mean()*100:3.0f}%")
print(f"  on PRE-PIVOT core earnings:       median {mm(W.er_core)*100:5.1f}%   | >=12%: {(W.er_core>=0.12).mean()*100:3.0f}%  | >=15%: {(W.er_core>=0.15).mean()*100:3.0f}%")
print(f"  uplift from valuing the core: median {(mm(W.er_core)-mm(W.er_now))*100:+.1f} pts of implied return")
cheap = W[(W.er_core >= 0.12)]
print(f"\n  CHEAP ON THE CORE (implied core return >=12%): {len(cheap)} of {len(W)} pivoted names")
print("  top 12 by implied core return:")
show = cheap.sort_values("er_core", ascending=False).head(12)
for _, r in show.iterrows():
    print(f"    {r.t:10s} CoreMoat {r.coremoat:.1f} (grp {r.companymoat:.1f})  pre-margin {r.pre_margin*100:4.0f}% -> now {r.cur_margin*100:4.0f}%  "
          f"core ER {r.er_core*100:4.0f}% (vs now {r.er_now*100:4.0f}%)")
print("\nCAVEATS: today's moat; period-end market cap; current SALES may include subsidiary")
print("revenue (would overstate core earnings if subs are large); pre-pivot margin assumes the")
print("peak was the clean core, not a cyclical high.")
