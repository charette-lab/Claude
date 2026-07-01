#!/usr/bin/env python3
"""stage2_realengine.py — buried-core Stage 2 with the REAL full engine (no proxy).

Same point-in-time backtest as stage2_buriedcore, but cheapness is the actual engine
expected return on the CORE: feed core earnings (pre-pivot NOI margin x sales) into
overearning.two_stage_return (supply/demand-normalized er_adj) with the quality-aware
artifact screen, exactly as the production decomposition does. Select elite-core names
that are wrecked, non-artifact, and have engine core-ER >= 12%. Compares selecting on the
CORE valuation vs the GROUP (reported) valuation, and the impairment split.

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/stage2_realengine.py
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
PRICE = [U+"972f0581-daily_volume_price_0.parquet", U+"257124b3-daily_volume_price_1.parquet",
         U+"86e54ec3-daily_volume_price_2.parquet", U+"13f82c18-daily_volume_price_0630.parquet"]
num = lambda v: v if isinstance(v, (int, float)) else None
IMPAIR = ["Adjusted_Impairment___Goodwill__ASR", "Adjusted_Impairment___Intangibles_excluding_Goodwill__ASR",
          "Adjusted_Impairment___Fixed_Assets", "Adjusted_Impairment___Investment_Property___Cash_Flow"]

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
col = lambda n: {str(r[ci["Ticker"]]): r[ci[n]] for r in rws[1:] if r[ci["Ticker"]] and n in ci and isinstance(r[ci[n]], (int, float))}
coremoat, companymoat = col("CoreMoat(v3.2)"), col("CompanyMoat(v3.2)")
ind = {str(r[ci["Ticker"]]): r[ci["Industry"]] for r in rws[1:] if r[ci["Ticker"]]}
ctry = {str(r[ci["Ticker"]]): r[ci["Country"]] for r in rws[1:] if r[ci["Ticker"]]}

print("loading prices...", flush=True)
px = pd.concat([pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"]) for f in PRICE])
px["Instrument"] = px["Instrument"].astype(str); px["Date"] = pd.to_datetime(px["Date"])
px["c"] = pd.to_numeric(px["Close Price"], errors="coerce")
W = px[px.c > 0].pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()
def pasof(t, d):
    if t not in W.columns: return None
    s = W[t].loc[:d].dropna(); return float(s.iloc[-1]) if len(s) else None

print("loading panel + precomputing point-in-time engine inputs...", flush=True)
by, idx = panel30.load(XLSB)
def gr(rows, i, k): return rows[i][idx[k]] if (k in idx and idx[k] < len(rows[i])) else None

def finrow(rows, i):
    g = lambda k: gr(rows, i, k)
    nopat, roiic = num(g("New Operating Income")), num(g("ROICm_total - 7 years"))
    rr, mc, ev = num(g("average_C - 7 year")), num(g("Market Capitalization")), num(g("EV"))
    sal = num(g("Sales"))
    if None in (nopat, roiic, rr) or not mc or ev is None or not sal or sal <= 0:
        return None
    ltd = num(g("Debt - Long-Term - Total")) or 0.0; std = num(g("Short-Term Debt & Current Portion of Long-Term Debt")) or 0.0
    tax = num(g("Income Tax Rate - Instrument"))
    ped = g("Period_End_Date") or g("Date")
    imp = sum(abs(num(g(k)) or 0.0) for k in IMPAIR)
    return pd.Timestamp(ped), dict(nopat=nopat, roiic=roiic, rr=rr, sales=sal, mc=mc,
                                   tax=(tax/100.0 if tax is not None else 0.25), gross=ltd+std, netdebt=ev-mc, imp=imp)

def base_fin(t, f, moat, nopat):
    return {"Company Name": t, "Instrument": t, "GICS Industry Group Name": ind.get(t),
            "Country of Headquarters": ctry.get(t), "New Operating Income": nopat, "ROICm 7": f["roiic"],
            "RR 7": f["rr"], "Moat Score": moat, "Gross debt": f["gross"], "Income Tax Rate - Instrument": f["tax"],
            "Net debt": f["netdebt"], "Sales": f["sales"]}

REC = {}   # t -> {year: rec}
for t, rows in by.items():
    t = str(t)
    if coremoat.get(t, 0) < 7.8: continue
    margins = []; per = {}
    for i in range(len(rows)):
        fr = finrow(rows, i)
        if not fr: continue
        ped, f = fr; yr = ped.year
        margins.append((yr, f["nopat"]/f["sales"]))
        if yr < 2014: continue
        mser = pd.Series({y: mm for y, mm in margins})
        valid = mser[mser.abs() <= 0.60]
        if len(valid) < 5: continue
        pre = valid.rolling(3, min_periods=2).mean().max()
        cur = mser.rolling(3, min_periods=1).mean().iloc[-1]
        if pre is None or not (0 < pre <= 0.50): continue
        try:
            sig = overearning.panel_signals(rows[:i+1], idx)
            mvh = history.verdict(history.summarize(rows[:i+1], idx), coremoat[t])[0]
        except Exception:
            continue
        per[yr] = dict(ped=ped, f=f, sig=sig, mvh=mvh, rstar=sig.get("roic_star"),
                       pre=pre, cur=cur, wrecked=(cur < 0.80*pre),
                       bf_core=base_fin(t, f, coremoat[t], pre*f["sales"]),
                       bf_grp=base_fin(t, f, companymoat.get(t, coremoat[t]), f["nopat"]))
    if per: REC[t] = per
print(f"  elite-core names with point-in-time records: {len(REC)}", flush=True)

def eng(bf, sig, mvh, rstar, mc):
    fin = dict(bf); fin["Market Cap"] = mc
    try:
        ts = overearning.two_stage_return(fin, sig, re=0.07, re2=0.12)
    except Exception:
        return None, False
    if not ts: return None, False
    erc = ts.get("er_current")
    art = bool(F.er_is_artifact(erc, rstar, mvh)) if erc is not None else False
    return ts.get("er_adj"), art

def hold(t, d0, d1):
    p0, p1 = pasof(t, d0), pasof(t, d1); return (p1/p0-1) if (p0 and p1 and p0 > 0) else None

rows_out = []; universe = [str(c) for c in W.columns]
for yt in range(2016, 2026):
    d0, d1 = pd.Timestamp(yt, 6, 30), pd.Timestamp(yt+1, 6, 30); Y = yt-1
    core_ret, grp_ret, imp_on, imp_off = {}, {}, [], []
    for t, per in REC.items():
        yrs = [y for y in per if y <= Y]
        if not yrs: continue
        rec = per[max(yrs)]
        p_pe, p_D = pasof(t, rec["ped"]), pasof(t, d0)
        if not p_pe or not p_D: continue
        mc_D = rec["f"]["mc"]*(p_D/p_pe)
        if mc_D <= 0: continue
        er_c, art_c = eng(rec["bf_core"], rec["sig"], rec["mvh"], rec["rstar"], mc_D)
        er_g, _ = eng(rec["bf_grp"], rec["sig"], rec["mvh"], rec["rstar"], mc_D)
        r = hold(t, d0, d1)
        if r is None: continue
        if rec["wrecked"] and not art_c and er_c is not None and er_c >= 0.12:   # cheap on the CORE, real engine
            core_ret[t] = r
            (imp_on if rec["f"]["imp"] > 0.02*(rec["f"]["mc"]+rec["f"]["netdebt"]) else imp_off).append(r)
        if er_g is not None and er_g >= 0.12:                                    # cheap on the GROUP (reported)
            grp_ret[t] = r
    if len(core_ret) < 4: continue
    uni = [v for v in (hold(t, d0, d1) for t in universe) if v is not None]
    elite = [v for v in (hold(t, d0, d1) for t in REC) if v is not None]
    rows_out.append(dict(year=yt, n=len(core_ret), core=np.mean(list(core_ret.values())),
                         grp=np.mean(list(grp_ret.values())) if grp_ret else np.nan,
                         elite=np.mean(elite), universe=np.mean(uni),
                         imp_on=np.mean(imp_on) if imp_on else np.nan,
                         imp_off=np.mean(imp_off) if imp_off else np.nan, n_imp=len(imp_on)))
R = pd.DataFrame(rows_out).set_index("year"); R.to_parquet(SCR+"/stage2_realengine.parquet")
def cagr(s): s = (1+pd.Series(s).dropna()); return s.prod()**(1/len(s))-1
print(f"\n{'year':6}{'n':>4}{'core-ENG%':>10}{'group-ENG%':>11}{'elite%':>8}{'univ%':>7}")
for y, r in R.iterrows():
    print(f"{y:<6}{int(r.n):>4}{r.core*100:>9.1f}{r.grp*100:>11.1f}{r.elite*100:>8.1f}{r.universe*100:>7.1f}")
print("\n--- CAGR (equal-weight, 1yr hold, gross) ---")
for nm, c in [("BURIED-CORE (real engine er_core>=12%)", R.core), ("Select on GROUP engine er>=12%", R.grp),
              ("Elite-core (all)", R.elite), ("Universe EW", R.universe)]:
    s = pd.Series(c).dropna(); eq = (1+s).cumprod(); dd = (eq/eq.cummax()-1).min()
    print(f"  {nm:40s} CAGR {cagr(c)*100:5.1f}%  vol {s.std(ddof=1)*100:4.1f}%  worstyr {s.min()*100:5.1f}%  maxDD {dd*100:5.1f}%")
print(f"\n  core-engine beat universe {int((R.core>R.universe).sum())}/{len(R)} yrs, beat elite {int((R.core>R.elite).sum())}/{len(R)}, "
      f"beat group-selection {int((R.core>R.grp).sum())}/{int(R.grp.notna().sum())}")
print(f"\n  impairment split (real-engine candidates): WITH {cagr(R.imp_on)*100:.1f}%  vs WITHOUT {cagr(R.imp_off)*100:.1f}%")
print("\nCAVEATS: gross price-only; today's CoreMoat backward; benign 2016-26 window.")
