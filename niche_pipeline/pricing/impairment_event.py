#!/usr/bin/env python3
"""impairment_event.py — engine expected return PRE vs POST an impairment, and the
forward return from investing after it. Answers: does the write-off create the
opportunity (ER jumps), and does buying post-impairment capture it?

For each elite-core (CoreMoat>=7.8) name we compute the REAL engine er_adj at each
fiscal year using that year's period-end market cap (point-in-time sig + artifact screen),
plus the impairment booked and the period-end price. An impairment EVENT year E is one
where impairment > 2% of EV. We then line up er at E-2..E+2 and the forward 1yr/2yr price
returns from E-1 (pre) vs E+1 (post) to see when the opportunity actually appears.

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/impairment_event.py
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
coremoat = col("CoreMoat(v3.2)"); companymoat = col("CompanyMoat(v3.2)")
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

print("loading panel + computing per-year engine ER...", flush=True)
by, idx = panel30.load(XLSB)
def gr(rows, i, k): return rows[i][idx[k]] if (k in idx and idx[k] < len(rows[i])) else None

recs = []
for t, rows in by.items():
    t = str(t)
    if coremoat.get(t, 0) < 7.8: continue
    for i in range(len(rows)):
        g = lambda k: gr(rows, i, k)
        nopat, roiic, rr = num(g("New Operating Income")), num(g("ROICm_total - 7 years")), num(g("average_C - 7 year"))
        mc, ev, sal = num(g("Market Capitalization")), num(g("EV")), num(g("Sales"))
        if None in (nopat, roiic, rr) or not mc or ev is None or not sal or sal <= 0: continue
        ped = pd.Timestamp(g("Period_End_Date") or g("Date")); yr = ped.year
        if yr < 2010: continue
        ltd = num(g("Debt - Long-Term - Total")) or 0.0; std = num(g("Short-Term Debt & Current Portion of Long-Term Debt")) or 0.0
        tax = num(g("Income Tax Rate - Instrument"))
        imp = sum(abs(num(g(k)) or 0.0) for k in IMPAIR)
        fin = {"Company Name": t, "Instrument": t, "GICS Industry Group Name": ind.get(t),
               "Country of Headquarters": ctry.get(t), "New Operating Income": nopat, "ROICm 7": roiic, "RR 7": rr,
               "Moat Score": companymoat.get(t, coremoat[t]), "Gross debt": ltd+std,
               "Income Tax Rate - Instrument": (tax/100.0 if tax is not None else 0.25), "Net debt": ev-mc,
               "Sales": sal, "Market Cap": mc}
        try:
            sig = overearning.panel_signals(rows[:i+1], idx)
            ts = overearning.two_stage_return(fin, sig, re=0.07, re2=0.12)
            er = ts.get("er_adj") if ts else None
        except Exception:
            er = None
        recs.append(dict(t=t, yr=yr, ped=ped, er=er, imp_frac=imp/ev if ev > 0 else 0.0,
                         margin=nopat/sal, price=pasof(t, ped)))
D = pd.DataFrame(recs)
print(f"  elite-core name-years with engine ER: {len(D)} / {D.t.nunique()} names", flush=True)

# forward price return helper
def fret(t, y0, h):
    d0, d1 = pd.Timestamp(y0, 12, 31), pd.Timestamp(y0+h, 12, 31)
    p0, p1 = pasof(t, d0), pasof(t, d1); return (p1/p0-1) if (p0 and p1 and p0 > 0) else None

byt = {t: g.set_index("yr") for t, g in D.groupby("t")}
ev_rows = []
for t, g in byt.items():
    for E in g.index:
        if g.loc[E, "imp_frac"] <= 0.02:            # material impairment event
            continue
        row = {"t": t, "E": E}
        for k in (-2, -1, 0, 1, 2):
            row[f"er{k:+d}"] = g.loc[E+k, "er"] if (E+k in g.index) else np.nan
            row[f"m{k:+d}"] = g.loc[E+k, "margin"] if (E+k in g.index) else np.nan
        row["fwd1_pre"] = fret(t, E-1, 1)           # invest end of year BEFORE impairment
        row["fwd1_post"] = fret(t, E+1, 1)          # invest end of year AFTER impairment (what my test did)
        row["fwd2_post"] = fret(t, E+1, 2)
        ev_rows.append(row)
E = pd.DataFrame(ev_rows)
Erec = E[(E.E >= 2015) & (E.E <= 2024)]              # events in the price window / my test window
med = lambda s: pd.Series(s).dropna().median()
print(f"\nimpairment events (elite-core, imp>2% of EV): {len(E)} total; {len(Erec)} in 2015-2024 (my test window)\n")
print("=== ENGINE EXPECTED RETURN around the impairment (median, real er_adj) ===")
print(f"  E-2:{med(Erec['er-2'])*100:5.1f}%   E-1:{med(Erec['er-1'])*100:5.1f}%   E(impair):{med(Erec['er+0'])*100:5.1f}%   "
      f"E+1:{med(Erec['er+1'])*100:5.1f}%   E+2:{med(Erec['er+2'])*100:5.1f}%")
print("=== NOI margin around the impairment (median) ===")
print(f"  E-2:{med(Erec['m-2'])*100:5.1f}%   E-1:{med(Erec['m-1'])*100:5.1f}%   E(impair):{med(Erec['m+0'])*100:5.1f}%   "
      f"E+1:{med(Erec['m+1'])*100:5.1f}%   E+2:{med(Erec['m+2'])*100:5.1f}%")
print("\n=== forward 1yr price return by ENTRY timing ===")
print(f"  invest END of E-1 (PRE impairment):  median {med(Erec['fwd1_pre'])*100:5.1f}%  (n={Erec['fwd1_pre'].notna().sum()})")
print(f"  invest END of E+1 (POST impairment): median {med(Erec['fwd1_post'])*100:5.1f}%  (n={Erec['fwd1_post'].notna().sum()})")
print(f"  invest END of E+1, hold 2yr (POST):  median {med(Erec['fwd2_post'])*100:5.1f}%  (n={Erec['fwd2_post'].notna().sum()})")
print(f"\n  share of events where ER rose from E-1 to E+1 (impairment made it cheaper): "
      f"{(Erec['er+1']>Erec['er-1']).mean()*100:.0f}%")
E.to_parquet(SCR+"/impairment_event.parquet")
print("\nNote: ER uses period-end market cap each year (no daily prices needed); forward returns 2016+ only.")
