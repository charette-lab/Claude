#!/usr/bin/env python3
"""build_hist_er.py — engine EXPECTED returns across 2000-2026, fundamentals only.

No stock prices needed: the engine's expected return er1 = f(fundamentals, market cap),
and the deep 30_file panels carry both back to ~2000. So we can preview how the Core-30
SIGNAL behaves through the dot-com bust and the GFC before the price history arrives.

Per (name, fiscal year) we compute er1 (5-yr price->intrinsic IRR) and wacc via the full
engine, then summarise the eligible opportunity set by year. Moat = today's v3.2 snapshot
(look-ahead on quality — flagged; to be replaced by point-in-time moat snapshots later).
This is NOT a return backtest; it is a signal/opportunity-set history.

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/build_hist_er.py
"""
import sys, os, warnings
import numpy as np, pandas as pd, openpyxl
warnings.filterwarnings("ignore")
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30, aip

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
num = lambda v: v if isinstance(v, (int, float)) else None
GATE, MOAT_MIN = 0.12, 6.5

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
moat = {str(r[ci["Ticker"]]): r[ci["CompanyMoat(v3.2)"]] for r in rws[1:]
        if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}

print("loading deep panel (1983-2026)...", flush=True)
by, idx = panel30.load(XLSB)
g = lambda r, k: (r[idx[k]] if k in idx and idx[k] < len(r) else None)
out = []
for t, rows in by.items():
    t = str(t); m = moat.get(t)
    if m is None:
        continue
    # keep the latest observation per fiscal year
    seen = {}
    for r in rows:
        d = g(r, "Period_End_Date") or g(r, "Date")
        d = pd.to_datetime(d, errors="coerce")
        if pd.isna(d):
            continue
        seen[d.year] = (d, r)
    for yr, (d, r) in seen.items():
        if yr < 2000:
            continue
        nopat, roiic, rr = num(g(r, "New Operating Income")), num(g(r, "ROICm_total - 7 years")), num(g(r, "average_C - 7 year"))
        mc, ev = num(g(r, "Market Capitalization")), num(g(r, "EV"))
        if None in (nopat, roiic, rr) or not mc or ev is None:
            continue
        ltd = num(g(r, "Debt - Long-Term - Total")) or 0.0
        std = num(g(r, "Short-Term Debt & Current Portion of Long-Term Debt")) or 0.0
        tax = num(g(r, "Income Tax Rate - Instrument"))
        fin = {"Company Name": t, "Instrument": t, "GICS Industry Group Name": g(r, "GICS Industry Group Name"),
               "Country of Headquarters": g(r, "Country of Headquarters"), "New Operating Income": nopat,
               "ROICm 7": roiic, "RR 7": rr, "Moat Score": m, "Gross debt": ltd+std,
               "Income Tax Rate - Instrument": (tax/100.0 if tax is not None else 0.25),
               "Net debt": ev-mc, "Sales": num(g(r, "Sales")), "Market Cap": mc}
        try:
            v = aip.value_and_return(fin, re=0.07, re2=0.12)
        except Exception:
            v = None
        if not v or v.get("er1") is None:
            continue
        out.append(dict(Instrument=t, year=yr, date=d, er1=float(v["er1"]),
                        wacc=float(v.get("wacc") or np.nan), moat=m))

E = pd.DataFrame(out)
E.to_parquet(SCR+"/hist_er_annual.parquet", index=False)
print(f"WROTE {len(E)} name-years / {E.Instrument.nunique()} names / {E.year.min()}-{E.year.max()}\n")

# ---- the opportunity set through the cycle ----
def summ(s):
    elig = s[(s.er1 >= GATE) & (s.moat >= MOAT_MIN)]
    return pd.Series(dict(names=s.Instrument.nunique(), median_er=s.er1.median(),
                          eligible=elig.Instrument.nunique(),
                          elig_pct=100*len(elig)/len(s) if len(s) else np.nan,
                          top30_er=s.er1.nlargest(30).median()))
tab = E.groupby("year").apply(summ)
tab["median_er"] = (tab["median_er"]*100).round(1); tab["top30_er"] = (tab["top30_er"]*100).round(1)
tab["elig_pct"] = tab["elig_pct"].round(0)
print("expected-return opportunity set by year (engine signal, no prices):")
print(tab.to_string(float_format=lambda x: f"{x:.0f}" if x == x else "nan"))
print("\nread: median_er = whole-universe median expected return; eligible = # names passing")
print("er>=12% AND moat>=6.5; top30_er = median expected return of the 30 highest-er names.")
print("Watch it collapse into the 2000 & 2021 peaks and spike at the 2002-03 & 2008-09 troughs.")
