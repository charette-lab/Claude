#!/usr/bin/env python3
"""build_floor.py — no-growth valuation floor per (security, fiscal-year), for Gate 2.

Gate 2 (IPS): P(30% drawdown) <= 20%, measured vs a fundamentally justified hard
floor. The floor is the NO-GROWTH value of the business: current normalized
operating earnings capitalized at the company's own cost of capital, less net debt
(the "value of assets in place"). floor_equity = NOPAT / WACC - net_debt.
floor_ratio (computed daily downstream) = floor_equity / market_cap; >= 0.70 means
a 30% fall lands at/below the floor.

Output: floor_equity.parquet [Instrument, ped, floor_equity, nopat, wacc, netdebt]

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/build_floor.py
"""
import sys, os
import numpy as np, pandas as pd, openpyxl
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30, aip

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
HIST2 = U+"29a7c3ae-hist_20260629_2.xlsx"
num = lambda v: v if isinstance(v, (int, float)) else None

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
moat = {str(r[ci["Ticker"]]): r[ci["CompanyMoat(v3.2)"]] for r in rws[1:]
        if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}
ind = {str(r[ci["Ticker"]]): r[ci["Industry"]] for r in rws[1:] if r[ci["Ticker"]]}
ctry = {str(r[ci["Ticker"]]): r[ci["Country"]] for r in rws[1:] if r[ci["Ticker"]]}
out = []


def go(by, idx, label):
    g = lambda r, k: (r[idx[k]] if k in idx and idx[k] < len(r) else None)
    n = 0
    for t, rows in by.items():
        t = str(t); m = moat.get(t)
        if m is None:
            continue
        for r in rows:
            nopat = num(g(r, "New Operating Income")); roiic = num(g(r, "ROICm_total - 7 years"))
            rr = num(g(r, "average_C - 7 year")); mc = num(g(r, "Market Capitalization")); ev = num(g(r, "EV"))
            if None in (nopat, roiic, rr) or not mc or ev is None:
                continue
            ped = g(r, "Period_End_Date") or g(r, "Date")
            ltd = num(g(r, "Debt - Long-Term - Total")) or 0.0
            std = num(g(r, "Short-Term Debt & Current Portion of Long-Term Debt")) or 0.0
            tax = num(g(r, "Income Tax Rate - Instrument"))
            netdebt = ev - mc
            fin = {"Company Name": t, "Instrument": t, "GICS Industry Group Name": g(r, "GICS Industry Group Name"),
                   "Country of Headquarters": g(r, "Country of Headquarters"), "New Operating Income": nopat,
                   "ROICm 7": roiic, "RR 7": rr, "Moat Score": m, "Gross debt": ltd+std,
                   "Income Tax Rate - Instrument": (tax/100.0 if tax is not None else 0.25),
                   "Net debt": netdebt, "Sales": num(g(r, "Sales")), "Market Cap": mc}
            try:
                v = aip.value_and_return(fin, re=0.07, re2=None)
                wacc = v.get("wacc") if v else None
            except Exception:
                wacc = None
            if wacc is None or wacc <= 0 or nopat is None:
                continue
            floor_eq = nopat / wacc - netdebt          # no-growth EV minus net debt
            out.append(dict(Instrument=t, ped=ped, floor_equity=floor_eq, nopat=nopat, wacc=wacc, netdebt=netdebt))
            n += 1
    print(f"  [{label}] {n}", flush=True)


print("loading panels...", flush=True)
by, idx = panel30.load(XLSB)
go(by, idx, "main")
h = pd.read_excel(HIST2); h["Instrument"] = h["Instrument"].astype(str)
hidx = {c: i for i, c in enumerate(h.columns)}; hidx["Period_End_Date"] = hidx.get("Date")
by_x = {t: [list(r) for r in g.sort_values("Date").itertuples(index=False, name=None)]
        for t, g in h.groupby("Instrument")}
go(by_x, hidx, "extra")

F = pd.DataFrame(out)
F["ped"] = pd.to_datetime(F["ped"], errors="coerce").astype("datetime64[ns]")
F = F.dropna(subset=["ped"]).sort_values("ped")
F.to_parquet(SCR+"/floor_equity.parquet", index=False)
print(f"WROTE {len(F)} rows / {F.Instrument.nunique()} securities -> floor_equity.parquet")
print(f"median wacc {F.wacc.median()*100:.1f}%  | positive floor_equity: {(F.floor_equity>0).mean()*100:.0f}%")
print("DONE_FLOOR")
