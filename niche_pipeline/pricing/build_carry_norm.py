#!/usr/bin/env python3
"""build_carry_norm.py — normalize the carry signal for supply/demand (over-earning).

Raw carry rewards a high CURRENT ROIIC, which over-earners carry on inflated earnings
that then mean-revert. So the raw-carry book picks over-earners and realises falling
earnings. This applies the full engine's supply/demand normalization to the carry:
per (security,fiscal-year) compute dn_carry = raw_carry - adjusted_carry at the
year-end price (point-in-time signals), and subtract it flat across the year — the
same flat-downgrade construction used for the normalized total ER.

Output: carry_grid_norm.parquet [Instrument, Date, market_cap, carry_raw,
expected_return(=normalized carry), artifact]

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/build_carry_norm.py
"""
import sys, os
import numpy as np, pandas as pd, openpyxl
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30, aip, overearning

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
HIST2 = U+"29a7c3ae-hist_20260629_2.xlsx"
num = lambda v: v if isinstance(v, (int, float)) else None

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
moat = {str(r[ci["Ticker"]]): r[ci["CompanyMoat(v3.2)"]] for r in rws[1:]
        if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}


def finrow(t, r, g, m):
    nopat = num(g(r, "New Operating Income")); roiic = num(g(r, "ROICm_total - 7 years"))
    rr = num(g(r, "average_C - 7 year")); mc = num(g(r, "Market Capitalization")); ev = num(g(r, "EV"))
    if None in (nopat, roiic, rr) or not mc or ev is None:
        return None
    ltd = num(g(r, "Debt - Long-Term - Total")) or 0.0
    std = num(g(r, "Short-Term Debt & Current Portion of Long-Term Debt")) or 0.0
    tax = num(g(r, "Income Tax Rate - Instrument"))
    return g(r, "Period_End_Date") or g(r, "Date"), {
        "Company Name": t, "Instrument": t, "GICS Industry Group Name": g(r, "GICS Industry Group Name"),
        "Country of Headquarters": g(r, "Country of Headquarters"), "New Operating Income": nopat,
        "ROICm 7": roiic, "RR 7": rr, "Moat Score": m, "Gross debt": ltd+std,
        "Income Tax Rate - Instrument": (tax/100.0 if tax is not None else 0.25),
        "Net debt": ev-mc, "Sales": num(g(r, "Sales")), "Market Cap": mc}


def dn_rows(by, idx, recs):
    g = lambda r, k: (r[idx[k]] if k in idx and idx[k] < len(r) else None)
    for t, rows in by.items():
        t = str(t); m = moat.get(t)
        if m is None:
            continue
        for i in range(len(rows)):
            fr = finrow(t, rows[i], g, m)
            if not fr:
                continue
            ped, fin = fr
            try:
                sig = overearning.panel_signals(rows[:i+1], idx)
                ts = overearning.two_stage_return(fin, sig, re=0.07, re2=0.12)
                stated = aip.value_and_return(fin, re=0.07, re2=None)
                raw_c = stated.get("er1_carry") if stated else None
                adj_c = ts.get("er_carry") if ts else None
                dn = max(0.0, raw_c - adj_c) if (raw_c is not None and adj_c is not None) else 0.0
            except Exception:
                dn = 0.0
            recs.append({"Instrument": t, "ped": ped, "dn_carry": dn})


recs = []
print("loading panels...", flush=True)
by, idx = panel30.load(XLSB)
dn_rows(by, idx, recs)
print(f"  main done ({len(recs)})", flush=True)
h = pd.read_excel(HIST2); h["Instrument"] = h["Instrument"].astype(str)
hidx = {c: i for i, c in enumerate(h.columns)}; hidx["Period_End_Date"] = hidx.get("Date")
by_x = {t: [list(r) for r in g.sort_values("Date").itertuples(index=False, name=None)]
        for t, g in h.groupby("Instrument")}
dn_rows(by_x, hidx, recs)
DN = pd.DataFrame(recs)
DN["ped"] = pd.to_datetime(DN["ped"], errors="coerce").astype("datetime64[ns]")
DN = DN.dropna(subset=["ped"]).sort_values("ped")
print(f"dn_carry rows: {len(DN)} | mean carry downgrade {DN.dn_carry.mean()*100:.2f}pp", flush=True)

carry = pd.read_parquet(SCR+"/carry_grid.parquet"); carry["Instrument"] = carry["Instrument"].astype(str)
carry["Date"] = pd.to_datetime(carry["Date"]).astype("datetime64[ns]")
carry = carry.rename(columns={"expected_return": "carry_raw"}).sort_values("Date")
m = pd.merge_asof(carry, DN.sort_values("ped"), left_on="Date", right_on="ped",
                  by="Instrument", direction="backward")
m["dn_carry"] = m["dn_carry"].fillna(0.0)
m["expected_return"] = m["carry_raw"] - m["dn_carry"]
out = m[["Instrument", "Date", "market_cap", "carry_raw", "expected_return", "artifact"]].sort_values(["Instrument", "Date"])
out.to_parquet(SCR+"/carry_grid_norm.parquet", index=False)
print(f"WROTE {len(out)} rows / {out.Instrument.nunique()} securities -> carry_grid_norm.parquet")
print(f"normalized carry >=12%: {(out.expected_return>=0.12).mean()*100:.0f}% (raw was higher)")
print("DONE_CARRYNORM")
