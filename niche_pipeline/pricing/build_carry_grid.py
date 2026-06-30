#!/usr/bin/env python3
"""build_carry_grid.py — month-end CARRY expected-return panel.

The engine splits expected return into CARRY (the IRR earned if the price never
re-rates — interim distributions + exit at the ENTRY enterprise value, i.e. pure
internal ROIIC compounding) and RE-RATE (the bonus from price closing to intrinsic).
This builds the daily-ish CARRY signal so a book can be SELECTED on internal
compounding rather than on re-rating.

Computed at month-end dates only (the rebalance grid), point-in-time: the carry at
date t uses the most recent fiscal-year fundamentals <= t and the market cap at t.
Artifact flag is carried over from the full panel (unreasonable raw total return).

Output: carry_grid.parquet [Instrument, Date, market_cap, expected_return(=carry), artifact]

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/build_carry_grid.py
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


def fins_from_panel(by, idx):
    g = lambda r, k: (r[idx[k]] if k in idx and idx[k] < len(r) else None)
    rows = []
    for t, rs in by.items():
        t = str(t); m = moat.get(t)
        if m is None:
            continue
        for r in rs:
            nopat = num(g(r, "New Operating Income")); roiic = num(g(r, "ROICm_total - 7 years"))
            rr = num(g(r, "average_C - 7 year")); mc = num(g(r, "Market Capitalization"))
            ev = num(g(r, "EV"))
            if None in (nopat, roiic, rr) or not mc or ev is None:
                continue
            ped = g(r, "Period_End_Date") or g(r, "Date")
            ltd = num(g(r, "Debt - Long-Term - Total")) or 0.0
            std = num(g(r, "Short-Term Debt & Current Portion of Long-Term Debt")) or 0.0
            tax = num(g(r, "Income Tax Rate - Instrument"))
            rows.append(dict(Instrument=t, ped=ped, moat=m, nopat=nopat, roiic=roiic, rr=rr,
                             sales=num(g(r, "Sales")), tax=(tax/100.0 if tax is not None else 0.25),
                             gross=ltd+std, netdebt=ev-mc,
                             ind=g(r, "GICS Industry Group Name"), ctry=g(r, "Country of Headquarters")))
    return rows


print("loading panels...", flush=True)
by, idx = panel30.load(XLSB)
rows = fins_from_panel(by, idx)
h = pd.read_excel(HIST2); h["Instrument"] = h["Instrument"].astype(str)
hidx = {c: i for i, c in enumerate(h.columns)}; hidx["Period_End_Date"] = hidx.get("Date")
by_x = {t: [list(r) for r in g.sort_values("Date").itertuples(index=False, name=None)]
        for t, g in h.groupby("Instrument")}
rows += fins_from_panel(by_x, hidx)
F = pd.DataFrame(rows)
F["ped"] = pd.to_datetime(F["ped"], errors="coerce").astype("datetime64[ns]")
F = F.dropna(subset=["ped"]).sort_values("ped")
print(f"fins: {len(F)} / {F.Instrument.nunique()} securities", flush=True)

# month-end market cap grid
mc = pd.read_parquet(SCR+"/daily_marketcap_ev.parquet", columns=["Instrument", "Date", "market_cap"])
mc["Instrument"] = mc["Instrument"].astype(str); mc["Date"] = pd.to_datetime(mc["Date"]).astype("datetime64[ns]")
me = mc["Date"].dt.to_period("M")
grid_mc = mc[mc.groupby(["Instrument", me])["Date"].transform("max") == mc["Date"]]   # last obs each month
grid_mc = grid_mc[grid_mc["market_cap"] > 0].sort_values("Date")

# attach the most-recent fiscal-year fundamentals to each month-end
m = pd.merge_asof(grid_mc, F.sort_values("ped"), left_on="Date", right_on="ped",
                  by="Instrument", direction="backward").dropna(subset=["ped"])
print(f"month-end rows with fundamentals: {len(m)}", flush=True)

# carry per (Instrument, ped): er1_carry is monotonic in price -> sample 12 points, interpolate
out = []
done = 0
for (t, ped), grp in m.groupby(["Instrument", "ped"]):
    r0 = grp.iloc[0]
    base = {"Company Name": t, "Instrument": t, "GICS Industry Group Name": r0["ind"],
            "Country of Headquarters": r0["ctry"], "New Operating Income": r0["nopat"],
            "ROICm 7": r0["roiic"], "RR 7": r0["rr"], "Moat Score": r0["moat"],
            "Gross debt": r0["gross"], "Income Tax Rate - Instrument": r0["tax"],
            "Net debt": r0["netdebt"], "Sales": r0["sales"]}
    mcs = grp["market_cap"].values
    lo, hi = np.nanmin(mcs), np.nanmax(mcs)
    if not (lo > 0):
        continue
    gx = np.unique(np.linspace(lo, hi, 12)) if hi > lo else np.array([lo])
    cy = []
    for x in gx:
        f = dict(base); f["Market Cap"] = x
        v = aip.value_and_return(f, re=0.07, re2=None)
        cy.append(v.get("er1_carry") if v else None)
    gd = [(x, c) for x, c in zip(gx, cy) if c is not None]
    if not gd:
        continue
    xs = np.array([x for x, _ in gd]); cs = np.array([c for _, c in gd])
    out.append(pd.DataFrame({"Instrument": t, "Date": grp["Date"].values,
                             "market_cap": mcs, "expected_return": np.interp(mcs, xs, cs)}))
    done += 1
    if done % 5000 == 0:
        print(f"  {done} (security,year) groups", flush=True)
carry = pd.concat(out, ignore_index=True)

# carry over the artifact flag from the full panel (asof by Instrument/Date)
full = pd.read_parquet(SCR+"/daily_expected_return_full.parquet", columns=["Instrument", "Date", "artifact"])
full["Instrument"] = full["Instrument"].astype(str); full["Date"] = pd.to_datetime(full["Date"]).astype("datetime64[ns]")
carry = carry.sort_values("Date")
carry = pd.merge_asof(carry, full.sort_values("Date"), on="Date", by="Instrument", direction="backward")
carry["artifact"] = carry["artifact"].fillna(False)
carry = carry.sort_values(["Instrument", "Date"])
carry.to_parquet(SCR+"/carry_grid.parquet", index=False)
print(f"WROTE {len(carry)} rows / {carry.Instrument.nunique()} securities -> carry_grid.parquet")
print(f"carry distribution: mean {carry.expected_return.mean()*100:.1f}%  median {carry.expected_return.median()*100:.1f}%  "
      f">=12%: {(carry.expected_return>=0.12).mean()*100:.0f}%")
print("DONE_CARRY")
