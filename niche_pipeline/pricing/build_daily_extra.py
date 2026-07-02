#!/usr/bin/env python3
"""build_daily_extra.py — value the names that live in the 132-col hist part-file
(hist_20260629_2.xlsx), not the 30-year .xlsb panel, and append them to the daily
MC/EV and ER panels.

These ~157 names (e.g. Workday) were absent from the .xlsb panel, so the engine
never valued them. The hist part-file carries the SAME rich schema (spaced column
names), so we extract per-(security,year) fundamentals exactly as build_daily_er
does, pair them with the moats from Universe_final and the 0630 daily prices, and
build daily MC/EV + daily ER on the identical methodology, then concat.

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/build_daily_extra.py
"""
import sys, os
import numpy as np, pandas as pd, openpyxl
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import aip

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
HIST2 = U + "29a7c3ae-hist_20260629_2.xlsx"
PRICES = [U + "13f82c18-daily_volume_price_0630.parquet",
          U + "972f0581-daily_volume_price_0.parquet",
          U + "257124b3-daily_volume_price_1.parquet",
          U + "86e54ec3-daily_volume_price_2.parquet"]
REQ = ["New Operating Income", "ROICm_total - 7 years", "average_C - 7 year",
       "Market Capitalization", "EV"]

# ---- moats (and industry/country fallback) from the unified scored book --------
ws = openpyxl.load_workbook(SCR + "/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rows = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rows[0])}
moat = {r[ci["Ticker"]]: r[ci["CompanyMoat(v3.2)"]] for r in rows[1:]
        if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}

# ---- which names still need valuing -------------------------------------------
er_have = set(pd.read_parquet(SCR + "/daily_expected_return.parquet",
                              columns=["Instrument"])["Instrument"].astype(str).unique())

print("loading hist part-file...", flush=True)
h = pd.read_excel(HIST2)
h["Instrument"] = h["Instrument"].astype(str)
h["Date"] = pd.to_datetime(h["Date"], errors="coerce").astype("datetime64[ns]")
h = h.dropna(subset=["Date"])
extra = sorted(set(h["Instrument"]) - er_have)
h = h[h["Instrument"].isin(extra)]

# ---- per-(security,year) fundamentals -----------------------------------------
fin = []
for t, g in h.groupby("Instrument"):
    m = moat.get(t)
    if m is None:
        continue
    for _, r in g.iterrows():
        if any(pd.isna(r[c]) for c in REQ):
            continue
        ltd = r.get("Debt - Long-Term - Total") or 0.0
        std = r.get("Short-Term Debt & Current Portion of Long-Term Debt") or 0.0
        tax = r.get("Income Tax Rate - Instrument")
        fin.append(dict(Instrument=t, ped=r["Date"], moat=m,
                        nopat=r["New Operating Income"], roiic=r["ROICm_total - 7 years"],
                        rr=r["average_C - 7 year"], sales=r.get("Sales"),
                        tax=(tax/100.0 if pd.notna(tax) else 0.25),
                        gross=(ltd if pd.notna(ltd) else 0)+(std if pd.notna(std) else 0),
                        netdebt=r["EV"]-r["Market Capitalization"], mc_ye=r["Market Capitalization"],
                        ind=r.get("GICS Industry Group Name"), ctry=r.get("Country of Headquarters")))
F = pd.DataFrame(fin).sort_values("ped")
F["ped"] = F["ped"].astype("datetime64[ns]")
print(f"per-(security,year) fins: {len(F)} across {F.Instrument.nunique()} securities", flush=True)

# ---- daily prices for these names ---------------------------------------------
pf = []
for p in PRICES:
    d = pd.read_parquet(p, columns=["Instrument", "Date", "Close Price"])
    d["Instrument"] = d["Instrument"].astype(str)
    pf.append(d[d["Instrument"].isin(F.Instrument.unique())])
px = pd.concat(pf, ignore_index=True).rename(columns={"Close Price": "close"})
px["Date"] = pd.to_datetime(px["Date"], errors="coerce").astype("datetime64[ns]")
px = px.dropna(subset=["Date", "close"])
px = px[px["close"] > 0].drop_duplicates(["Instrument", "Date"]).sort_values("Date")
print(f"daily price rows: {len(px)} / {px.Instrument.nunique()} securities", flush=True)

# ---- daily MC/EV: shares = yearly MC / close@period-end (nearest 10d) ----------
dd = px.rename(columns={"Date": "d", "close": "c"})[["Instrument", "d", "c"]].sort_values("d")
yr = F[["Instrument", "ped", "mc_ye", "netdebt"]].rename(columns={"mc_ye": "mc"}).sort_values("ped")
yr2 = pd.merge_asof(yr, dd, left_on="ped", right_on="d", by="Instrument",
                    direction="nearest", tolerance=pd.Timedelta("10D"))
yr2["shares"] = yr2["mc"] / yr2["c"]
step = yr2.dropna(subset=["shares"])[["Instrument", "ped", "shares", "netdebt"]].sort_values("ped")
mc = pd.merge_asof(px.sort_values("Date"), step, left_on="Date", right_on="ped",
                   by="Instrument", direction="backward")
mc = mc.sort_values(["Instrument", "Date"])
mc[["shares", "netdebt"]] = mc.groupby("Instrument")[["shares", "netdebt"]].bfill()
mc = mc.dropna(subset=["shares"])
mc["market_cap"] = mc["shares"] * mc["close"]
mc["enterprise_value"] = mc["market_cap"] + mc["netdebt"]
mcev = mc[["Instrument", "Date", "close", "shares", "netdebt", "market_cap", "enterprise_value"]] \
    .rename(columns={"netdebt": "net_debt"})
print(f"daily MC/EV rows: {len(mcev)} / {mcev.Instrument.nunique()} securities", flush=True)

# ---- daily ER: attach most-recent year-end valuation, 20-pt price interp -------
anc = pd.merge_asof(mcev[["Instrument", "Date", "market_cap"]].sort_values("Date"),
                    F.sort_values("ped"), left_on="Date", right_on="ped",
                    by="Instrument", direction="backward").dropna(subset=["ped", "mc_ye"])

def er_at(base, mcv):
    f = dict(base); f["Market Cap"] = mcv
    v = aip.value_and_return(f, re=0.07, re2=0.12)
    return v.get("er1") if v else None

out = []
for (t, ped), grp in anc.groupby(["Instrument", "ped"]):
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
    grid = np.unique(np.linspace(lo, hi, 20)) if hi > lo else np.array([lo])
    gd = [(x, er_at(base, x)) for x in grid]
    gd = [(x, e) for x, e in gd if e is not None]
    if not gd:
        continue
    gx = np.array([x for x, _ in gd]); ge = np.array([e for _, e in gd])
    out.append(pd.DataFrame({"Instrument": t, "Date": grp["Date"].values,
                             "market_cap": mcs, "expected_return": np.interp(mcs, gx, ge)}))
er = pd.concat(out, ignore_index=True)
print(f"daily ER rows: {len(er)} / {er.Instrument.nunique()} securities", flush=True)

# ---- append to the panels (de-dup on Instrument+Date, new wins) ----------------
def append(path, new, keys):
    old = pd.read_parquet(path)
    old["Instrument"] = old["Instrument"].astype(str)
    old["Date"] = pd.to_datetime(old["Date"]).astype("datetime64[ns]")
    new = new.copy(); new["Date"] = pd.to_datetime(new["Date"]).astype("datetime64[ns]")
    comb = pd.concat([old[old["Instrument"].isin(set(new["Instrument"])) == False], new],
                     ignore_index=True).sort_values(["Instrument", "Date"])
    comb.to_parquet(path, index=False)
    return old["Instrument"].nunique(), comb["Instrument"].nunique()

a, b = append(SCR + "/daily_marketcap_ev.parquet", mcev, ["Instrument", "Date"])
print(f"MC/EV panel: {a} -> {b} securities", flush=True)
a, b = append(SCR + "/daily_expected_return.parquet", er, ["Instrument", "Date"])
print(f"ER panel:    {a} -> {b} securities", flush=True)
print("DONE_EXTRA")
