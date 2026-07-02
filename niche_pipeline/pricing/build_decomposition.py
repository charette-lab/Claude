#!/usr/bin/env python3
"""build_decomposition.py — persist the DAILY full-engine return decomposition.

For every (security, day): er_total (full-engine er1: 5yr price->intrinsic, terminal
included, supply/demand-normalized), carry (internal-growth component = engine
er1_carry, exit at entry EV), and rerate (= er_total - carry, the revaluation).
Built once so downstream selection never recomputes the engine.

Method mirrors the existing panels exactly:
  Pass 1 (per security, fiscal-year): point-in-time signals -> two_stage_return to
    get the supply/demand carry downgrade dn_carry (raw er1_carry - adjusted carry)
    and the artifact flag.
  Pass 2 (per security, fiscal-year): er1_carry is monotonic in price, so sample it
    at 12 market-cap points and np.interp to the daily grid; carry = raw - dn_carry.
  er_total + artifact come from daily_expected_return_full (the validated full er1);
  rerate = er_total - carry.

Output: daily_return_decomposition.parquet
  [Instrument, Date, market_cap, er_total, carry, rerate, artifact]

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/build_decomposition.py
"""
import sys, os
import numpy as np, pandas as pd, openpyxl
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30, aip, overearning

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
HIST2 = U+"29a7c3ae-hist_20260629_2.xlsx"
NP = 12
num = lambda v: v if isinstance(v, (int, float)) else None

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
moat = {str(r[ci["Ticker"]]): r[ci["CompanyMoat(v3.2)"]] for r in rws[1:]
        if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}
ind = {str(r[ci["Ticker"]]): r[ci["Industry"]] for r in rws[1:] if r[ci["Ticker"]]}
ctry = {str(r[ci["Ticker"]]): r[ci["Country"]] for r in rws[1:] if r[ci["Ticker"]]}


def finrow(t, r, g, m):
    nopat = num(g(r, "New Operating Income")); roiic = num(g(r, "ROICm_total - 7 years"))
    rr = num(g(r, "average_C - 7 year")); mc = num(g(r, "Market Capitalization")); ev = num(g(r, "EV"))
    if None in (nopat, roiic, rr) or not mc or ev is None:
        return None
    ltd = num(g(r, "Debt - Long-Term - Total")) or 0.0
    std = num(g(r, "Short-Term Debt & Current Portion of Long-Term Debt")) or 0.0
    tax = num(g(r, "Income Tax Rate - Instrument"))
    return (g(r, "Period_End_Date") or g(r, "Date")), dict(
        nopat=nopat, roiic=roiic, rr=rr, sales=num(g(r, "Sales")),
        tax=(tax/100.0 if tax is not None else 0.25), gross=ltd+std, netdebt=ev-mc, mc=mc)


def base_fin(t, f, m):
    return {"Company Name": t, "Instrument": t, "GICS Industry Group Name": ind.get(t),
            "Country of Headquarters": ctry.get(t), "New Operating Income": f["nopat"],
            "ROICm 7": f["roiic"], "RR 7": f["rr"], "Moat Score": m, "Gross debt": f["gross"],
            "Income Tax Rate - Instrument": f["tax"], "Net debt": f["netdebt"], "Sales": f["sales"]}


# ---- Pass 1: per (security, fiscal-year) carry downgrade + fundamentals ----
dn = {}   # (inst, ped) -> dn_carry
FIN = []  # rows for the daily merge
def pass1(by, idx, label):
    g = lambda r, k: (r[idx[k]] if k in idx and idx[k] < len(r) else None)
    n = 0
    for t, rows in by.items():
        t = str(t); m = moat.get(t)
        if m is None:
            continue
        for i in range(len(rows)):
            fr = finrow(t, rows[i], g, m)
            if not fr:
                continue
            ped, f = fr
            fin = base_fin(t, f, m); fin["Market Cap"] = f["mc"]
            try:
                sig = overearning.panel_signals(rows[:i+1], idx)
                ts = overearning.two_stage_return(fin, sig, re=0.07, re2=0.12)
                stated = aip.value_and_return(fin, re=0.07, re2=None)
                rc = stated.get("er1_carry") if stated else None
                ac = ts.get("er_carry") if ts else None
                d = max(0.0, rc - ac) if (rc is not None and ac is not None) else 0.0
            except Exception:
                d = 0.0
            dn[(t, pd.Timestamp(ped))] = d
            row = dict(Instrument=t, ped=ped, **f); row["moat"] = m
            FIN.append(row); n += 1
    print(f"  [pass1 {label}] {n} (security,year)", flush=True)


print("loading panels...", flush=True)
by, idx = panel30.load(XLSB)
pass1(by, idx, "main")
h = pd.read_excel(HIST2); h["Instrument"] = h["Instrument"].astype(str)
hidx = {c: i for i, c in enumerate(h.columns)}; hidx["Period_End_Date"] = hidx.get("Date")
by_x = {t: [list(r) for r in g.sort_values("Date").itertuples(index=False, name=None)]
        for t, g in h.groupby("Instrument")}
pass1(by_x, hidx, "extra")

F = pd.DataFrame(FIN)
F["ped"] = pd.to_datetime(F["ped"], errors="coerce").astype("datetime64[ns]")
F = F.dropna(subset=["ped"]).sort_values("ped")
print(f"fins: {len(F)} / {F.Instrument.nunique()} securities", flush=True)

# ---- Pass 2: daily carry via 12-point interpolation of er1_carry ----
daily = pd.read_parquet(SCR+"/daily_marketcap_ev.parquet", columns=["Instrument", "Date", "market_cap"])
daily["Instrument"] = daily["Instrument"].astype(str); daily["Date"] = pd.to_datetime(daily["Date"]).astype("datetime64[ns]")
daily = daily.sort_values("Date")
m = pd.merge_asof(daily, F.sort_values("ped"), left_on="Date", right_on="ped", by="Instrument", direction="backward")
m = m.dropna(subset=["ped"])

def carry_at(base, mc):
    f = dict(base); f["Market Cap"] = mc
    v = aip.value_and_return(f, re=0.07, re2=None)
    return v.get("er1_carry") if v else None

out = []; done = 0
for (t, ped), grp in m.groupby(["Instrument", "ped"]):
    r0 = grp.iloc[0]; base = base_fin(t, r0, r0["moat"])
    mcs = grp["market_cap"].values; lo, hi = np.nanmin(mcs), np.nanmax(mcs)
    if not (lo > 0):
        continue
    gx = np.unique(np.linspace(lo, hi, NP)) if hi > lo else np.array([lo])
    gd = [(x, carry_at(base, x)) for x in gx]
    gd = [(x, c) for x, c in gd if c is not None]
    if not gd:
        continue
    xs = np.array([x for x, _ in gd]); cs = np.array([c for _, c in gd])
    craw = np.interp(mcs, xs, cs)
    out.append(pd.DataFrame({"Instrument": t, "Date": grp["Date"].values,
                             "carry": craw - dn.get((t, pd.Timestamp(ped)), 0.0)}))
    done += 1
    if done % 5000 == 0:
        print(f"  [pass2] {done} groups", flush=True)
C = pd.concat(out, ignore_index=True)
print(f"daily carry rows: {len(C)} / {C.Instrument.nunique()} securities", flush=True)

# ---- assemble: er_total + artifact from the full panel; rerate = er_total - carry ----
full = pd.read_parquet(SCR+"/daily_expected_return_full.parquet")[["Instrument", "Date", "market_cap", "expected_return", "artifact"]]
full["Instrument"] = full["Instrument"].astype(str); full["Date"] = pd.to_datetime(full["Date"]).astype("datetime64[ns]")
full = full.rename(columns={"expected_return": "er_total"})
dec = full.merge(C, on=["Instrument", "Date"], how="left")
dec["rerate"] = dec["er_total"] - dec["carry"]
dec = dec[["Instrument", "Date", "market_cap", "er_total", "carry", "rerate", "artifact"]].sort_values(["Instrument", "Date"])
dec.to_parquet(SCR+"/daily_return_decomposition.parquet", index=False)
cov = dec["carry"].notna().mean()
print(f"WROTE {len(dec)} rows / {dec.Instrument.nunique()} securities -> daily_return_decomposition.parquet")
print(f"carry coverage: {cov*100:.1f}% | mean er_total {dec.er_total.mean()*100:.1f}%  "
      f"mean carry {dec.carry.mean()*100:.1f}%  mean rerate {dec.rerate.mean()*100:.1f}%")
print("DONE_DECOMP")
