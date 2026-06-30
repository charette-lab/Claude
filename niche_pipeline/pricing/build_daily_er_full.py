#!/usr/bin/env python3
"""build_daily_er_full.py — daily expected return from the FULL engine.

Augments the raw fade ER (daily_expected_return.parquet) with the two extra layers
the complete pipeline applies:

  SUPPLY+DEMAND NORMALIZATION (overearning.two_stage_return): a transient scarcity
    rent (capital-cycle / supply-demand imbalance) is faded down to the reproduction
    equilibrium, and a genAI-compressed software moat fades faster. Exactly as the
    pipeline: a per-(security,fiscal-year) downgrade dn = er_current - er_adj, applied
    flat across that year's daily prices.  Point-in-time: signals use only history up
    to each fiscal year-end (no look-ahead).

  UNREASONABLE-RETURN SCREEN (frameworks.er_is_artifact): a quality-aware ceiling on
    the raw ER (roic_star from the same signals), flagging extrapolation artifacts
    (e.g. an AI-boom NOPAT giving a triple-digit IRR) for exclusion from the book.

Output: daily_expected_return_full.parquet
  [Instrument, Date, market_cap, er_current, expected_return(=er_adj), artifact]

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/build_daily_er_full.py
"""
import sys, os
import numpy as np, pandas as pd, openpyxl
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30, aip, overearning, history, frameworks as F

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
HIST2 = U+"29a7c3ae-hist_20260629_2.xlsx"

# moats / industry / country
ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
moat = {str(r[ci["Ticker"]]): r[ci["CompanyMoat(v3.2)"]] for r in rws[1:]
        if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}

num = lambda v: v if isinstance(v, (int, float)) else None


def fin_from_row(t, r, g, m):
    nopat = num(g(r, "New Operating Income")); roiic = num(g(r, "ROICm_total - 7 years"))
    rr = num(g(r, "average_C - 7 year")); mc = num(g(r, "Market Capitalization"))
    ev = num(g(r, "EV")); sales = num(g(r, "Sales"))
    tax = num(g(r, "Income Tax Rate - Instrument"))
    ltd = num(g(r, "Debt - Long-Term - Total")) or 0.0
    std = num(g(r, "Short-Term Debt & Current Portion of Long-Term Debt")) or 0.0
    if None in (nopat, roiic, rr) or not mc or ev is None:
        return None, None
    ped = g(r, "Period_End_Date") or g(r, "Date")
    fin = {"Company Name": t, "Instrument": t,
           "GICS Industry Group Name": g(r, "GICS Industry Group Name"),
           "Country of Headquarters": g(r, "Country of Headquarters"),
           "New Operating Income": nopat, "ROICm 7": roiic, "RR 7": rr, "Moat Score": m,
           "Gross debt": ltd+std, "Income Tax Rate - Instrument": (tax/100.0 if tax is not None else 0.25),
           "Net debt": ev-mc, "Sales": sales, "Market Cap": mc}
    return ped, fin


def process(by, idx, label, recs):
    g = lambda r, k: (r[idx[k]] if k in idx and idx[k] < len(r) else None)
    n = 0
    for t, rows in by.items():
        t = str(t); m = moat.get(t)
        if m is None:
            continue
        for i in range(len(rows)):
            ped, fin = fin_from_row(t, rows[i], g, m)
            if fin is None or ped is None:
                continue
            mvh = None
            try:
                sig = overearning.panel_signals(rows[:i+1], idx)        # point-in-time
                ts = overearning.two_stage_return(fin, sig, re=0.07, re2=0.12)
                mvh = history.verdict(history.summarize(rows[:i+1], idx), m)[0]
            except Exception:
                ts = None
            if not ts:
                erc = aip.value_and_return(fin, re=0.07, re2=0.12)
                erc = erc.get("er1") if erc else None
                dn = 0.0; rstar = None
            else:
                erc, era = ts.get("er_current"), ts.get("er_adj")
                dn = max(0.0, (erc - era)) if (erc is not None and era is not None) else 0.0
                rstar = sig.get("roic_star")
            # quality-aware artifact screen (100% ceiling for history-CONFIRMED high-ROIC* names)
            artifact = bool(F.er_is_artifact(erc, rstar, mvh)) if erc is not None else False
            recs.append({"Instrument": t, "ped": ped, "dn": dn, "artifact": artifact})
            n += 1
        if n and n % 5000 < 1:
            print(f"  [{label}] ~{n}", flush=True)
    print(f"  [{label}] done: {n} (security,year) rows", flush=True)


recs = []
print("loading main panel...", flush=True)
by, idx = panel30.load(XLSB)
process(by, idx, "main", recs)

print("loading hist part-file...", flush=True)
h = pd.read_excel(HIST2)
h["Instrument"] = h["Instrument"].astype(str)
hcols = list(h.columns); hidx = {c: i for i, c in enumerate(hcols)}
hidx["Period_End_Date"] = hidx.get("Date")
by_x = {t: [list(r) for r in g.sort_values("Date").itertuples(index=False, name=None)]
        for t, g in h.groupby("Instrument")}
process(by_x, hidx, "extra", recs)

DN = pd.DataFrame(recs)
DN["ped"] = pd.to_datetime(DN["ped"], errors="coerce").astype("datetime64[ns]")
DN = DN.dropna(subset=["ped"]).sort_values("ped")
print(f"downgrade rows: {len(DN)} | names adjusted (dn>0): {DN.Instrument[DN.dn>1e-4].nunique()} "
      f"| artifact rows: {int(DN.artifact.sum())}", flush=True)

# attach to daily er_current
er = pd.read_parquet(SCR+"/daily_expected_return.parquet")
er["Instrument"] = er["Instrument"].astype(str)
er["Date"] = pd.to_datetime(er["Date"]).astype("datetime64[ns]")
er = er.rename(columns={"expected_return": "er_current"}).sort_values("Date")
m = pd.merge_asof(er, DN.sort_values("ped"), left_on="Date", right_on="ped",
                  by="Instrument", direction="backward")
m["dn"] = m["dn"].fillna(0.0); m["artifact"] = m["artifact"].fillna(False)
m["expected_return"] = m["er_current"] - m["dn"]
out = m[["Instrument", "Date", "market_cap", "er_current", "expected_return", "artifact"]]
out.to_parquet(SCR+"/daily_expected_return_full.parquet", index=False)
print(f"WROTE {len(out)} rows / {out.Instrument.nunique()} securities -> daily_expected_return_full.parquet")
d = (out["er_current"] - out["expected_return"])
print(f"mean daily downgrade: {d.mean()*100:.2f}pp | rows downgraded: {(d>1e-4).mean()*100:.1f}% | "
      f"artifact rows: {out.artifact.mean()*100:.1f}%")
print("DONE_FULL")
