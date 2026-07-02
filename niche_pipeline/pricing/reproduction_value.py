#!/usr/bin/env python3
"""reproduction_value.py — Greenwald asset-reproduction value + franchise value, added to the framework.

Reproduction Value of operating assets (what a competitor would pay to rebuild the business):
    RV = Gross Reproduction Cost (inflation/age-adjusted PP&E)
       + Operating Working Capital
       + R&D Capital Base        (capitalized R&D, via the panel's R&D decay rate)
       + SG&A Capital Base       (capitalized customer/brand build, via SG&A cap+decay rate)
The panel already carries every component (it uses them for the ROIIC growth split); we just
ASSEMBLE them into a standalone reproduction value and pair it with the engine's operating value.

Franchise (moat) value = EPV/operating value - reproduction value:
    FranchiseValue = OpValue(engine) - RV
    FranchiseShare = FranchiseValue / OpValue        (>0 => a real competitive advantage exists;
                                                       ~0 => commodity; <0 => assets worth more than
                                                       the earnings stream, i.e. value destruction)
This is an INDEPENDENT, accounting-based moat gauge to cross-check CompanyMoat/CoreMoat and the
buried-core thesis, and RV/EV gives an asset floor under the price.

  python3 pricing/reproduction_value.py
"""
import sys, os
import numpy as np, pandas as pd, openpyxl

U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
HIST = U+"bd173c55-hist_20260629_2.xlsx"

# ---- reproduction components from the fundamental history (latest row per name) ----
wb = openpyxl.load_workbook(HIST, data_only=True, read_only=True)["Original"]
it = wb.iter_rows(values_only=True); hdr = [str(c) for c in next(it)]
df = pd.DataFrame(list(it), columns=hdr)
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df = df.sort_values("Date").groupby("Instrument").tail(1)
num = ["Gross Reproduction Cost", "Operating Working Capital", "R&D Capital Base", "SG&A Capital Base",
       "Invested Capital", "Cash & Short Term Investments", "Sales", "Property Plant & Equipment - Gross - Total"]
for c in num:
    df[c] = pd.to_numeric(df[c], errors="coerce")
df["RV_tangible"] = df[["Gross Reproduction Cost", "Operating Working Capital"]].sum(axis=1, min_count=1)
df["RV_intangible"] = df[["R&D Capital Base", "SG&A Capital Base"]].sum(axis=1, min_count=1)
df["RV"] = df[["Gross Reproduction Cost", "Operating Working Capital", "R&D Capital Base", "SG&A Capital Base"]].sum(axis=1, min_count=1)
rv = df.set_index("Instrument")[["RV", "RV_tangible", "RV_intangible", "R&D Capital Base", "SG&A Capital Base",
                                 "Invested Capital", "Sales"]]

# ---- engine operating value + moats from the Scored universe ----
ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rows = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rows[0])}
def col(name):
    return [(str(r[ci["Ticker"]]), r[ci[name]]) for r in rows[1:] if r[ci["Ticker"]]]
sc = pd.DataFrame({"Instrument": [t for t, _ in col("OpValue")]})
for c in ["OpValue", "CompanyMoat(v3.2)", "CoreMoat(v3.2)", "OpVal/EV", "ER@12%(as-is)", "MoatGap"]:
    sc[c] = pd.to_numeric([v for _, v in col(c)], errors="coerce")
sc = sc.set_index("Instrument")
sc["EV"] = sc["OpValue"] / sc["OpVal/EV"]     # OpVal/EV = engine operating value / enterprise value

j = rv.join(sc, how="inner").dropna(subset=["RV", "OpValue"])
j = j[(j["OpValue"] > 0)]
j["FranchiseValue"] = j["OpValue"] - j["RV"]
j["FranchiseShare"] = j["FranchiseValue"] / j["OpValue"]
j["RV_over_EV"] = j["RV"] / j["EV"]
j["IntangibleShareRV"] = j["RV_intangible"] / j["RV"]
print(f"names with both reproduction value and engine OpValue: {len(j)}\n")

# ---- does the accounting franchise value agree with the moat score? (cross-check) ----
print("=== Franchise Share (moat value / operating value) by CompanyMoat band ===")
j["moatband"] = pd.cut(j["CompanyMoat(v3.2)"], [0, 5, 6.5, 7.8, 8.5, 11],
                       labels=["<5", "5-6.5", "6.5-7.8", "7.8-8.5", "8.5+"])
g = j.groupby("moatband", observed=True).agg(n=("FranchiseShare", "size"),
        FranchiseShare=("FranchiseShare", "median"), RV_over_EV=("RV_over_EV", "median"),
        IntangShareRV=("IntangibleShareRV", "median"))
print(g.round(2).to_string())
corr = j[["CompanyMoat(v3.2)", "FranchiseShare"]].corr().iloc[0, 1]
print(f"\ncorrelation(CompanyMoat, FranchiseShare) = {corr:.2f}   (positive => accounting franchise value validates the moat score)")

# ---- asset-floor view + the moat/no-moat/value-trap buckets ----
def bucket(fs):
    return "franchise (moat)" if fs > 0.25 else ("thin/none" if fs > -0.10 else "value-trap (assets>EPV)")
j["kind"] = j["FranchiseShare"].apply(bucket)
print("\n=== value composition buckets ===")
print(j["kind"].value_counts().to_string())

# ---- widest franchises (highest moat value share, engine agrees) ----
print("\n=== top 12 by Franchise Share (biggest moats on the accounting cross-check) ===")
top = j.sort_values("FranchiseShare", ascending=False).head(12)
print(f"{'ticker':>12} {'Moat':>4} {'Frnch%':>6} {'RV/EV':>6} {'IntgRV%':>7}  {'RV($b)':>7} {'OpVal($b)':>9}")
for t, r in top.iterrows():
    print(f"{t:>12} {r['CompanyMoat(v3.2)']:4.1f} {r['FranchiseShare']*100:5.0f}% {r['RV_over_EV']:6.2f} {r['IntangibleShareRV']*100:6.0f}% {r['RV']/1e9:7.2f} {r['OpValue']/1e9:9.2f}")

j.reset_index()[["Instrument", "RV", "RV_tangible", "RV_intangible", "OpValue", "FranchiseValue",
                 "FranchiseShare", "RV_over_EV", "IntangibleShareRV", "CompanyMoat(v3.2)", "CoreMoat(v3.2)",
                 "MoatGap"]].to_csv(SCR+"/reproduction_value.csv", index=False)
print("\nwrote reproduction_value.csv")
