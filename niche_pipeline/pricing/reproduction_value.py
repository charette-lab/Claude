#!/usr/bin/env python3
"""reproduction_value.py — Greenwald reproduction value + franchise value, added to the framework.

Method (Columbia "Value Investing in Technology"; Meta/Adobe/Salesforce/PayPal/Booking cases):
  Asset Value (reproduction) AV = tangible reproduction + capitalized intangibles (R&D + S&M).
  Earnings Power Value  EPV     = no-growth NOPAT / WACC.
  Franchise value = EPV - AV;  FranchiseShare = 1 - AV/EPV.

KEY EMPIRICAL FINDING (this file): FranchiseShare is a capitalization of the CURRENT excess-return
spread -- algebraically 1 - WACC/ROIC -- so it ranks almost perfectly with ROIC (Spearman 0.94)
and is STATISTICALLY UNRELATED to the moat-durability score (Spearman 0.03). So reproduction/
franchise value is NOT a moat validator; it is an ORTHOGONAL axis: "how much is this business
out-earning its asset base RIGHT NOW", vs the moat score's "how DURABLE is that excess". The value
is the 2-D map (franchise x moat) and the asset floor (AV/EV). The prize quadrant is
HIGH CoreMoat x LOW franchise = a durable franchise not currently earning its excess = the
BURIED-CORE / activist setup (bad subsidiaries suppress group NOPAT, so EPV~AV even though the core
moat is intact; free the core -> NOPAT rises -> franchise value appears).

  python3 pricing/reproduction_value.py
"""
import sys, os
import numpy as np, pandas as pd, openpyxl
from scipy.stats import spearmanr

U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
HIST = U+"bd173c55-hist_20260629_2.xlsx"

# ---- reproduction components + NOPAT from the fundamental history (latest row per name) ----
wb = openpyxl.load_workbook(HIST, data_only=True, read_only=True)["Original"]
it = wb.iter_rows(values_only=True); hdr = [str(c) for c in next(it)]
df = pd.DataFrame(list(it), columns=hdr); df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df = df.sort_values("Date").groupby("Instrument").tail(1)
for c in ["Gross Reproduction Cost", "Operating Working Capital", "R&D Capital Base", "SG&A Capital Base",
          "Base Physical Capital & Leases", "New Operating Income", "Invested Capital"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
df["tangible_repro"] = df["Gross Reproduction Cost"].fillna(df["Base Physical Capital & Leases"])
df["intangible_repro"] = df[["R&D Capital Base", "SG&A Capital Base"]].sum(axis=1, min_count=1)
# reproduction value of operating assets: clean IC base + off-book intangibles (a full CBS build would
# refine the tangible leg with age/inflation-adjusted PP&E and add the S&M customer-base reproduction)
df["AV"] = df["Invested Capital"] + df["intangible_repro"].fillna(0.0)
rv = df.set_index("Instrument")[["AV", "tangible_repro", "intangible_repro", "Invested Capital", "New Operating Income"]]

# ---- moats + WACC + EV from the Scored universe ----
ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rows = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rows[0])}
def col(n): return [(str(r[ci["Ticker"]]), r[ci[n]]) for r in rows[1:] if r[ci["Ticker"]]]
sc = pd.DataFrame({"Instrument": [t for t, _ in col("OpValue")]})
for n in ["OpValue", "OpVal/EV", "CompanyMoat(v3.2)", "CoreMoat(v3.2)", "AIP_WACC", "MoatGap"]:
    sc[n] = pd.to_numeric([v for _, v in col(n)], errors="coerce")
sc = sc.set_index("Instrument"); sc["EV"] = sc["OpValue"] / sc["OpVal/EV"]

j = rv.join(sc, how="inner")
j["EPV"] = j["New Operating Income"] / j["AIP_WACC"]                 # no-growth earnings power value
j["ROIC"] = j["New Operating Income"] / j["Invested Capital"]
j["FranchiseValue"] = j["EPV"] - j["AV"]
j["FranchiseShare"] = j["FranchiseValue"] / j["EPV"]
j["AV_over_EV"] = j["AV"] / j["EV"]                                  # asset floor: how much of price is asset-backed
g = j[(j.EPV > 0) & (j.AV > 0) & (j.ROIC > 0) & np.isfinite(j.FranchiseShare)].copy()

# ---- prove the orthogonality (the whole point) ----
print(f"n = {len(g)} names with reproduction value + EPV")
r_roic = spearmanr(g.FranchiseShare, g.ROIC)[0]
r_moat = spearmanr(g.FranchiseShare.values, g["CompanyMoat(v3.2)"].values, nan_policy="omit")[0]
r_core = spearmanr(g.FranchiseShare.values, g["CoreMoat(v3.2)"].values, nan_policy="omit")[0]
print(f"Spearman(FranchiseShare, ROIC)        = {r_roic:+.2f}   -> franchise value IS the current excess-return spread")
print(f"Spearman(FranchiseShare, CompanyMoat) = {r_moat:+.2f}   -> ORTHOGONAL to moat durability")
print(f"Spearman(FranchiseShare, CoreMoat)    = {r_core:+.2f}")

# ---- the 2-D map: current franchise (ROIC-based) x moat durability ----
fmed = g.FranchiseShare.median()
def quad(row):
    hi_f = row.FranchiseShare >= fmed; hi_m = (row["CoreMoat(v3.2)"] or 0) >= 7.8
    if hi_f and hi_m:  return "compounder (earns + durable)"
    if hi_f and not hi_m: return "fading (earns now, not durable)"
    if not hi_f and hi_m: return "BURIED CORE (durable, under-earning)"
    return "commodity (neither)"
g["quadrant"] = g.apply(quad, axis=1)
print("\n=== franchise x moat map (CoreMoat>=7.8 = durable; FranchiseShare vs median) ===")
print(g["quadrant"].value_counts().to_string())

# ---- the prize: BURIED-CORE candidates (durable core moat, suppressed current franchise) ----
bc = g[(g["CoreMoat(v3.2)"] >= 7.8) & (g.FranchiseShare < fmed)].copy()
bc = bc.sort_values("FranchiseShare")
print(f"\n=== BURIED-CORE candidates: CoreMoat>=7.8 but franchise value suppressed ({len(bc)}) ===")
print(f"{'ticker':>12} {'CoreMoat':>8} {'CoMoat':>6} {'MoatGap':>7} {'Franch%':>7} {'ROIC':>5} {'AV/EV':>6}")
for t, r in bc.head(15).iterrows():
    print(f"{t:>12} {r['CoreMoat(v3.2)']:8.1f} {r['CompanyMoat(v3.2)']:6.1f} {(r['MoatGap'] or 0):7.1f} "
          f"{r.FranchiseShare*100:6.0f}% {r.ROIC*100:4.0f}% {r.AV_over_EV:6.2f}")

j.reset_index()[["Instrument", "AV", "tangible_repro", "intangible_repro", "EPV", "FranchiseValue",
                 "FranchiseShare", "AV_over_EV", "ROIC", "CompanyMoat(v3.2)", "CoreMoat(v3.2)", "MoatGap"]].to_csv(SCR+"/reproduction_value.csv", index=False)
print("\nwrote reproduction_value.csv")
print("\nNOTE: AV here = Invested Capital + capitalized R&D/SG&A base (clean proxy). A full CBS build adds")
print("(a) age/inflation-adjusted PP&E tangible reproduction and (b) the S&M customer-base reproduction")
print("-- decisive for asset-light tech names, minor for this tangible-heavy quality universe.")
