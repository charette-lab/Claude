#!/usr/bin/env python3
"""reproduction_tech.py — Greenwald reproduction value + EPV for TECH (Software & Services).

Follows the Columbia "Value Investing in Technology" method (Greenwald), built from the panel's
existing intangible capitalization. Validated against the course decks: this reproduces Adobe's
AV ~$37bn and EPV ~$89bn almost exactly from our columns.

  Asset Value of equity (reproduction):
      AV = Book Equity + Product portfolio (R&D Capital Base) + Customer/brand portfolio (SG&A Capital Base)
      [ − acquisition goodwill: TODO once a goodwill-balance column is available ]

  Earnings Power Value of equity (no-growth):
      Adjusted operating income = EBITA + growth R&D + growth S&M   (add back the GROWTH portion of
          intangible spend; expense only maintenance = decay-rate x capital base)
      Adjusted NOPAT = Adj OI x (1 − tax)
      EPV = Adjusted NOPAT / WACC + Cash − Debt

  Franchise value = EPV − AV  (>0 ⇒ barriers to entry). Three-tier read: Market cap vs EPV vs AV.
  Moat confirmation: marginal ROIC (panel ROICm) vs WACC.  Cheapness: Market cap vs EPV (growth free
  when price ≤ EPV, the Booking/Meta signal).

  python3 pricing/reproduction_tech.py
"""
import sys, os
import numpy as np, pandas as pd, openpyxl

U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
HIST = U+"bd173c55-hist_20260629_2.xlsx"
DEFAULT_WACC = 0.09

C = {  # panel column -> short name
    "R&D Capital Base": "P_rd", "SG&A Capital Base": "P_sga",
    "R&D Decay Rate": "d_rd", "SG&A Decay Rate": "d_sga",
    "R&D Growth": "g_rd", "SG&A Growth": "g_sga",
    "R&D Expense - Expensed & Capitalized - Total - Suppl": "rd_exp",
    "EBITA": "ebita", "Operating Income": "oi", "Cash & Short Term Investments": "cash",
    "Income Tax Rate - Instrument": "tax", "Shareholders' Equity - Attributable to Parent ShHold - Total": "bookeq",
    "Market Capitalization": "mktcap", "Sales": "sales",
    "Debt - Long-Term - Total": "ltd", "Short-Term Debt & Current Portion of Long-Term Debt": "std",
    "Capitalized Lease Obligations - Long-Term": "ll", "Capitalized Leases - Current Portion": "lc",
    "ROICm_total - 7 years": "roicm7",
    "Operating Working Capital": "owc", "Gross Reproduction Cost": "ppe_repro",
    "Base Physical Capital & Leases": "ppe_base",
}

wb = openpyxl.load_workbook(HIST, data_only=True, read_only=True)["Original"]
it = wb.iter_rows(values_only=True); hdr = [str(c) for c in next(it)]
df = pd.DataFrame(list(it), columns=hdr); df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df = df.sort_values("Date").groupby("Instrument").tail(1)
df = df[df["GICS Industry Group Name"].astype(str) == "Software & Services"].copy()
for c in C:
    df[C[c]] = pd.to_numeric(df[c], errors="coerce")
d = df.set_index("Instrument")

# WACC + moats from the Scored universe
ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rows = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rows[0])}
def col(n): return {str(r[ci["Ticker"]]): r[ci[n]] for r in rows[1:] if r[ci["Ticker"]]}
wacc = {k: pd.to_numeric(pd.Series([v]), errors="coerce")[0] for k, v in col("AIP_WACC").items()}
cmoat, coremoat = col("CompanyMoat(v3.2)"), col("CoreMoat(v3.2)")

def num(x, default=np.nan):
    try:
        v = float(x); return v if np.isfinite(v) else default
    except Exception:
        return default

out = []
for t, r in d.iterrows():
    P = num(r.P_rd, 0) + num(r.P_sga, 0)                       # capitalized intangibles (product + customer/brand)
    w = num(wacc.get(t), DEFAULT_WACC) or DEFAULT_WACC
    tax = num(r.tax, 0.25);  tax = tax if 0 <= tax < 0.6 else 0.25
    adj_oi = num(r.ebita, np.nan) + num(r.g_rd, 0) + num(r.g_sga, 0)   # add back growth intangible spend
    adj_nopat = adj_oi * (1 - tax)
    debt = num(r.ltd, 0) + num(r.std, 0) + num(r.ll, 0) + num(r.lc, 0)
    epv = adj_nopat / w + num(r.cash, 0) - debt                # EPV of equity
    mc = num(r.mktcap, np.nan)
    # AV of equity, two ways:
    av_kept = num(r.bookeq, np.nan) + P                        # book equity (goodwill INCLUDED) + intangibles -- upper bound
    tang = num(r.ppe_repro, np.nan)
    if not np.isfinite(tang):
        tang = num(r.ppe_base, 0)
    av_repro = num(r.cash, 0) + num(r.owc, 0) + tang + P - debt  # bottom-up reproduction, NO goodwill -- lower bound
    out.append(dict(Instrument=t, Sales=num(r.sales), AV_kept=av_kept, AV_repro=av_repro, EPV=epv, MktCap=mc,
                    Fr_kept=epv-av_kept, Fr_repro=epv-av_repro,
                    EPVoverAV_kept=epv/av_kept if av_kept > 0 else np.nan,
                    EPVoverAV_repro=epv/av_repro if av_repro > 0 else np.nan,
                    MktCap_over_EPV=mc/epv if epv > 0 else np.nan,
                    AVrepro_over_MktCap=av_repro/mc if mc > 0 else np.nan,
                    ROICm7=num(r.roicm7), WACC=w, CoreMoat=num(coremoat.get(t)), CompanyMoat=num(cmoat.get(t))))
R = pd.DataFrame(out).set_index("Instrument")

# ---- goodwill kept vs removed: the acquirers where it matters most ----
print("=== goodwill KEPT (book equity) vs REMOVED (bottom-up reproduction) — worked names ===")
print(f"{'ticker':>9} {'AV_kept':>8} {'AV_repro':>8} {'EPV':>6} {'Fr_kept':>8} {'Fr_repro':>8} {'EPV/AV k>r':>12}")
for s in ["ADBE.OQ", "CRM.N", "SAPG.DE", "ORCL.N", "MSFT.OQ", "INTU.OQ"]:
    if s in R.index:
        x = R.loc[s]
        print(f"{s:>9} {x.AV_kept/1e9:7.0f}b {x.AV_repro/1e9:7.0f}b {x.EPV/1e9:5.0f}b {x.Fr_kept/1e9:7.0f}b "
              f"{x.Fr_repro/1e9:7.0f}b   {x.EPVoverAV_kept:4.1f}x>{x.EPVoverAV_repro:4.1f}x")

# ---- book-level shift ----
v = R[(R.EPV > 0) & (R.AV_kept > 0) & (R.AV_repro > 0) & R.MktCap.notna()].copy()
print(f"\ntech names valued: {len(v)}")
print(f"EPV>AV (franchise present):   goodwill-kept {100*(v.EPV>v.AV_kept).mean():.0f}%  ->  goodwill-removed {100*(v.EPV>v.AV_repro).mean():.0f}%")
print(f"median EPV/AV:                goodwill-kept {v.EPVoverAV_kept.median():.2f}x  ->  goodwill-removed {v.EPVoverAV_repro.median():.2f}x")
print(f"median goodwill drag (AV_kept/AV_repro): {(v.AV_kept/v.AV_repro).median():.2f}x  "
      f"(how much book-equity AV overstates reproduction)")

# ---- cheapness screen on the goodwill-free reproduction ----
cheap = v[(v.MktCap_over_EPV <= 1.0) & (v.EPV > v.AV_repro) & (v.CompanyMoat >= 6.5)].sort_values("MktCap_over_EPV")
print(f"\n=== CHEAP: price<=EPV, franchise (goodwill-free), moat>=6.5 — {len(cheap)} names ===")
print(f"{'ticker':>11} {'MC/EPV':>6} {'EPV/AVr':>7} {'ROICm':>6} {'WACC':>5} {'AVr/MC':>6} {'Moat':>4}")
for t, x in cheap.head(15).iterrows():
    print(f"{t:>11} {x.MktCap_over_EPV:6.2f} {x.EPVoverAV_repro:7.1f} {x.ROICm7*100 if np.isfinite(x.ROICm7) else 0:5.0f}% "
          f"{x.WACC*100:4.0f}% {x.AVrepro_over_MktCap:6.2f} {x.CompanyMoat if np.isfinite(x.CompanyMoat) else 0:4.1f}")

R.reset_index().to_csv(SCR+"/reproduction_tech.csv", index=False)
print("\nwrote reproduction_tech.csv")
print("AV_kept = book equity (goodwill IN) + intangibles [upper bound]; AV_repro = cash + working capital")
print("+ PP&E reproduction + R&D base + SG&A base - debt [bottom-up, NO goodwill, lower bound]. Deck-style")
print("(keep saleable-product goodwill) sits between; for organic builders the two nearly coincide.")
