#!/usr/bin/env python3
"""hidden_core.py — test the 'great core, wrecked by a management pivot' thesis.

Consolidated data can't isolate the core segment, but it CAN fingerprint the situation:
  CORE INTACT   : moat>=7.8  AND  gross margin >= 0.9x its 10yr norm  AND  long-run
                  ROIIC (21/14yr) high  -> the product/franchise economics are fine.
  GROUP WRECKED : EBITA margin <= 0.80x its 10yr norm  (destruction is BELOW gross,
                  i.e. in the operating/subsidiary layer)  OR recent ROIIC (3yr) has
                  collapsed >=5pts below the through-cycle core while reinvestment is high.
A candidate = CORE INTACT and GROUP WRECKED.

STAGE 1 (no prices): does the core REASSERT? For candidates in year T, measure how much
of the EBITA-margin gap-to-normal is recovered by T+3 / T+5, and compare HIGH-moat
candidates against LOW-moat companies with the SAME depression (control). If the moat
names mean-revert up and the junk names don't, the thesis mechanism holds.

  AIP_VALUE_ENGINE=.../roiic_dcf.py python3 pricing/hidden_core.py
"""
import sys, warnings
import numpy as np, pandas as pd, openpyxl
warnings.filterwarnings("ignore")
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
num = lambda v: v if isinstance(v, (int, float)) else None

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
def _col(name):
    return {str(r[ci["Ticker"]]): r[ci[name]] for r in rws[1:]
            if r[ci["Ticker"]] and name in ci and isinstance(r[ci[name]], (int, float))}
coremoat = _col("CoreMoat(v3.2)")        # the CORE franchise moat (the thesis driver)
companymoat = _col("CompanyMoat(v3.2)")  # the consolidated GROUP moat (dragged down by subs)
moatgap = _col("MoatGap")                # CoreMoat - CompanyMoat: how buried the core is

print("loading deep panel...", flush=True)
by, idx = panel30.load(XLSB)

def G(r, *names):
    for k in names:
        if k in idx and idx[k] < len(r):
            v = r[idx[k]]
            if isinstance(v, (int, float)):
                return v
    return None

# build per-company yearly metric series
COMP = {}
for t, rs in by.items():
    t = str(t); ser = {}
    for r in rs:
        d = pd.to_datetime(G(r, "Period_End_Date") or (r[idx["Date"]] if "Date" in idx else None), errors="coerce")
        if pd.isna(d) or d.year < 2000:
            continue
        ser[d.year] = dict(
            em=G(r, "EBITA_Margin"), emn=G(r, "EBITA_Average_Margin 10 yr", "EBITA_Average_Margin_Last_10_yr"),
            gm=G(r, "Gross Profit Margin", "Gross_Profit_Margin"), gmn=G(r, "Gross_Average_Margin_Last_10_yr"),
            r3=G(r, "ROICm_total - 3 years", "ROICm_total___3_years"),
            r21=G(r, "ROICm_total - 21 years", "ROICm_total___21_years", "ROICm_total - 14 years", "ROICm_total___14_years"),
            rr=G(r, "average_C - 7 year", "average_C___7_year"))
    if len(ser) >= 6:
        COMP[t] = ser

# normalise ROIIC/margins that may be in percent
def pct(x):
    if x is None: return None
    return x/100.0 if abs(x) > 1.5 else x

MOAT_HI, GATE_LONG_ROIC, DEP = 7.8, 0.10, 0.80
cand, ctrl = [], []          # (ticker, year) tuples
records = []
for t, ser in COMP.items():
    cm = coremoat.get(t); com = companymoat.get(t); mg = moatgap.get(t)
    for y, f in ser.items():
        em, emn = pct(f["em"]), pct(f["emn"]); gm, gmn = pct(f["gm"]), pct(f["gmn"])
        r3, r21 = pct(f["r3"]), pct(f["r21"])
        if em is None or emn is None or emn <= 0:
            continue
        group_wrecked = (em <= DEP*emn) or (r3 is not None and r21 is not None and r3 <= r21-0.05)
        gross_ok = (gm is None or gmn is None) or (gm >= 0.9*gmn)
        core_strong = (cm is not None and cm >= MOAT_HI) and (r21 is not None and r21 >= GATE_LONG_ROIC) and gross_ok
        if not group_wrecked:
            continue
        # forward recovery of the EBITA-margin gap-to-normal
        gap0 = emn - em
        rec = {}
        for H in (3, 5):
            g2 = ser.get(y+H)
            if g2 and pct(g2["em"]) is not None and gap0 > 1e-4:
                rec[H] = (pct(g2["em"]) - em)/gap0        # 1.0 = fully back to 10yr norm
        row = dict(t=t, y=y, coremoat=cm, companymoat=com, moatgap=mg, em=em, emn=emn,
                   gap=gap0, r3=r3, r21=r21, **{f"rec{H}": rec.get(H) for H in (3, 5)})
        if core_strong:
            cand.append(row)
        elif cm is not None and cm < 6.5:                  # control: same depression, WEAK core
            ctrl.append(row)
        records.append(row)

C = pd.DataFrame(cand); K = pd.DataFrame(ctrl)
print(f"\nCANDIDATES (CoreMoat>=7.8, long-run ROIIC>=10%, gross intact, group wrecked): {len(C)} company-years, {C.t.nunique()} names")
print(f"CONTROL   (CoreMoat<6.5, same group depression):                            {len(K)} company-years, {K.t.nunique()} names")
print(f"\ncandidates per year (is it a real, populated strategy?):")
cpy = C.groupby("y").t.nunique()
print("  " + "  ".join(f"{y}:{n}" for y, n in cpy.items()))

def med(df, col):
    s = df[col].dropna()
    return (s.median()*100, len(s))

print("\n=== STAGE 1: does the core REASSERT? (share of the EBITA-margin gap recovered) ===")
print(f"  {'':28s} {'T+3yr':>14} {'T+5yr':>14}")
for name, df in [("HIGH-moat candidates", C), ("LOW-moat control (same drop)", K)]:
    m3, n3 = med(df, "rec3"); m5, n5 = med(df, "rec5")
    print(f"  {name:28s}  {m3:6.0f}% (n={n3:4d}) {m5:6.0f}% (n={n5:4d})")
print("\n  read: 100% = margin fully recovered to its 10-yr norm; 0% = no recovery; <0 = got worse.")
print("  If HIGH-moat >> LOW-moat, the moat lets the buried core reassert -> thesis mechanism supported.")
# also: ROIIC recovery
for H in (3, 5):
    ch = C[f"rec{H}"].dropna(); kh = K[f"rec{H}"].dropna()
    if len(ch) and len(kh):
        from scipy.stats import mannwhitneyu
        try:
            p = mannwhitneyu(ch, kh, alternative="greater").pvalue
            print(f"  T+{H}: high-moat recovery > control, Mann-Whitney p = {p:.3f}")
        except Exception:
            pass
# DIRECT thesis test: does a bigger CoreMoat-vs-group gap predict more group recovery?
R = pd.DataFrame(records).dropna(subset=["moatgap"])
print("\n=== DIRECT TEST: does the MoatGap (core minus group) predict the group closing it? ===")
try:
    R["gap_tier"] = pd.qcut(R["moatgap"], 3, labels=["small gap", "mid gap", "LARGE gap (buried core)"])
    for tier in ["small gap", "mid gap", "LARGE gap (buried core)"]:
        s = R[R.gap_tier == tier]
        m5 = s["rec5"].dropna()
        print(f"  {tier:26s} n={len(s):5d}  median MoatGap {s.moatgap.median():.2f}  "
              f"margin recovery T+5 = {m5.median()*100:4.0f}% (n={len(m5)})")
    print("  -> if LARGE-gap names recover more, the market's group-based pricing understates a core that reasserts.")
except Exception as e:
    print("  (tercile split failed:", e, ")")
pd.DataFrame(records).to_parquet(SCR+"/hidden_core_candidates.parquet", index=False)
print("\nwrote hidden_core_candidates.parquet (candidate list for the Stage-2 return test once prices land)")
