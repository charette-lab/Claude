#!/usr/bin/env python3
"""nopat_growth.py — how fast does New Operating Income actually compound?

Tests the internal-engine thesis directly on the deep panel (2000-2026): if the
universe is moat-selected, operating income should grow faster than the market.
Per company we take the 3-year-average NOPAT at the START and END of its history
(smooths cyclicality) and compute the CAGR between the window midpoints. We report
the DISTRIBUTION (median company), not a currency-summed aggregate — because the
panel is multi-currency and summing local-currency NOPAT is meaningless; a per-company
growth RATE is currency-internal and comparable. Then we split by moat to see whether
higher moat -> faster compounding.

Honest caveats printed with the result: survivorship inflates it (dead names absent),
growth is NOMINAL local-currency (includes local inflation), and includes INORGANIC
(M&A) growth, not just organic.

  python3 pricing/nopat_growth.py
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
MIN_SPAN = 12          # need a long enough window for a meaningful CAGR

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
moat = {str(r[ci["Ticker"]]): r[ci["CompanyMoat(v3.2)"]] for r in rws[1:]
        if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}

print("loading deep panel...", flush=True)
by, idx = panel30.load(XLSB)
g = lambda r, k: (r[idx[k]] if k in idx and idx[k] < len(r) else None)
rows = []
for t, rs in by.items():
    t = str(t)
    ser = {}
    for r in rs:
        d = pd.to_datetime(g(r, "Period_End_Date") or g(r, "Date"), errors="coerce")
        n = num(g(r, "New Operating Income"))
        if pd.isna(d) or n is None or d.year < 2000:
            continue
        ser[d.year] = n
    yrs = sorted(ser)
    if len(yrs) < MIN_SPAN or (yrs[-1]-yrs[0]) < MIN_SPAN:
        continue
    first3 = np.mean([ser[y] for y in yrs[:3]]); last3 = np.mean([ser[y] for y in yrs[-3:]])
    span = np.mean(yrs[-3:]) - np.mean(yrs[:3])
    rows.append(dict(Instrument=t, moat=moat.get(t), start=first3, end=last3, span=span,
                     cagr=((last3/first3)**(1/span)-1) if (first3 > 0 and last3 > 0 and span > 0) else np.nan))
D = pd.DataFrame(rows)
ok = D.dropna(subset=["cagr"])
neg_start = (D.start <= 0).sum()

def line(name, s):
    if len(s) == 0:
        print(f"  {name:26s}  n=0"); return
    print(f"  {name:26s}  n={len(s):4d}  median {s.cagr.median()*100:5.1f}%   "
          f"mean {s.cagr.mean()*100:5.1f}%   25th {s.cagr.quantile(.25)*100:5.1f}%   75th {s.cagr.quantile(.75)*100:5.1f}%   "
          f">0: {(s.cagr>0).mean()*100:3.0f}%")

print(f"\nNew Operating Income compounding, per company, {MIN_SPAN}+ yr history (avg span {ok.span.mean():.0f} yr)")
print(f"companies measured: {len(ok)}  (excluded {neg_start} with non-positive starting income)\n")
line("ALL companies", ok)
line("  moat >= 7.8 (highest)", ok[ok.moat >= 7.8])
line("  moat 7.0 - 7.8", ok[(ok.moat >= 7.0) & (ok.moat < 7.8)])
line("  moat 6.5 - 7.0", ok[(ok.moat >= 6.5) & (ok.moat < 7.0)])
line("  moat < 6.5 (weakest)", ok[ok.moat < 6.5])
print(f"\n  Core-30 gate universe (moat >= 6.5):")
line("  moat >= 6.5", ok[ok.moat >= 6.5])
# monotonicity check: median cagr by moat decile
ok2 = ok.dropna(subset=["moat"]).copy()
ok2["mb"] = pd.qcut(ok2.moat, 5, labels=["Q1 low", "Q2", "Q3", "Q4", "Q5 high"], duplicates="drop")
print("\n  median NOPAT CAGR by moat quintile (does higher moat compound faster?):")
mq = ok2.groupby("mb").cagr.median()*100
for k, v in mq.items():
    print(f"    {k:8s}: {v:5.1f}%")
print("\nCAVEATS: survivor-only panel (inflates growth); NOMINAL local-currency (includes")
print("local inflation ~2-3%/yr); includes M&A/inorganic growth; accounting NOPAT.")
D.to_parquet(SCR+"/nopat_growth.parquet", index=False)
