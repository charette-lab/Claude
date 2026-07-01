#!/usr/bin/env python3
"""stage2_buriedcore.py — do the buried-core candidates actually OUTPERFORM? (2016-2026)

Point-in-time, no look-ahead. Each year we trade on June 30 using the most recent fiscal
year that ended >=6 months earlier (reporting lag). A name is a BURIED-CORE candidate if:
  CoreMoat>=7.8  AND  currently 'wrecked' (latest 3yr NOI margin < 80% of its pre-pivot peak,
  computed only from history up to that FY)  AND  cheap on the core: normalized core earnings
  yield = pre-pivot margin x sales / EV(at trade date) in the top half of wrecked elite cores.
Equal-weight the candidates, hold one year, measure the realized price return, compare to the
universe and to elite cores that are NOT wrecked. Also split the basket by whether a recent
impairment was booked, to test if impairment adds any signal.

  python3 pricing/stage2_buriedcore.py
"""
import sys, warnings
import numpy as np, pandas as pd, openpyxl
warnings.filterwarnings("ignore")
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
import panel30

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB = [U+"fb2aec33-30_file_1.xlsb", U+"a810c35f-30_file_2.xlsb", U+"62545fc2-30_file_3.xlsb"]
PRICE = [U+"972f0581-daily_volume_price_0.parquet", U+"257124b3-daily_volume_price_1.parquet",
         U+"86e54ec3-daily_volume_price_2.parquet", U+"13f82c18-daily_volume_price_0630.parquet"]
num = lambda v: v if isinstance(v, (int, float)) else None
IMPAIR = ["Adjusted_Impairment___Goodwill__ASR", "Adjusted_Impairment___Intangibles_excluding_Goodwill__ASR",
          "Adjusted_Impairment___Fixed_Assets", "Adjusted_Impairment___Investment_Property___Cash_Flow"]

ws = openpyxl.load_workbook(SCR+"/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
coremoat = {str(r[ci["Ticker"]]): r[ci["CoreMoat(v3.2)"]] for r in rws[1:]
            if r[ci["Ticker"]] and isinstance(r[ci["CoreMoat(v3.2)"]], (int, float))}

print("loading prices...", flush=True)
px = pd.concat([pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"]) for f in PRICE])
px["Instrument"] = px["Instrument"].astype(str); px["Date"] = pd.to_datetime(px["Date"])
px["c"] = pd.to_numeric(px["Close Price"], errors="coerce")
W = px[px.c > 0].pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()
def price_asof(tkr, d):
    if tkr not in W.columns: return None
    s = W[tkr].loc[:d].dropna()
    return float(s.iloc[-1]) if len(s) else None

print("loading deep panel...", flush=True)
by, idx = panel30.load(XLSB)
def GET(r, *names):
    for k in names:
        if k in idx and idx[k] < len(r) and isinstance(r[idx[k]], (int, float)):
            return r[idx[k]]
    return None

# per name: annual records {year: {...}} from fundamentals
FUND = {}
for t, rs in by.items():
    t = str(t)
    if coremoat.get(t, 0) < 7.8:
        continue
    ser = {}
    for r in rs:
        d = pd.to_datetime(GET(r, "Period_End_Date") or (r[idx["Date"]] if "Date" in idx else None), errors="coerce")
        noi, sal = num(GET(r, "New Operating Income")), num(GET(r, "Sales"))
        ev, mc = num(GET(r, "EV")), num(GET(r, "Market Capitalization", "Market_Capitalization"))
        if pd.isna(d) or noi is None or sal is None or sal <= 0 or ev is None or mc is None:
            continue
        imp = sum(abs(num(GET(r, k)) or 0.0) for k in IMPAIR)
        ser[d.year] = dict(ped=d, m=noi/sal, sal=sal, ev=ev, mc=mc, nd=ev-mc, imp=imp)
    if len(ser) >= 6:
        FUND[t] = ser

def prepivot(ser, upto):
    yrs = [y for y in sorted(ser) if y <= upto]
    if len(yrs) < 5: return None, None
    m = pd.Series({y: ser[y]["m"] for y in yrs if abs(ser[y]["m"]) <= 0.60})
    if len(m) < 5: return None, None
    roll = m.rolling(3, min_periods=2).mean()
    pre = roll.max()
    cur = pd.Series({y: ser[y]["m"] for y in yrs}).rolling(3, min_periods=1).mean().loc[yrs[-1]]
    return (pre if (pre is not None and 0 < pre <= 0.50) else None), cur

# ---- annual point-in-time backtest ----
def hold_ret(tkr, d0, d1):
    p0, p1 = price_asof(tkr, d0), price_asof(tkr, d1)
    return (p1/p0 - 1) if (p0 and p1 and p0 > 0) else None

rows = []
universe = [str(c) for c in W.columns]
for yt in range(2016, 2026):
    d0, d1 = pd.Timestamp(yt, 6, 30), pd.Timestamp(yt+1, 6, 30)
    fy = yt-1                                            # latest FY ended >=6mo earlier
    cands, impair_flag, yields = [], {}, {}
    for t, ser in FUND.items():
        if fy not in ser and (fy-1) not in ser: continue
        y = fy if fy in ser else fy-1
        pre, cur = prepivot(ser, y)
        if pre is None or cur is None or cur >= 0.80*pre:   # must be 'wrecked'
            continue
        s = ser[y]; p_pe = price_asof(t, s["ped"]); p_tr = price_asof(t, d0)
        if not p_pe or not p_tr: continue
        ev_tr = s["nd"] + s["mc"]*(p_tr/p_pe)               # EV at trade date (net debt + repriced mktcap)
        if ev_tr <= 0: continue
        y_core = pre*s["sal"]/ev_tr                         # normalized core earnings yield
        yields[t] = y_core
        recent_imp = sum(ser[yy]["imp"] for yy in ser if y-2 <= yy <= y)
        impair_flag[t] = recent_imp > 0.02*s["ev"]          # impairment >2% of EV in last 3y
    if len(yields) < 5: continue
    cut = np.median(list(yields.values()))
    picks = [t for t, v in yields.items() if v >= cut]      # cheap half of wrecked elite cores
    rets = {t: hold_ret(t, d0, d1) for t in picks}
    rets = {t: v for t, v in rets.items() if v is not None}
    if not rets: continue
    # benchmarks over same window
    uni = [hold_ret(t, d0, d1) for t in universe]; uni = [v for v in uni if v is not None]
    elite = [hold_ret(t, d0, d1) for t in FUND]; elite = [v for v in elite if v is not None]
    wi = [rets[t] for t in rets if impair_flag.get(t)]; woi = [rets[t] for t in rets if not impair_flag.get(t)]
    rows.append(dict(year=yt, n=len(rets), buried=np.mean(list(rets.values())),
                     universe=np.mean(uni), elite=np.mean(elite),
                     buried_med=np.median(list(rets.values())),
                     with_impair=np.mean(wi) if wi else np.nan, no_impair=np.mean(woi) if woi else np.nan,
                     n_imp=len(wi)))
R = pd.DataFrame(rows).set_index("year")
def cagr(s): s = (1+pd.Series(s).dropna()); return s.prod()**(1/len(s))-1
def stats(s):
    s = pd.Series(s).dropna(); eq = (1+s).cumprod(); dd = eq/eq.cummax()-1
    return cagr(s)*100, s.std(ddof=1)*100, dd.min()*100
print(f"\n{'year':6}{'n':>4}{'buried%':>9}{'universe%':>10}{'elite%':>8}")
for y, r in R.iterrows():
    print(f"{y:<6}{int(r.n):>4}{r.buried*100:>8.1f}{r.universe*100:>10.1f}{r.elite*100:>8.1f}")
print("\n--- annual-return summary (equal-weight, 1yr hold, gross price return) ---")
for name, col in [("BURIED-CORE basket", R.buried), ("Elite-core (all)", R.elite), ("Universe EW", R.universe)]:
    c, v, d = stats(col); print(f"  {name:22s} CAGR {c:5.1f}%  vol {v:4.1f}%  worst yr {col.min()*100:5.1f}%")
print(f"\n  buried-core beat universe in {int((R.buried>R.universe).sum())}/{len(R)} years; "
      f"beat elite-core in {int((R.buried>R.elite).sum())}/{len(R)} years")
print("\n--- impairment test (does a recent impairment add signal?) ---")
wi, woi = cagr(R.with_impair)*100, cagr(R.no_impair)*100
print(f"  candidates WITH recent impairment: CAGR {wi:5.1f}%   |  WITHOUT: {woi:5.1f}%   (avg {R.n_imp.mean():.0f} names/yr had one)")
R.to_parquet(SCR+"/stage2_buriedcore.parquet")
print("\nCAVEATS: gross price return (no dividends/FX/costs); today's CoreMoat backward;")
print("cheapness proxy = normalized yield (not yet the full supply/demand-normalized engine).")
