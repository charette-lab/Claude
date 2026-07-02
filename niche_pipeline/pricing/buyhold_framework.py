#!/usr/bin/env python3
"""buyhold_framework.py — the wide-moat buy-cheap-and-hold strategy, but forced to respect the
AIP risk framework (so the concentration/factor/severity risks are controlled).

Same core rule (CompanyMoat>=7.8, buy at er_total>=20%, hold until er_total<=0), PLUS the
framework's risk controls, which require trimming winners (you can no longer let one name run to
35% of NAV):
  - HOLDING CAP  : no position above 10% of the book (trim the excess back to cash each quarter)
  - FACTOR CAP   : <= TAG_CAP names may share any one of the 11 risk tags  (~20% weight cap = the 6/30 rule)
  - COUNTRY CAP  : <= COUNTRY_CAP names per country
  - SEVERITY     : drop Extreme (severity-4) names -- the wipeout screen
  - TAGS         : a name must carry the 11-tag classification to be sized (framework needs it)
Held names are kept while er_total>0 (buy-cheap/hold spirit) and processed first each quarter to
minimise churn; new entrants (er>=20%) fill remaining cash under the caps, best-er first. Cash
earns CASH_RATE. Point-in-time, quarterly, full history.

  BH_MIN_MCAP=100000000 python3 pricing/buyhold_framework.py
"""
import sys, os
import numpy as np, pandas as pd
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import backtest_core30 as B

SCR = B.SCR
MOAT_MIN, ER_BUY, ER_HOLD, POS_CAP = 7.8, 0.20, 0.0, 0.10
TAG_CAP = int(os.environ.get("BH_TAG_CAP", "2"))         # <=2 names/tag ~ 20% weight cap (6/30 rule)
COUNTRY_CAP = int(os.environ.get("BH_COUNTRY_CAP", "3"))
MIN_MCAP = float(os.environ.get("BH_MIN_MCAP", "0"))
CASH_RATE = float(os.environ.get("BH_CASH", "0.02")); ppy = 4; cash_q = (1+CASH_RATE)**0.25-1

tags, maxsev, moat, country = B.load_meta()
dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet", columns=["Instrument", "Date", "er_total", "artifact", "market_cap"])
dec["Instrument"] = dec["Instrument"].astype(str); dec["Date"] = pd.to_datetime(dec["Date"]).astype("datetime64[ns]")
fr = []
for f in B.PRICE_FILES:
    p = pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"]); p["Instrument"] = p["Instrument"].astype(str); fr.append(p)
px = pd.concat(fr); px["Date"] = pd.to_datetime(px["Date"]).astype("datetime64[ns]"); px["c"] = pd.to_numeric(px["Close Price"], errors="coerce")
px = px[px["c"] > 0].dropna(subset=["c"]).drop_duplicates(["Instrument", "Date"])
W = px.pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()
s = pd.Series(sorted(dec["Date"].unique()))
grid = pd.DatetimeIndex(sorted(list(s.groupby(s.dt.to_period("Q")).max().values)))
asof = lambda wide: wide.reindex(grid, method="ffill")
er_g = asof(dec.pivot_table(index="Date", columns="Instrument", values="er_total", aggfunc="last"))
art_g = asof(dec.pivot_table(index="Date", columns="Instrument", values="artifact", aggfunc="last"))
mc_g = asof(dec.pivot_table(index="Date", columns="Instrument", values="market_cap", aggfunc="last"))
px_g = asof(W)

# wide-moat, has-tags, not-Extreme candidate columns
wide = [c for c in er_g.columns if (moat.get(c) or 0) >= MOAT_MIN and c in tags and len(tags[c]) > 0 and maxsev.get(c, 9) < 4]
print(f"framework-eligible wide-moat names (tags present, sev<4): {len(wide)}  | TAG_CAP {TAG_CAP} COUNTRY_CAP {COUNTRY_CAP} floor ${MIN_MCAP/1e6:.0f}M")

def caps_ok(r, tagc, cc):
    for tg in tags.get(r, []):
        if tagc.get(tg, 0)+1 > TAG_CAP:
            return False
    if country.get(r) is not None and cc.get(country.get(r), 0)+1 > COUNTRY_CAP:
        return False
    return True

holdings = {}; cash = 1.0
rets, ncnt, cashpct, conc = [], [], [], []
prev_total = 1.0
for i, t in enumerate(grid):
    erT, artT, pxT, mcT = er_g.loc[t], art_g.loc[t], px_g.loc[t], mc_g.loc[t]
    for r, h in list(holdings.items()):
        p = pxT.get(r)
        if p is not None and not pd.isna(p) and h["px"] and not pd.isna(h["px"]) and h["px"] > 0:
            h["val"] *= float(p)/float(h["px"]); h["px"] = float(p)
    cash *= (1+cash_q)
    total = sum(h["val"] for h in holdings.values()) + cash
    if i > 0:
        rets.append((t, total/prev_total-1.0))
    # --- sell exits (er<=0 or price gone) ---
    for r in list(holdings):
        e = erT.get(r)
        if (e is not None and not pd.isna(e) and float(e) <= ER_HOLD):
            cash += holdings[r]["val"]; del holdings[r]
    total = sum(h["val"] for h in holdings.values()) + cash
    # --- trim to 10% holding cap ---
    for r, h in holdings.items():
        if h["val"] > POS_CAP*total:
            cash += h["val"]-POS_CAP*total; h["val"] = POS_CAP*total
    # --- current tag/country counts from held ---
    tagc, cc = {}, {}
    for r in holdings:
        for tg in tags.get(r, []):
            tagc[tg] = tagc.get(tg, 0)+1
        cc[country.get(r)] = cc.get(country.get(r), 0)+1
    # --- deploy cash into new entrants (er>=20%, caps ok, floor), best er first ---
    total = sum(h["val"] for h in holdings.values()) + cash
    cand = []
    for r in wide:
        if r in holdings:
            continue
        e = erT.get(r); p = pxT.get(r); mv = mcT.get(r)
        if e is None or pd.isna(e) or float(e) < ER_BUY or bool(artT.get(r)) or p is None or pd.isna(p) or p <= 0:
            continue
        if MIN_MCAP > 0 and (mv is None or pd.isna(mv) or float(mv) < MIN_MCAP):
            continue
        cand.append((r, float(e), float(p)))
    cand.sort(key=lambda x: -x[1])
    for r, e, p in cand:
        buy = POS_CAP*total
        if cash < buy or buy <= 0:
            continue
        if not caps_ok(r, tagc, cc):
            continue
        holdings[r] = {"val": buy, "px": p}; cash -= buy
        for tg in tags.get(r, []):
            tagc[tg] = tagc.get(tg, 0)+1
        cc[country.get(r)] = cc.get(country.get(r), 0)+1
    prev_total = sum(h["val"] for h in holdings.values()) + cash
    tt = prev_total
    ncnt.append(len(holdings)); cashpct.append(cash/tt if tt > 0 else 1.0)
    vv = np.sort(np.array([h["val"] for h in holdings.values()]))[::-1]/tt if (tt > 0 and holdings) else np.array([])
    hhi = float((vv**2).sum()) if len(vv) else 0.0
    conc.append(dict(top1=float(vv[0]) if len(vv) else 0.0, top3=float(vv[:3].sum()) if len(vv) else 0.0,
                     effN=(1.0/hhi if hhi > 0 else 0.0)))

ret_s = pd.Series([x for _, x in rets], index=pd.DatetimeIndex([d for d, _ in rets]), name="bh_framework")

def perf(r):
    r = r.dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    dds = np.sqrt((np.minimum(r, 0.0)**2).mean())*np.sqrt(ppy); cagr = eq.iloc[-1]**(1/y)-1
    return dict(CAGR=cagr*100, Vol=r.std(ddof=1)*np.sqrt(ppy)*100, DownVol=dds*100, MaxDD=dd.min()*100,
                Sharpe=(r.mean()*ppy-0.03)/(r.std(ddof=1)*np.sqrt(ppy)), Sortino=(r.mean()*ppy-0.03)/dds if dds > 0 else np.nan,
                Calmar=cagr/abs(dd.min()) if dd.min() < 0 else np.nan, FinalNAV=eq.iloc[-1])

st = perf(ret_s)
FTAG = f"_mcap{int(MIN_MCAP/1e6)}M" if MIN_MCAP > 0 else ""
print(f"\nperiod {ret_s.index.min():%Y}-{ret_s.index.max():%Y} | avg positions {np.mean(ncnt):.1f} | avg cash {np.mean(cashpct)*100:.0f}% | avg effN {np.mean([c['effN'] for c in conc]):.1f} | avg top1 {np.mean([c['top1'] for c in conc])*100:.0f}%")
for k in ["CAGR", "Vol", "DownVol", "MaxDD", "Sharpe", "Sortino", "Calmar", "FinalNAV"]:
    u = "%" if k in ("CAGR", "Vol", "DownVol", "MaxDD") else ""
    print(f"  {k:9s}: " + (f"{st[k]:6.1f}{u}" if u else f"{st[k]:6.2f}"))
pd.DataFrame({"date": grid, "n_pos": ncnt, "cash_pct": cashpct,
             "top1": [c["top1"] for c in conc], "top3": [c["top3"] for c in conc], "effN": [c["effN"] for c in conc]}).to_csv(SCR+f"/bh_framework_state{FTAG}.csv", index=False)
ret_s.to_frame().to_parquet(SCR+f"/bh_framework_Q{FTAG}.parquet")
print(f"wrote bh_framework_Q{FTAG}.parquet")
