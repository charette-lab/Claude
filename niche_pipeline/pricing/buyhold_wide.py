#!/usr/bin/env python3
"""buyhold_wide.py — "buy the best companies cheap and keep them."

Rule (exactly as specified):
  UNIVERSE : CompanyMoat(v3.2) >= 7.8  (wide-moat only -- the best companies)
  BUY      : real-engine expected return er_total >= 20%  (cheap), 10% of portfolio per NEW name
  HOLD     : never trim a winner; let it ride
  SELL     : only when that name's er_total falls to <= 0  (no longer any expected return)
  SIZING   : no NEW position above 10%; winners are allowed to drift above 10% (we don't sell)
  CASH     : whatever isn't invested sits in cash (earns a modest CASH_RATE)

Point-in-time, quarterly, full history. er_total is the supply/demand-normalised real engine
(er_adj). Buys/sells are value-neutral transfers between cash and holdings at the current price,
so the period return comes only from price moves on holdings + cash yield.

  python3 pricing/buyhold_wide.py
"""
import sys, os
import numpy as np, pandas as pd
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import backtest_core30 as B

SCR = B.SCR
MOAT_MIN = 7.8
ER_BUY = 0.20
ER_SELL = 0.0
POS_CAP = 0.10
MIN_MCAP = float(os.environ.get("BH_MIN_MCAP", "0"))   # market-cap floor at ENTRY, USD (0 = off)
CASH_RATE = float(os.environ.get("BH_CASH", "0.02"))   # annual cash yield
ppy = 4; cash_q = (1+CASH_RATE)**0.25 - 1

# ---- meta (moat) ----
_, _, moat, _ = B.load_meta()

# ---- decomposition: er_total (=er_adj), artifact ----
dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet", columns=["Instrument", "Date", "er_total", "artifact", "market_cap"])
dec["Instrument"] = dec["Instrument"].astype(str); dec["Date"] = pd.to_datetime(dec["Date"]).astype("datetime64[ns]")

# ---- price panel (realized returns) ----
fr = []
for f in B.PRICE_FILES:
    p = pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"]); p["Instrument"] = p["Instrument"].astype(str)
    fr.append(p)
px = pd.concat(fr); px["Date"] = pd.to_datetime(px["Date"]).astype("datetime64[ns]")
px["c"] = pd.to_numeric(px["Close Price"], errors="coerce")
px = px[px["c"] > 0].dropna(subset=["c"]).drop_duplicates(["Instrument", "Date"])
W = px.pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()

s = pd.Series(sorted(dec["Date"].unique()))
grid = pd.DatetimeIndex(sorted(list(s.groupby(s.dt.to_period("Q")).max().values)))
asof = lambda wide: wide.reindex(grid, method="ffill")
er_g = asof(dec.pivot_table(index="Date", columns="Instrument", values="er_total", aggfunc="last"))
art_g = asof(dec.pivot_table(index="Date", columns="Instrument", values="artifact", aggfunc="last"))
mc_g = asof(dec.pivot_table(index="Date", columns="Instrument", values="market_cap", aggfunc="last")) if MIN_MCAP > 0 else None
px_g = asof(W)

# wide-moat candidate columns (static moat label)
wide = [c for c in er_g.columns if moat.get(c, 0) is not None and (moat.get(c, 0) or 0) >= MOAT_MIN]
print(f"wide-moat (>= {MOAT_MIN}) names in engine coverage: {len(wide)}")

holdings = {}          # inst -> {"val": value, "px": last price}
cash = 1.0
rets, ncnt, cashpct, conc = [], [], [], []
snaps = []             # per-period holdings snapshot (date, inst, val, px) for attribution
prev_total = 1.0
for i, t in enumerate(grid):
    erT, artT, pxT = er_g.loc[t], art_g.loc[t], px_g.loc[t]
    # 1) mark holdings to today's price; grow cash
    for r, h in list(holdings.items()):
        p = pxT.get(r)
        if p is not None and not pd.isna(p) and h["px"] and not pd.isna(h["px"]) and h["px"] > 0:
            h["val"] *= float(p)/float(h["px"]); h["px"] = float(p)
    cash *= (1+cash_q)
    total = sum(h["val"] for h in holdings.values()) + cash
    if i > 0:
        rets.append((t, total/prev_total - 1.0))
    ncnt.append(len(holdings)); cashpct.append(cash/total if total > 0 else 1.0)
    # concentration of the invested book (weights as % of total portfolio incl. cash)
    vv = np.sort(np.array([h["val"] for h in holdings.values()], dtype=float))[::-1]
    w = vv/total if (total > 0 and len(vv)) else np.array([])
    hhi = float((w**2).sum()) if len(w) else 0.0
    conc.append(dict(top1=float(w[0]) if len(w) >= 1 else 0.0,
                     top3=float(w[:3].sum()) if len(w) else 0.0,
                     top5=float(w[:5].sum()) if len(w) else 0.0,
                     hhi=hhi, effN=(1.0/hhi if hhi > 0 else 0.0)))
    for r, h in holdings.items():                       # snapshot for return attribution
        snaps.append((t, r, h["val"], h["px"]))
    # 2) SELL: er has fallen to <= 0
    for r in list(holdings):
        e = erT.get(r)
        if e is not None and not pd.isna(e) and float(e) <= ER_SELL:
            cash += holdings[r]["val"]; del holdings[r]
    # 3) BUY: wide-moat, er >= 20%, not artifact, priced, not held -- best er first, 10% chunks from cash
    total = sum(h["val"] for h in holdings.values()) + cash
    mcT = mc_g.loc[t] if mc_g is not None else None
    cand = []
    for r in wide:
        if r in holdings:
            continue
        e = erT.get(r); p = pxT.get(r)
        if e is None or pd.isna(e) or float(e) < ER_BUY:
            continue
        if bool(artT.get(r)) or p is None or pd.isna(p) or float(p) <= 0:
            continue
        if MIN_MCAP > 0:                                # size floor at entry (no look-ahead)
            mv = mcT.get(r)
            if mv is None or pd.isna(mv) or float(mv) < MIN_MCAP:
                continue
        cand.append((r, float(e), float(p)))
    cand.sort(key=lambda x: -x[1])
    buy = POS_CAP*total
    for r, e, p in cand:
        if cash < buy or buy <= 0:
            break
        holdings[r] = {"val": buy, "px": p}; cash -= buy
    prev_total = sum(h["val"] for h in holdings.values()) + cash

ret_s = pd.Series([x for _, x in rets], index=pd.DatetimeIndex([d for d, _ in rets]), name="buyhold_wide")

def perf(r):
    r = r.dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    dds = np.sqrt((np.minimum(r, 0.0)**2).mean())*np.sqrt(ppy); cagr = eq.iloc[-1]**(1/y)-1
    return dict(CAGR=cagr*100, Vol=r.std(ddof=1)*np.sqrt(ppy)*100, DownVol=dds*100, MaxDD=dd.min()*100,
                Sharpe=(r.mean()*ppy-0.03)/(r.std(ddof=1)*np.sqrt(ppy)),
                Sortino=(r.mean()*ppy-0.03)/dds if dds > 0 else np.nan,
                Calmar=cagr/abs(dd.min()) if dd.min() < 0 else np.nan, FinalNAV=eq.iloc[-1])

st = perf(ret_s)
core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]; core.index = pd.to_datetime(core.index)
core = core[(core.index.year >= 2000) & (core.index.year <= 2026)]; stc = perf(core)

print(f"\nperiod {ret_s.index.min():%Y-%m}..{ret_s.index.max():%Y-%m} | cash yield {CASH_RATE*100:.1f}%/yr | avg positions {np.mean(ncnt):.1f} | avg cash {np.mean(cashpct)*100:.0f}%")
print(f"\n{'metric':>9} | {'BUY&HOLD wide':>13} | {'Core-30':>8}")
print("-"*38)
for k in ["CAGR", "Vol", "DownVol", "MaxDD", "Sharpe", "Sortino", "Calmar", "FinalNAV"]:
    u = "%" if k in ("CAGR", "Vol", "DownVol", "MaxDD") else ""
    fmt = (lambda v: f"{v:6.1f}{u}") if u else (lambda v: f"{v:6.2f}")
    print(f"{k:>9} | {fmt(st[k]):>13} | {fmt(stc[k]):>8}")

# cash trajectory (valuation timing) at a few dates
cs = pd.Series(cashpct, index=grid)
print("\ncash weight at key dates:")
for y in [2000, 2003, 2007, 2009, 2013, 2018, 2021, 2023, 2026]:
    sub = cs[cs.index.year == y]
    if len(sub):
        print(f"  {y}: {sub.iloc[-1]*100:3.0f}% cash")
FTAG = f"_mcap{int(MIN_MCAP/1e6)}M" if MIN_MCAP > 0 else ""
ret_s.to_frame().to_parquet(SCR+f"/buyhold_wide_Q{FTAG}.parquet")
state = pd.DataFrame({"date": grid, "n_pos": ncnt, "cash_pct": cashpct})
for k in ["top1", "top3", "top5", "hhi", "effN"]:
    state[k] = [c[k] for c in conc]
state.to_csv(SCR+f"/buyhold_wide_state{FTAG}.csv", index=False)
pd.DataFrame(snaps, columns=["date", "inst", "val", "px"]).to_parquet(SCR+f"/buyhold_wide_snaps{FTAG}.parquet")

# per-year concentration table (year-end snapshot)
state["year"] = pd.to_datetime(state["date"]).dt.year
gy = state.groupby("year").last()
print("\n=== CONCENTRATION per year (year-end weights, % of total incl. cash) ===")
print(f"{'year':>4} | {'n':>3} | {'top1':>5} | {'top3':>5} | {'top5':>5} | {'effN':>5} | {'cash':>5}")
print("-"*50)
for y, r in gy.iterrows():
    print(f"{y:>4} | {r.n_pos:3.0f} | {r.top1*100:4.0f}% | {r.top3*100:4.0f}% | {r.top5*100:4.0f}% | {r.effN:5.1f} | {r.cash_pct*100:4.0f}%")
print("\nwrote buyhold_wide_Q.parquet, buyhold_wide_state.csv")
