#!/usr/bin/env python3
"""parking_lot.py — the Core re-tuned as a PARKING LOT, not a return engine.

Different job -> opposite objective. The return-optimized Core-30 ranks by er/vol and
maximizes compounding. A parking lot holds dry powder while 1-2 activist ideas/year are
sourced; it is judged on capital preservation, LIQUIDITY (exit fast to fund a strike),
low beta, and par-when-called — return is secondary. So we re-tune:

  * universe: same quality floor (moat, non-artifact) but require LARGE-CAP (top-half
    market cap) so every name is liquid and exitable in size;
  * a light return floor only (er >= 6%), NOT the 12% gate — we are not reaching for return;
  * RANK BY LOWEST VOLATILITY (min-vol tilt), not by er/vol;
  * inverse-vol weights, tighter caps, same slot/country diversification.

Then two fund structures for the SAME parking lot:
  * CLOSED / permanent capital (their fund, 80%+ permanent): can hold the equity parking
    lot at high weight — no redemptions, slow idiosyncratic deployment, rides out its drawdowns;
  * OPEN / redeemable: must hold more T-bills so dry powder is ~par when called.

  python3 pricing/parking_lot.py
"""
from __future__ import annotations
import json, os, sys
import numpy as np, pandas as pd
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import backtest_core30 as B
from frameworks import RISK_TAGS_SHORT

SCR = B.SCR
N = 30
MIN_MOAT = 6.5
PARK_GATE = 0.06          # light floor only; a parking lot isn't reaching for return
MCAP_Q = 0.50             # require top-half market cap (liquidity / exit-in-size)
SLOT_CAP, COUNTRY_CAP = 6, 10
W_MIN, W_MAX = 0.010, 0.050
BILL = 0.04               # T-bill / ultra-short parking yield (annual)


def build():
    tags, maxsev, moat, country = B.load_meta()
    dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet")
    dec["Instrument"] = dec["Instrument"].astype(str); dec["Date"] = pd.to_datetime(dec["Date"]).astype("datetime64[ns]")
    px = pd.concat([pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"]) for f in B.PRICE_FILES])
    px["Instrument"] = px["Instrument"].astype(str); px["Date"] = pd.to_datetime(px["Date"]).astype("datetime64[ns]")
    px["c"] = pd.to_numeric(px["Close Price"], errors="coerce")
    W = px[px.c > 0].pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()
    vol = W.pct_change().rolling(126, min_periods=60).std()*np.sqrt(252)
    hist = W.notna().cumsum()
    s = pd.Series(sorted(dec["Date"].unique()))
    grid = pd.DatetimeIndex(sorted(s.groupby(s.dt.to_period("Q")).max().values))
    aso = lambda w: w.reindex(grid, method="ffill")
    er_g = aso(dec.pivot_table(index="Date", columns="Instrument", values="er_total", aggfunc="last"))
    art_g = aso(dec.pivot_table(index="Date", columns="Instrument", values="artifact", aggfunc="last"))
    mc_g = aso(dec.pivot_table(index="Date", columns="Instrument", values="market_cap", aggfunc="last"))
    vol_g, hist_g, px_g = aso(vol), aso(hist), aso(W)

    rows, holds, sizes = [], {}, []
    for i in range(len(grid)-1):
        t, t1 = grid[i], grid[i+1]
        erT, vT, mcT = er_g.loc[t], vol_g.loc[t], mc_g.loc[t]
        mc_cut = mcT.dropna().quantile(MCAP_Q)
        elig = []
        for r in erT.index:
            if r not in moat or moat[r] < MIN_MOAT or r not in tags or maxsev.get(r, 0) >= 4:
                continue
            e, v, h, mc = erT.get(r), vT.get(r), hist_g.loc[t].get(r), mcT.get(r)
            if pd.isna(e) or e < PARK_GATE or bool(art_g.loc[t].get(r)):
                continue
            if pd.isna(v) or v <= 0 or pd.isna(h) or h < 60 or pd.isna(px_g.loc[t].get(r)):
                continue
            if pd.isna(mc) or mc < mc_cut:                 # large-cap / liquidity gate
                continue
            elig.append((r, float(v)))                     # RANK BY LOWEST VOLATILITY (min-vol parking lot)
        elig.sort(key=lambda x: x[1])
        book, buckets, cc = [], {}, {}
        for r, _ in elig:
            if len(book) >= N:
                break
            ts = tags.get(r, frozenset()); co = country.get(r)
            if any(buckets.get(x, 0)+1 > SLOT_CAP for x in ts):
                continue
            if co is not None and cc.get(co, 0)+1 > COUNTRY_CAP:
                continue
            book.append(r)
            for x in ts: buckets[x] = buckets.get(x, 0)+1
            cc[co] = cc.get(co, 0)+1
        if not book:
            continue
        iv = np.array([1.0/float(vT[r]) for r in book]); w = np.clip(iv/iv.sum(), W_MIN, W_MAX); w = w/w.sum()
        p0 = px_g.loc[t].reindex(book).to_numpy(dtype="float64"); p1 = px_g.loc[t1].reindex(book).to_numpy(dtype="float64")
        r_i = p1/p0-1.0; ok = ~np.isnan(r_i)
        if ok.sum() == 0:
            continue
        rows.append((t1, float(np.nansum(w[ok]*r_i[ok])/w[ok].sum())))
        holds[str(pd.Timestamp(t).date())] = book; sizes.append(len(book))
    idx = pd.DatetimeIndex([d for d, _ in rows])
    return pd.Series([x for _, x in rows], index=idx, name="ParkingLot"), np.mean(sizes)


def msci_q(index):
    m = pd.read_parquet(SCR+"/msci_imi_monthly.parquet").set_index("date")["ret"]; m.index = pd.to_datetime(m.index)
    mq = (1+m).groupby(m.index.to_period("Q")).prod()-1; mq.index = mq.index.to_timestamp("Q")
    j = mq.reindex(pd.PeriodIndex(index.to_period("Q")).to_timestamp("Q"))
    j.index = index
    return j


def stats(r, ppy=4):
    r = pd.Series(r).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy; dd = eq/eq.cummax()-1
    return dict(CAGR=eq.iloc[-1]**(1/y)-1, Vol=r.std(ddof=1)*np.sqrt(ppy), MaxDD=dd.min(),
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)),
                Calmar=(eq.iloc[-1]**(1/y)-1)/abs(dd.min()))


park, avgn = build()
ret_core = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]; ret_core.index = pd.to_datetime(ret_core.index)
ret_core.index = ret_core.index.to_period("Q").to_timestamp("Q")
park.index = park.index.to_period("Q").to_timestamp("Q")
mkt = msci_q(park.index)

print("="*78)
print(f"PARKING-LOT variant vs RETURN-optimized Core-30  (avg {avgn:.0f} names, {len(park)} quarters)")
print("="*78)
comp = pd.DataFrame({"Parking-lot (min-vol, large-cap)": stats(park),
                     "Return Core-30 (er/vol)": stats(ret_core.reindex(park.index).dropna())}).T
beta_p = np.cov(park.values, mkt.values)[0, 1]/np.var(mkt.values)
beta_c = np.cov(ret_core.reindex(park.index).dropna().values, mkt.reindex(ret_core.reindex(park.index).dropna().index).values)[0, 1]/np.var(mkt.values)
comp["beta_mkt"] = [beta_p, beta_c]
for c in ["CAGR", "Vol", "MaxDD"]: comp[c] = (comp[c]*100).round(1)
for c in ["Sharpe", "Calmar", "beta_mkt"]: comp[c] = comp[c].round(2)
print(comp[["CAGR", "Vol", "MaxDD", "Sharpe", "Calmar", "beta_mkt"]].to_string())

# behaviour in the 4 worst market quarters (prime counter-cyclical deployment windows)
df = pd.DataFrame({"park": park, "core": ret_core.reindex(park.index), "mkt": mkt}).dropna()
worst = df.sort_values("mkt").head(4)
print("\n  par-when-called test — value in the 4 worst market quarters (when you'd deploy into a crash):")
for d, row in worst.iterrows():
    print(f"    {d:%Y-%m}:  market {row.mkt*100:6.1f}%   parking-lot {row.park*100:6.1f}%   return-core {row.core*100:6.1f}%")

# tiered: equity parking lot blended with T-bills -> open vs closed configurations
print("\n"+"="*78)
print("OPEN vs CLOSED — same parking lot, different equity/bills mix")
print("="*78)
qbill = (1+BILL)**0.25-1
for label, weq in [("CLOSED / permanent  (90% equity parking lot)", 0.90),
                   ("CLOSED / permanent  (75% equity)", 0.75),
                   ("balanced            (50% equity)", 0.50),
                   ("OPEN / redeemable   (25% equity, 75% bills)", 0.25),
                   ("OPEN / redeemable   (100% T-bills)", 0.00)]:
    blend = weq*park + (1-weq)*qbill
    s = stats(blend); worst_q = (weq*worst["park"] + (1-weq)*qbill).min()
    print(f"  {label:46s} CAGR {s['CAGR']*100:5.1f}%  MaxDD {s['MaxDD']*100:6.1f}%  worst deploy-qtr {worst_q*100:6.1f}%")

# whole-fund translation: with ~20% of a permanent fund parked, what does the parking choice add?
print("\n  whole-fund translation (permanent fund, ~20% parked on average):")
for weq, tag in [(0.90, "closed: 90% equity parking"), (0.00, "open: T-bills only")]:
    plret = (weq*park + (1-weq)*qbill)
    add = 0.20*(stats(plret)["CAGR"]) + 0.80*0.18   # 80% deployed at ~18% book gross
    print(f"    {tag:28s}: parked sleeve CAGR {stats(plret)['CAGR']*100:4.1f}%  -> blended fund ~{add*100:4.1f}% (illustrative)")

import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
fig, (a1, a2) = plt.subplots(2, 1, figsize=(11, 8.5), height_ratios=[3, 1.5], sharex=True)
x0 = df.index.min()-pd.offsets.QuarterEnd()
series = {"Parking-lot (min-vol, large-cap)": (park, "#1f77b4"),
         "Return Core-30 (er/vol)": (ret_core.reindex(park.index).dropna(), "#2ca02c"),
         "T-bills 4%": (pd.Series(qbill, index=park.index), "#7f7f7f")}
for k, (r, c) in series.items():
    eq = (1+r).cumprod(); idx = pd.DatetimeIndex([x0]).append(pd.DatetimeIndex(r.index))
    a1.plot(idx, np.r_[1.0, eq.values], label=k, color=c, lw=2.2 if "Parking" in k else 1.6)
    dd = eq/eq.cummax()-1; a2.plot(r.index, dd*100, color=c, lw=1.6)
a1.set_yscale("log"); a1.legend(loc="upper left"); a1.grid(True, which="both", alpha=0.25); a1.set_ylabel("Growth of 1.0 (log)")
a1.set_title("Parking-lot variant (preservation/liquidity) vs return-optimized Core vs T-bills\n"
             "min-vol + large-cap re-tune: lower return, but shallower drawdown and lower beta for dry powder", fontsize=10.5)
a2.grid(True, alpha=0.25); a2.set_ylabel("drawdown (%)")
plt.tight_layout(); fig.savefig(SCR+"/parking_lot.png", dpi=130); print("\nwrote parking_lot.png")
