#!/usr/bin/env python3
"""mcap_sweep.py — find the OPTIMAL point-in-time market-cap floor for the Core-30.

Loads the price/decomposition data ONCE, then re-runs the exact Core-30 selection/weighting
loop for a ladder of USD market-cap floors, reporting CAGR / Vol / Sharpe / MaxDD / Calmar /
avg book size for each. The optimum is the floor that maximises risk-adjusted return (Calmar /
Sharpe) before the book starts thinning below the 30-name target and giving up CAGR.

  python3 pricing/mcap_sweep.py
"""
import sys, os, json
import numpy as np, pandas as pd
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import backtest_core30 as B
from frameworks import CORE_N, GATE1_IRR

SCR = B.SCR
FLOORS = [0, 25e6, 50e6, 75e6, 100e6, 150e6, 200e6, 300e6, 500e6, 750e6, 1000e6]
cadence = "Q"; ppy = 4

# ---------- load once ----------
tags, maxsev, moat, country = B.load_meta()
dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet")
dec["Instrument"] = dec["Instrument"].astype(str); dec["Date"] = pd.to_datetime(dec["Date"]).astype("datetime64[ns]")
fr = []
for f in B.PRICE_FILES:
    p = pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"]); p["Instrument"] = p["Instrument"].astype(str)
    fr.append(p)
px = pd.concat(fr).rename(columns={"Close Price": "c"})
px["Date"] = pd.to_datetime(px["Date"]).astype("datetime64[ns]"); px["c"] = pd.to_numeric(px["c"], errors="coerce")
px = px[px["c"] > 0].dropna().drop_duplicates(["Instrument", "Date"])
W = px.pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()
ret = W.pct_change()
vol = ret.rolling(B.VOL_WIN, min_periods=60).std()*np.sqrt(252)
hist = W.notna().cumsum()
s = pd.Series(sorted(dec["Date"].unique()))
grid = pd.DatetimeIndex(sorted(list(s.groupby(s.dt.to_period("Q")).max().values)))
asof = lambda wide: wide.reindex(grid, method="ffill")
er_g = asof(dec.pivot_table(index="Date", columns="Instrument", values="er_total", aggfunc="last"))
art_g = asof(dec.pivot_table(index="Date", columns="Instrument", values="artifact", aggfunc="last"))
mc_g = asof(dec.pivot_table(index="Date", columns="Instrument", values="market_cap", aggfunc="last"))
vol_g, hist_g, px_g = asof(vol), asof(hist), asof(W)


def run_floor(min_mcap):
    rows, books_n = [], []
    for i in range(len(grid)-1):
        t, t1 = grid[i], grid[i+1]
        erT, vT, mcT = er_g.loc[t], vol_g.loc[t], mc_g.loc[t]
        elig = []
        for r in erT.index:
            if r not in moat or moat[r] < B.MIN_MOAT or r not in tags or maxsev.get(r, 0) >= 4:
                continue
            e = erT.get(r); v = vT.get(r); h = hist_g.loc[t].get(r)
            if pd.isna(e) or e < GATE1_IRR or bool(art_g.loc[t].get(r)):
                continue
            if pd.isna(v) or v <= 0 or pd.isna(h) or h < B.MIN_HIST or pd.isna(px_g.loc[t].get(r)):
                continue
            if min_mcap > 0:
                mcv = mcT.get(r)
                if pd.isna(mcv) or float(mcv) < min_mcap:
                    continue
            elig.append((r, float(e)/float(v)))
        elig.sort(key=lambda x: -x[1])
        book, buckets, cc = [], {}, {}
        for r, _ in elig:
            if len(book) >= CORE_N:
                break
            ts = tags.get(r, frozenset()); co = country.get(r)
            if any(buckets.get(x, 0)+1 > B.SLOT_CAP for x in ts) or (co is not None and cc.get(co, 0)+1 > B.COUNTRY_CAP):
                continue
            book.append(r)
            for x in ts:
                buckets[x] = buckets.get(x, 0)+1
            cc[co] = cc.get(co, 0)+1
        if not book:
            continue
        iv = np.array([1.0/float(vT[r]) for r in book]); w = np.clip(iv/iv.sum(), B.W_MIN, B.W_MAX); w = w/w.sum()
        p0 = px_g.loc[t].reindex(book).to_numpy(dtype="float64"); p1 = px_g.loc[t1].reindex(book).to_numpy(dtype="float64")
        r_i = p1/p0 - 1.0; ok = ~np.isnan(r_i)
        if ok.sum() == 0:
            continue
        rows.append((t1, float(np.nansum(w[ok]*r_i[ok])/w[ok].sum()))); books_n.append(len(book))
    ret_s = pd.Series([x for _, x in rows], index=pd.DatetimeIndex([d for d, _ in rows]))
    st = B.perf(ret_s, ppy)
    return st, float(np.mean(books_n))


print(f"{'Floor':>8} | {'CAGR':>6} | {'Vol':>6} | {'Sharpe':>6} | {'MaxDD':>7} | {'Calmar':>6} | {'NAV':>6} | {'book':>4}")
print("-"*72)
results = []
for fl in FLOORS:
    st, bn = run_floor(fl)
    lbl = "none" if fl == 0 else f"${fl/1e6:.0f}M"
    results.append((fl, lbl, st, bn))
    print(f"{lbl:>8} | {st['CAGR']*100:5.1f}% | {st['Vol']*100:5.1f}% | {st['Sharpe']:6.2f} | "
          f"{st['MaxDD']*100:6.1f}% | {st['Calmar']:6.2f} | {st['FinalNAV']:6.1f} | {bn:4.1f}")
print("-"*72)
best_calmar = max(results, key=lambda x: x[2]['Calmar'])
best_sharpe = max(results, key=lambda x: x[2]['Sharpe'])
best_cagr = max(results, key=lambda x: x[2]['CAGR'])
print(f"best Calmar : {best_calmar[1]:>6}  (Calmar {best_calmar[2]['Calmar']:.2f}, CAGR {best_calmar[2]['CAGR']*100:.1f}%, MaxDD {best_calmar[2]['MaxDD']*100:.1f}%)")
print(f"best Sharpe : {best_sharpe[1]:>6}  (Sharpe {best_sharpe[2]['Sharpe']:.2f})")
print(f"best CAGR   : {best_cagr[1]:>6}  (CAGR {best_cagr[2]['CAGR']*100:.1f}%)")

# save a small csv for the record
pd.DataFrame([{"floor_usd": fl, "label": lbl, **st, "avg_book": bn} for fl, lbl, st, bn in results]).to_csv(SCR+"/mcap_sweep.csv", index=False)
print("\nwrote mcap_sweep.csv")
