#!/usr/bin/env python3
"""backtest_system.py — the integrated system end-to-end (SYSTEM.md).

Core (robust compounder index, ~75%) + Satellite (concentrated conviction, ~25%)
+ external put hedge on MSCI World IMI. One combined equity/drawdown curve.

  AIP_VALUE_ENGINE not needed (reads the persisted decomposition).
  python3 pricing/backtest_system.py
"""
from __future__ import annotations
import os, sys, json
import numpy as np, pandas as pd
from scipy.stats import norm
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from frameworks import RISK_TAGS_SHORT, interpolate_weights
from backtest_core30 import load_meta, SCR, U, PRICE_FILES

# ---- params ----
SAT_FRAC = float(os.environ.get("SAT_FRAC", "0.25")); CORE_FRAC = 1.0 - SAT_FRAC
CORE_GATE, SAT_GATE = 0.12, 0.18
CORE_MOAT = 6.5           # match the tuned Resilient Core-30
CORE_N, SAT_N = 30, 10
SLOT_CAP, COUNTRY_CAP = 6, 10
CORE_WMAX, SAT_WMAX = 0.075, 0.08
VOL_WIN = 126
HEDGE_MONEY, HEDGE_TN = 0.90, 2         # 10% OTM, semi-annual roll
VRP, SKEW, RF = 0.03, 0.05, 0.02


def bsput(S, K, s, r, t):
    if t <= 1e-9:
        return max(K-S, 0.0)
    d1 = (np.log(S/K)+(r+0.5*s*s)*t)/(s*np.sqrt(t))
    return K*np.exp(-r*t)*norm.cdf(-(d1-s*np.sqrt(t))) - S*norm.cdf(-d1)


def perf(r, ppy=4):
    r = pd.Series(r).dropna(); eq = (1+r).cumprod(); y = len(r)/ppy
    dd = eq/eq.cummax()-1; roll = eq.pct_change(ppy).dropna()
    return dict(CAGR=eq.iloc[-1]**(1/y)-1, Vol=r.std(ddof=1)*np.sqrt(ppy),
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)), MaxDD=dd.min(),
                Calmar=(eq.iloc[-1]**(1/y)-1)/abs(dd.min()), Pct1y=(roll >= 0.12).mean() if len(roll) else np.nan)


def capnorm(w, cap, total):
    """scale dict of raw weights to sum=total, cap each at `cap`, renormalise."""
    if not w:
        return {}
    s = sum(w.values()); w = {k: v/s*total for k, v in w.items()}
    w = {k: min(v, cap) for k, v in w.items()}
    s = sum(w.values())
    return {k: v/s*total for k, v in w.items()} if s > 0 else {}


def run(out=SCR):
    tags, maxsev, moat, country = load_meta()
    dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet")
    dec["Instrument"] = dec["Instrument"].astype(str); dec["Date"] = pd.to_datetime(dec["Date"]).astype("datetime64[ns]")
    px = pd.concat([pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"]) for f in PRICE_FILES])
    px["Instrument"] = px["Instrument"].astype(str); px["Date"] = pd.to_datetime(px["Date"]).astype("datetime64[ns]")
    px["c"] = pd.to_numeric(px["Close Price"], errors="coerce")
    W = px[px.c > 0].pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()
    vol = W.pct_change().rolling(VOL_WIN, min_periods=60).std()*np.sqrt(252)
    s = pd.Series(sorted(dec["Date"].unique())); grid = pd.DatetimeIndex(sorted(s.groupby(s.dt.to_period("Q")).max().values))
    aso = lambda w: w.reindex(grid, method="ffill")
    er_g = aso(dec.pivot_table(index="Date", columns="Instrument", values="er_total", aggfunc="last"))
    ca_g = aso(dec.pivot_table(index="Date", columns="Instrument", values="carry", aggfunc="last"))
    ar_g = aso(dec.pivot_table(index="Date", columns="Instrument", values="artifact", aggfunc="last"))
    vol_g, px_g = aso(vol), aso(W)

    rets, dates, holds, satsz, coresz = [], [], {}, [], []
    for i in range(len(grid)-1):
        t, t1 = grid[i], grid[i+1]
        erT, caT, vT, arT, p0, p1 = er_g.loc[t], ca_g.loc[t], vol_g.loc[t], ar_g.loc[t], px_g.loc[t], px_g.loc[t1]
        base = []
        for r in erT.index:
            e = erT.get(r); v = vT.get(r)
            if r not in moat or moat[r] < CORE_MOAT or maxsev.get(r, 0) >= 4:
                continue
            if pd.isna(e) or e < CORE_GATE or bool(arT.get(r)) or pd.isna(v) or v <= 0 or pd.isna(p0.get(r)):
                continue
            base.append(r)
        # SATELLITE: er>=18%, carry>0, top by er
        sat_e = sorted([r for r in base if erT[r] >= SAT_GATE and (caT.get(r) or 0) > 0], key=lambda r: -erT[r])[:SAT_N]
        satw = capnorm(interpolate_weights({r: float(erT[r]) for r in sat_e}), SAT_WMAX, SAT_FRAC) if sat_e else {}
        cfrac = 1.0 - sum(satw.values())
        # CORE: rank by er/sqrt(vol); fill under combined slot cap (+sat tags) + country cap
        buckets, cc = {}, {}
        for r in sat_e:
            for tg in tags.get(r, ()): buckets[tg] = buckets.get(tg, 0)+1
            cc[country.get(r)] = cc.get(country.get(r), 0)+1
        core_rank = sorted([r for r in base if satw.get(r, 0) <= 0], key=lambda r: -erT[r]/vT[r])
        core = []
        for r in core_rank:
            if len(core) >= CORE_N: break
            ts = tags.get(r, frozenset()); co = country.get(r)
            if any(buckets.get(x, 0)+1 > SLOT_CAP for x in ts): continue
            if co is not None and cc.get(co, 0)+1 > COUNTRY_CAP: continue
            core.append(r)
            for x in ts: buckets[x] = buckets.get(x, 0)+1
            cc[co] = cc.get(co, 0)+1
        corew = capnorm({r: 1.0/float(vT[r]) for r in core}, CORE_WMAX, cfrac) if core else {}
        w = {**satw, **corew}
        # realized quarter return
        ret = {r: (float(p1[r])/float(p0[r])-1) for r in w if pd.notna(p0.get(r)) and pd.notna(p1.get(r)) and float(p0[r]) > 0}
        tw = sum(w[r] for r in ret)
        if tw <= 0: continue
        rets.append(sum(w[r]*ret[r] for r in ret)/tw); dates.append(t1)
        holds[str(pd.Timestamp(t).date())] = {"satellite": {r: round(satw[r], 3) for r in satw},
                                              "core": sorted(core)}
        satsz.append(len(sat_e)); coresz.append(len(core))
    idx = pd.DatetimeIndex(dates)
    unhedged = pd.Series(rets, index=idx, name="System")

    # ---- external MSCI World IMI put hedge ----
    m = pd.read_parquet(SCR+"/msci_imi_monthly.parquet").set_index("date")["ret"]
    mq = ((1+m).cumprod().reindex(idx, method="ffill")); mq = (mq/mq.shift(1)-1).fillna(0.0)
    b, u = unhedged.values, mq.values
    beta = np.cov(b, u)[0, 1]/np.var(u); sig = np.std(u, ddof=1)*np.sqrt(4)
    si = sig+VRP+SKEW*(1-HEDGE_MONEY)/0.10
    S = np.cumprod(np.r_[1.0, 1+u]); E = 1.0; pv = 0.0; con = None; nav = [1.0]
    for q in range(len(b)):
        if q % HEDGE_TN == 0:
            if con: E += pv; pv = 0.0; con = None
            N = beta*E; K = HEDGE_MONEY*S[q]; cost = bsput(S[q], K, si, RF, HEDGE_TN/4.0)/S[q]*N
            E -= cost; pv = cost; con = (K, q, N, S[q])
        E *= (1+b[q])
        if con:
            K, q0, N, S0 = con; tau = max((q0+HEDGE_TN-(q+1))/4.0, 0.0); pv = bsput(S[q+1], K, si, RF, tau)/S0*N
        nav.append(E+pv)
    hedged = pd.Series(np.array(nav)[1:]/np.array(nav)[:-1]-1, index=idx, name="System+hedge")

    # ---- report ----
    bench = {"Universe EW": pd.read_parquet(SCR+"/bt_returns_Q.parquet")["Universe_EW"],
             "Framework slot-cap": pd.read_parquet(SCR+"/bt_returns_Q.parquet")["Core30_framework"],
             "Resilient Core-30": pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]}
    allr = {"System (Core+Sat, unhedged)": unhedged, "System + MSCI put hedge": hedged, **bench}
    st = pd.DataFrame({k: perf(v) for k, v in allr.items()}).T
    for c in ["CAGR", "Vol", "MaxDD", "Pct1y"]: st[c] = (st[c]*100).round(1)
    for c in ["Sharpe", "Calmar"]: st[c] = st[c].round(2)
    print(f"[build] {len(unhedged)} quarters | avg satellite {np.mean(satsz):.1f}, avg core {np.mean(coresz):.1f}, "
          f"book beta to MSCI {beta:.2f}")
    print(st[["CAGR", "Vol", "Sharpe", "MaxDD", "Calmar", "Pct1y"]].to_string())
    unhedged.to_frame().to_parquet(out+"/bt_system_Q.parquet")
    hedged.to_frame().to_parquet(out+"/bt_system_hedged_Q.parquet")
    json.dump(holds, open(out+"/bt_system_holds.json", "w"))

    import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
    R = pd.DataFrame(allr).dropna(); eq = (1+R).cumprod()
    col = {"System (Core+Sat, unhedged)": "#ff7f0e", "System + MSCI put hedge": "#1f77b4",
           "Resilient Core-30": "#2ca02c", "Framework slot-cap": "#9467bd", "Universe EW": "#7f7f7f"}
    fig, (a1, a2) = plt.subplots(2, 1, figsize=(11, 8.5), height_ratios=[3, 1.5], sharex=True)
    for k in R.columns:
        a1.plot(eq.index, eq[k], label=k, color=col[k], lw=2.4 if "System" in k else 1.5)
    yrs = np.arange(len(eq)); a1.plot(eq.index, 1.12**(yrs/4.), ls="--", color="k", lw=0.8, label="12% target")
    a1.set_yscale("log"); a1.legend(loc="upper left", fontsize=9); a1.grid(True, which="both", alpha=0.25); a1.set_ylabel("Growth of 1.0 (log)")
    a1.set_title("Integrated system: Core (75%) + Satellite (25%) + MSCI World IMI put hedge\nvs benchmarks (quarterly, full engine)", fontsize=11)
    for k in ["System + MSCI put hedge", "System (Core+Sat, unhedged)"]:
        e = eq[k]; dd = e/e.cummax()-1; a2.plot(dd.index, dd*100, color=col[k], lw=1.7, label=k)
    a2.grid(True, alpha=0.25); a2.set_ylabel("drawdown (%)"); a2.legend(fontsize=8)
    plt.tight_layout(); fig.savefig(out+"/backtest_system.png", dpi=130); print("wrote backtest_system.png")
    return st, holds


if __name__ == "__main__":
    run()
