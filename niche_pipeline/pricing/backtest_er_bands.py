#!/usr/bin/env python3
"""backtest_er_bands.py — pure-ER Core index run on the IPS rebalancing machinery.

Per instruction: the SELECTOR is Expected IRR alone (risk tags ignored — tag
coverage is incomplete), but the rest of the framework drives maintenance. So we
keep the full ER universe (all priced securities, no tag requirement) and run the
IPS Phase-4 "Quarterly Checkup with Tolerance Bands":

  * Target  : 30 names, equal-weighted at 3.33%.
  * Gate 1  : Expected IRR >= 12% to be eligible / to replace.
  * Hold with drift; act only when a band is breached (no calendar churn):
      Trigger 2 (Momentum cap): a name drifts to >= 5.0% -> trim back to 3.33%,
                 redeploy the proceeds into highest-ER holdings below target.
      Trigger 3 (Thesis violation): a name bleeds to <= 2.0% -> liquidate and
                 replace with the highest-ER name not held (from the screened
                 universe), at 3.33%.
  * Trigger 1 (risk-tag ceiling) is DROPPED — that is the "ignore risk factors"
    part; everything else of the framework is retained.

Selection has no look-ahead (ER@t uses price@t + prior FY fundamentals); future
prices only mark the book. Returns are close-to-close, local ccy, price-only, no
costs (same caveats as backtest.py).

  python3 pricing/backtest_er_bands.py --cadence Q
"""
from __future__ import annotations
import argparse, json, os, sys
import numpy as np, pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from backtest import (ER_PATH, SCRATCH, load_close, wide_on_grid, rebalance_dates,
                      perf_stats)
from frameworks import GATE1_IRR, CORE_N

TARGET = 1.0 / CORE_N      # 3.33%
HI_BAND = 0.05             # Trigger 2 momentum cap
LO_BAND = 0.02             # Trigger 3 thesis violation


def simulate(er_w, px_w, grid, n=CORE_N, target=TARGET, gate=GATE1_IRR,
             hi=HI_BAND, lo=LO_BAND):
    shares, prev_nav = {}, 1.0
    rets, dates, holdings, turn, trig = [], [], {}, [], []
    for i, t in enumerate(grid):
        px, er = px_w.loc[t], er_w.loc[t]
        val = {nm: shares[nm] * px[nm] for nm in shares if pd.notna(px.get(nm))}
        nav = sum(val.values()) if shares else prev_nav
        if i > 0:
            rets.append(nav / prev_nav - 1.0); dates.append(t)
        if nav <= 0:
            break
        w = {nm: val[nm] / nav for nm in val}
        elig = er[er >= gate].dropna()
        elig = elig[px.reindex(elig.index).notna().values]
        ranked = list(elig.sort_values(ascending=False).index)
        before = set(shares)
        n_liq = n_trim = 0
        cash = nav - sum(val.values())     # uninvested NAV (=nav at inception, ~0 in steady state)
        # Trigger 3: liquidate bleeders (and any holding that lost its price)
        for nm in list(shares):
            if nm not in w or w[nm] <= lo:
                cash += val.get(nm, 0.0); del shares[nm]; n_liq += 1
        # Trigger 2: trim winners back to target
        for nm in list(shares):
            cw = shares[nm] * px[nm] / nav
            if cw >= hi:
                excess = (cw - target) * nav
                shares[nm] -= excess / px[nm]; cash += excess; n_trim += 1
        # fill vacated slots with highest-ER names not held
        held = set(shares)
        for nm in ranked:
            if len(shares) >= n or cash <= 1e-12:
                break
            if nm in held:
                continue
            buy = min(target * nav, cash)
            shares[nm] = buy / px[nm]; cash -= buy; held.add(nm)
        # redeploy remaining cash into below-target holdings, highest-ER first
        if cash > 1e-12:
            order = [nm for nm in ranked if nm in shares] + \
                    [nm for nm in shares if nm not in ranked]
            for nm in order:
                cw = shares[nm] * px[nm] / nav
                if cw < target:
                    add = min((target - cw) * nav, cash)
                    shares[nm] += add / px[nm]; cash -= add
                if cash <= 1e-12:
                    break
        # any residual cash -> stay fully invested (spread pro-rata)
        if cash > 1e-12 and shares:
            tot = sum(shares[nm] * px[nm] for nm in shares)
            for nm in shares:
                shares[nm] += cash * (shares[nm] * px[nm] / tot) / px[nm]
        after = set(shares)
        turn.append((len(after - before) + len(before - after)) / 2.0)
        trig.append((n_liq, n_trim))
        holdings[str(pd.Timestamp(t).date())] = sorted(shares)
        prev_nav = nav
    return dict(rets=rets, dates=dates, holdings=holdings, turn=turn, trig=trig)


def load_moats():
    """moat (CompanyMoat v3.2) per RIC from the unified scored book."""
    import openpyxl
    ws = openpyxl.load_workbook(os.path.join(SCRATCH, "Universe_final.xlsx"),
                                data_only=True, read_only=True)["Scored"]
    rows = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rows[0])}
    return {str(r[ci["Ticker"]]): r[ci["CompanyMoat(v3.2)"]] for r in rows[1:]
            if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}


def run(cadence="Q", out=SCRATCH, min_moat=None):
    ppy = {"M": 12, "Q": 4, "A": 1}[cadence]
    print(f"[load] ER panel + prices (cadence={cadence})", flush=True)
    er = pd.read_parquet(ER_PATH)
    er["Instrument"] = er["Instrument"].astype(str)
    er["Date"] = er["Date"].astype("datetime64[ns]")
    px = load_close()
    universe = sorted(set(er["Instrument"]) & set(px["Instrument"]))
    if min_moat is not None:
        moats = load_moats()
        universe = sorted(t for t in universe if moats.get(t, 0) > min_moat)
        print(f"[moat] restricting to moat > {min_moat}: {len(universe)} securities", flush=True)
    print(f"[universe] {len(universe)} priced ER securities (NO tag filter)", flush=True)

    grid = rebalance_dates(er["Date"].unique(), cadence)
    grid = [d for d in grid if d >= er["Date"].min()]
    er_w = wide_on_grid(er.rename(columns={"expected_return": "v"}), "v", grid, universe)
    px_w = wide_on_grid(px.rename(columns={"close": "v"}), "v", grid, universe)
    grid = list(er_w.index)
    print(f"[grid] {len(grid)} rebalance dates {pd.Timestamp(grid[0]).date()} -> {pd.Timestamp(grid[-1]).date()}", flush=True)

    sim = simulate(er_w, px_w, grid)
    idx = pd.DatetimeIndex(sim["dates"])
    ret = pd.Series(sim["rets"], index=idx, name="ER_bands")
    stats = perf_stats(ret, ppy)

    turn = np.array(sim["turn"][1:])           # exclude the initial fill
    liq = np.array([a for a, _ in sim["trig"]][1:])
    trim = np.array([b for _, b in sim["trig"]][1:])
    avg_book = np.mean([len(v) for v in sim["holdings"].values()])
    notrade = int(((liq == 0) & (trim == 0)).sum())

    print("\n==== PURE-ER CORE INDEX (IPS bands, risk tags ignored) ====")
    print(f"  period: {idx[0].date()} -> {idx[-1].date()}  ({len(ret)} {cadence}-periods)")
    for k in ["CAGR", "Vol", "Sharpe", "MaxDD", "Pct_1y_ge_12pct", "FinalNAV"]:
        v = stats[k]
        print(f"  {k:18s}: {v*100:7.2f}%" if k not in ("Sharpe", "FinalNAV")
              else f"  {k:18s}: {v:7.2f}")
    print(f"\n  REBALANCING ACTIVITY:")
    print(f"  avg book size              : {avg_book:.1f}")
    print(f"  avg names changed / checkup: {turn.mean():.1f}  ({turn.mean()/avg_book*100:.0f}% of book)")
    print(f"  one-way turnover           : ~{turn.mean()*ppy:.0f} names/yr ({turn.mean()*ppy/avg_book*100:.0f}% of book/yr)")
    print(f"  avg liquidations (2% band) : {liq.mean():.2f} / checkup")
    print(f"  avg trims (5% band)        : {trim.mean():.2f} / checkup")
    print(f"  checkups with NO trades    : {notrade} of {len(turn)}")
    print(f"  implied avg holding period : ~{avg_book/(turn.mean()*ppy)*12:.0f} months" if turn.mean() > 0 else "")

    sfx = f"_{cadence}" + (f"_moat{min_moat:g}" if min_moat is not None else "")
    ret.to_frame().to_parquet(os.path.join(out, f"bt_erbands_returns{sfx}.parquet"))
    json.dump(sim["holdings"], open(os.path.join(out, f"bt_erbands_holdings{sfx}.json"), "w"))
    pd.Series(stats).to_csv(os.path.join(out, f"bt_erbands_stats{sfx}.csv"))
    return ret, stats, sim


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--cadence", default="Q", choices=["M", "Q", "A"])
    ap.add_argument("--out", default=SCRATCH)
    ap.add_argument("--min-moat", type=float, default=None)
    a = ap.parse_args()
    run(a.cadence, a.out, a.min_moat)
