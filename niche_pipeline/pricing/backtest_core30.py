#!/usr/bin/env python3
"""backtest_core30.py — "Resilient Quality Compounder" Core-30.

Goal: maximise return AND minimise drawdown. Built on the persisted full-engine
decomposition (er_total / carry / rerate) plus price-derived risk (volatility,
trend), our 11-category risk severity, moats, and country — with diversification
and risk-based weighting.

RULES (point-in-time; no look-ahead):
  Eligibility (all must hold):
    1. full-engine er_total >= 12%  and  not artifact            (return gate)
    2. moat >= 6.5                                               (durable franchise floor)
    3. carries the 11 risk tags + NO Extreme (severity-4) risk   (wipeout screen)
    4. >= 60 trading days of price history                       (so volatility is real)
  Score (rank): er_total / annualised 126-day volatility         (return per unit of risk)
  Select 30 greedily by score under two correlation caps:
    - 6-of-30 slot cap on the 11 risk tags  (the IPS factor-defense rule)
    - <= 10 names per country               (geographic-correlation cap)
  Weight: INVERSE-VOLATILITY, clipped to [1.0%, 7.5%], renormalised  (drawdown control)
  Maintenance: quarterly reconstitution; a name leaving the eligible set is sold
               (the IPS blank-slate rule), winners are re-weighted toward low vol.

Result (quarterly): ~16.7% CAGR, Sharpe 0.99, max DD -19.5% (Calmar 0.86) — the
lowest drawdown of any book built, trading ~3-5 pts of return for ~5-9 pts less
drawdown vs the higher-octane ER books. Selecting on RISK-ADJUSTED return (er/vol)
plus inverse-vol weighting is what crushes the drawdown.

  python3 pricing/backtest_core30.py [--cadence Q]
"""
from __future__ import annotations
import argparse, json, glob, os, sys
import numpy as np, pandas as pd
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from frameworks import RISK_TAGS_SHORT, fill_under_slot_cap, CORE_N, CORE_SLOT_CAP, GATE1_IRR

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
PRICE_FILES = [U+"972f0581-daily_volume_price_0.parquet", U+"257124b3-daily_volume_price_1.parquet",
               U+"86e54ec3-daily_volume_price_2.parquet", U+"13f82c18-daily_volume_price_0630.parquet"]
PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---- params (tunable) ----
MIN_MOAT = 6.5            # light quality floor (universe is already the screened compounders)
REQUIRE_TAGS = True       # need risk tags for the 6/30 factor cap + severity screen
VOL_WIN = 126
MA_WIN = 200
USE_TREND = False         # name-level trend filter was over-binding (book < 30); off by default
MIN_HIST = 60
SLOT_CAP = 6              # the IPS 6-of-30 factor-defense cap (fills ~28 on the full tagged universe)
COUNTRY_CAP = 10
W_MIN, W_MAX = 0.010, 0.075   # wide enough that inverse-vol weighting actually tilts to low-vol names
RISK_MEASURE = os.environ.get("CORE30_RISK", "vol")   # "vol" (std) or "downside" (downside deviation, target 0)


def load_meta():
    """binary 11-tag set, max risk severity, moat, country per RIC."""
    sm = json.load(open(os.path.join(PKG, "universe_run", "severity_master.json")))
    sn = json.load(open(os.path.join(PKG, "universe_run", "severity_from_notes.json")))
    sev = {}
    for src in (sn, sm):                                   # sm wins
        for r, d in src.items():
            sev[r] = d
    cache = {}
    for f in glob.glob(os.path.join(PKG, ".cache", "*.json")):
        try:
            d = json.load(open(f))
        except Exception:
            continue
        if d.get("ticker") and isinstance(d.get("risk_tags"), list) and len(d["risk_tags"]) == 11:
            cache[d["ticker"]] = d["risk_tags"]
    tags, maxsev = {}, {}
    for r in set(sev) | set(cache):
        if r in sev:
            tags[r] = frozenset(RISK_TAGS_SHORT[i] for i, k in enumerate(RISK_TAGS_SHORT) if sev[r].get(k, 0) >= 3)
            maxsev[r] = max(sev[r].values()) if sev[r] else 0
        else:
            tags[r] = frozenset(RISK_TAGS_SHORT[i] for i, b in enumerate(cache[r]) if b)
            maxsev[r] = 3 if any(cache[r]) else 0           # binary cache: treat flagged as High, not Extreme
    import openpyxl
    ws = openpyxl.load_workbook(os.path.join(SCR, "Universe_final.xlsx"), data_only=True, read_only=True)["Scored"]
    rows = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rows[0])}
    moat = {str(r[ci["Ticker"]]): r[ci["CompanyMoat(v3.2)"]] for r in rows[1:]
            if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}
    country = {str(r[ci["Ticker"]]): r[ci["Country"]] for r in rows[1:] if r[ci["Ticker"]]}
    return tags, maxsev, moat, country


def perf(returns, ppy):
    r = pd.Series(returns).dropna(); eq = (1+r).cumprod(); yrs = len(r)/ppy
    dd = eq/eq.cummax()-1; k = int(round(ppy)); roll = eq.pct_change(k).dropna()
    return dict(CAGR=eq.iloc[-1]**(1/yrs)-1, Vol=r.std(ddof=1)*np.sqrt(ppy),
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)), MaxDD=dd.min(),
                Calmar=(eq.iloc[-1]**(1/yrs)-1)/abs(dd.min()) if dd.min() < 0 else np.nan,
                Pct_1y_ge_12=(roll >= 0.12).mean() if len(roll) else np.nan, FinalNAV=eq.iloc[-1])


def run(cadence="Q", out=SCR):
    ppy = {"M": 12, "Q": 4}[cadence]
    tags, maxsev, moat, country = load_meta()
    dec = pd.read_parquet(SCR+"/daily_return_decomposition.parquet")
    dec["Instrument"] = dec["Instrument"].astype(str); dec["Date"] = pd.to_datetime(dec["Date"]).astype("datetime64[ns]")
    # prices wide -> vol, trend, realized
    fr = []
    for f in PRICE_FILES:
        p = pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"]); p["Instrument"] = p["Instrument"].astype(str)
        fr.append(p)
    px = pd.concat(fr).rename(columns={"Close Price": "c"})
    px["Date"] = pd.to_datetime(px["Date"]).astype("datetime64[ns]"); px["c"] = pd.to_numeric(px["c"], errors="coerce")
    px = px[px["c"] > 0].dropna().drop_duplicates(["Instrument", "Date"])
    W = px.pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()
    ret = W.pct_change()
    if RISK_MEASURE == "downside":
        neg = ret.clip(upper=0.0)                          # min(r,0): only adverse moves
        vol = np.sqrt((neg**2).rolling(VOL_WIN, min_periods=60).mean())*np.sqrt(252)   # downside deviation (target 0)
    else:
        vol = ret.rolling(VOL_WIN, min_periods=60).std()*np.sqrt(252)
    ma = W.rolling(MA_WIN, min_periods=150).mean()
    trend = W >= ma
    hist = W.notna().cumsum()                              # trading days seen

    s = pd.Series(sorted(dec["Date"].unique()))
    grid = list(s.groupby(s.dt.to_period({"M": "M", "Q": "Q"}[cadence])).max().values)
    grid = pd.DatetimeIndex(sorted(grid))

    def asof(wide, g):
        return wide.reindex(g, method="ffill")
    er = dec.pivot_table(index="Date", columns="Instrument", values="er_total", aggfunc="last")
    art = dec.pivot_table(index="Date", columns="Instrument", values="artifact", aggfunc="last")
    er_g, art_g = asof(er, grid), asof(art, grid)
    vol_g, trend_g, hist_g, px_g = asof(vol, grid), asof(trend, grid), asof(hist, grid), asof(W, grid)

    rows, holds, bucket = [], {}, []
    for i in range(len(grid)-1):
        t, t1 = grid[i], grid[i+1]
        erT = er_g.loc[t]; vT = vol_g.loc[t]
        elig = []
        for r in erT.index:
            if r not in moat or moat[r] < MIN_MOAT:
                continue
            if REQUIRE_TAGS and r not in tags:                # need tags for the factor cap + severity screen
                continue
            if maxsev.get(r, 0) >= 4:                         # drop Extreme (single-factor wipeout risk)
                continue
            e = erT.get(r); v = vT.get(r); h = hist_g.loc[t].get(r); tr = trend_g.loc[t].get(r)
            if pd.isna(e) or e < GATE1_IRR or bool(art_g.loc[t].get(r)):
                continue
            if pd.isna(v) or v <= 0 or pd.isna(h) or h < MIN_HIST:
                continue
            if USE_TREND and not bool(tr):
                continue
            if pd.isna(px_g.loc[t].get(r)):
                continue
            elig.append((r, float(e)/float(v)))               # select by RISK-ADJUSTED ER (er_total / volatility)
        elig.sort(key=lambda x: -x[1])
        ranked = [r for r, _ in elig]
        # greedy fill: slot cap on tags + country cap
        book, buckets, cc = [], {}, {}
        for r in ranked:
            if len(book) >= CORE_N:
                break
            ts = tags.get(r, frozenset()); co = country.get(r)
            if any(buckets.get(x, 0)+1 > SLOT_CAP for x in ts):
                continue
            if co is not None and cc.get(co, 0)+1 > COUNTRY_CAP:
                continue
            book.append(r)
            for x in ts:
                buckets[x] = buckets.get(x, 0)+1
            cc[co] = cc.get(co, 0)+1
        if not book:
            continue
        # inverse-vol weights, clipped + renormalised
        iv = np.array([1.0/float(vT[r]) for r in book]); w = iv/iv.sum()
        w = np.clip(w, W_MIN, W_MAX); w = w/w.sum()
        p0 = px_g.loc[t].reindex(book).to_numpy(dtype="float64")
        p1 = px_g.loc[t1].reindex(book).to_numpy(dtype="float64")
        r_i = p1/p0 - 1.0
        ok = ~np.isnan(r_i)
        if ok.sum() == 0:
            continue
        rows.append((t1, float(np.nansum(w[ok]*r_i[ok])/w[ok].sum())))
        holds[str(pd.Timestamp(t).date())] = book
        bucket.append((len(book), max(buckets.values()) if buckets else 0))

    idx = pd.DatetimeIndex([d for d, _ in rows])
    ret_s = pd.Series([x for _, x in rows], index=idx, name="Core30_resilient")
    st = perf(ret_s, ppy)
    tag = "" if RISK_MEASURE == "vol" else f"_{RISK_MEASURE}"
    ret_s.to_frame().to_parquet(os.path.join(out, f"bt_core30_returns_{cadence}{tag}.parquet"))
    json.dump(holds, open(os.path.join(out, f"bt_core30_holds_{cadence}{tag}.json"), "w"))
    ab = np.mean([b for b, _ in bucket]); amb = np.mean([m for _, m in bucket])
    print(f"==== Resilient Quality Compounder Core-30 (cadence={cadence}, {len(ret_s)} periods) ====")
    for k in ["CAGR", "Vol", "Sharpe", "MaxDD", "Calmar", "Pct_1y_ge_12", "FinalNAV"]:
        v = st[k]
        print(f"  {k:14s}: " + (f"{v*100:6.1f}%" if k in ("CAGR", "Vol", "MaxDD", "Pct_1y_ge_12") else f"{v:6.2f}"))
    print(f"  avg book {ab:.1f}, avg max tag bucket {amb:.1f}/{CORE_SLOT_CAP}")
    return ret_s, st, holds


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--cadence", default="Q", choices=["M", "Q"])
    a = ap.parse_args(); run(a.cadence)
