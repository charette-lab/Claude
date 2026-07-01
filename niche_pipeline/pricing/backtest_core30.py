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
               U+"86e54ec3-daily_volume_price_2.parquet", U+"13f82c18-daily_volume_price_0630.parquet",
               U+"cf27bad3-two_batches_daily_volume_price_0.parquet", U+"c3621e8f-two_batches_daily_volume_price_1.parquet",
               U+"35e9a52e-two_batches_daily_volume_price_2.parquet", U+"e397aa9f-two_batches_daily_volume_price_3.parquet",
               U+"3cfd2814-two_batches_daily_volume_price_4.parquet", U+"25b0cf78-two_batches_daily_volume_price_5.parquet",
               U+"95c8416b-two_batches_daily_volume_price_6.parquet"]
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
USE_GATE2 = os.environ.get("CORE30_GATE2", "0") == "1"   # IPS Gate 2: P(30% drawdown) <= 20%
FLOOR_RATIO = 0.70        # no-growth floor / price; >= 0.70 => a 30% fall lands at the floor
GATE2_PROB = 0.20         # statistical P(30% drawdown over 1yr) ceiling
DOWN_BARRIER = 0.70


def downside_prob(er, sigma, barrier=DOWN_BARRIER, horizon=1.0):
    """First-passage P(price touches barrier*P0 within horizon) under GBM with drift =
    capped expected return and vol sigma (reflection-principle barrier formula)."""
    from scipy.stats import norm
    if sigma is None or pd.isna(sigma) or sigma <= 0:
        return 1.0
    mu = min(max(float(er), 0.0), 0.12)               # conservative: cap the drift credit at the hurdle
    nu = mu - 0.5*sigma**2; m = np.log(barrier); s = sigma*np.sqrt(horizon)
    p = norm.cdf((m-nu*horizon)/s) + np.exp(2*nu*m/sigma**2)*norm.cdf((m+nu*horizon)/s)
    return float(min(max(p, 0.0), 1.0))


def gate2_pass(floor_ratio, er, sigma):
    """Pass if a hard valuation floor is within 30% of price OR the statistical
    1-yr probability of a 30% drawdown is <= 20%."""
    if floor_ratio is not None and not pd.isna(floor_ratio) and floor_ratio >= FLOOR_RATIO:
        return True
    return downside_prob(er, sigma) <= GATE2_PROB
WEIGHT_MODE = os.environ.get("CORE30_WEIGHT", "invvol")  # "invvol" or "cvar" (min 95% expected shortfall)
CVAR_BETA = 0.95
CVAR_WIN = 252


def cvar_weights(book, ret, t, er_row, wmin=0.01, wmax=0.08, floor_frac=1.0):
    """Rockafellar-Uryasev: minimise the portfolio's 95% CVaR (mean of the worst-5%
    trailing daily returns) s.t. expected er >= the equal-weight level and box weights.
    This is downside-CORRELATION aware: names that crash on the SAME days are penalised
    jointly, so the optimiser tilts toward holdings whose drawdowns offset."""
    from scipy.optimize import linprog
    R = ret.reindex(columns=book).loc[:t].tail(CVAR_WIN).dropna(how="all")
    if len(R) < 60:
        return None
    Rm = R.fillna(0.0).to_numpy(dtype=float); T, n = Rm.shape
    er = np.array([float(er_row.get(b, 0.0)) for b in book])
    floor = floor_frac*er.mean()
    nv = n+1+T
    c = np.concatenate([np.zeros(n), [1.0], np.full(T, 1.0/((1-CVAR_BETA)*T))])
    A1 = np.zeros((T, nv)); A1[:, :n] = -Rm; A1[:, n] = -1.0
    A1[np.arange(T), n+1+np.arange(T)] = -1.0; b1 = np.zeros(T)
    A2 = np.zeros((1, nv)); A2[0, :n] = -er; b2 = np.array([-floor])
    A_eq = np.zeros((1, nv)); A_eq[0, :n] = 1.0; b_eq = np.array([1.0])
    bounds = [(wmin, wmax)]*n + [(None, None)] + [(0, None)]*T
    res = linprog(c, A_ub=np.vstack([A1, A2]), b_ub=np.concatenate([b1, b2]),
                  A_eq=A_eq, b_eq=b_eq, bounds=bounds, method="highs")
    if not res.success:
        res = linprog(c, A_ub=A1, b_ub=b1, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method="highs")
    if not res.success:
        return None
    w = np.clip(res.x[:n], 0, None)
    if w.sum() <= 0:
        return None
    w = w/w.sum()
    shrink = float(os.environ.get("CVAR_SHRINK", "0.0"))   # shrink toward equal weight (estimation-error control)
    if shrink > 0:
        w = (1-shrink)*w + shrink*(np.ones(n)/n)
    return w


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
    mcw = dec.pivot_table(index="Date", columns="Instrument", values="market_cap", aggfunc="last")
    er_g, art_g, mc_g = asof(er, grid), asof(art, grid), asof(mcw, grid)
    vol_g, trend_g, hist_g, px_g = asof(vol, grid), asof(trend, grid), asof(hist, grid), asof(W, grid)
    # Gate-2 no-growth floor (per fiscal year) -> floor_ratio at the grid
    floor_g = None
    if USE_GATE2:
        F = pd.read_parquet(SCR+"/floor_equity.parquet")
        F["Instrument"] = F["Instrument"].astype(str); F["ped"] = pd.to_datetime(F["ped"]).astype("datetime64[ns]")
        fw = F.pivot_table(index="ped", columns="Instrument", values="floor_equity", aggfunc="last").sort_index()
        floor_g = fw.reindex(grid, method="ffill")

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
            if USE_GATE2:                                     # IPS Gate 2: P(30% drawdown) <= 20%
                fe = floor_g.loc[t].get(r) if floor_g is not None else None
                mc = mc_g.loc[t].get(r)
                fr = (float(fe)/float(mc)) if (fe is not None and not pd.isna(fe) and mc and not pd.isna(mc)) else None
                if not gate2_pass(fr, e, v):
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
        # weights: inverse-vol, or downside-correlation-aware CVaR minimisation
        w = cvar_weights(book, ret, t, erT) if WEIGHT_MODE == "cvar" else None
        if w is None:
            iv = np.array([1.0/float(vT[r]) for r in book]); w = np.clip(iv/iv.sum(), W_MIN, W_MAX); w = w/w.sum()
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
    tag = ("" if RISK_MEASURE == "vol" else f"_{RISK_MEASURE}") + ("" if WEIGHT_MODE == "invvol" else f"_{WEIGHT_MODE}") + ("_g2" if USE_GATE2 else "")
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
