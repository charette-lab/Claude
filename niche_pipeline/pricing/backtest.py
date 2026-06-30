#!/usr/bin/env python3
"""backtest.py — evaluate the Constrained Quality Compounder Index through time.

Builds the framework's Core index on a rebalance grid and measures the REALIZED
forward return, to answer: does the system actually work over time?

The system (per the IPS), Gate 2 (downside hard-stop) DISABLED per instruction:
  * Universe : the high-moat researched names that carry the 11 binary risk tags
               (a stock must be tagged before purchase, so untagged names are out).
  * Gate 1   : forward Expected IRR >= 12%  (from the daily ROIIC-engine ER panel).
  * Rank     : by Expected IRR, best first.
  * Build    : 30 equal-weighted names under the 20% factor rule expressed as the
               6-of-30 slot cap (no risk tag held by more than 6 of the 30 names).
  * Sweep    : capital not used by the 30 names is irrelevant here — the Core IS
               the 30-name index; we evaluate that index's realized return.

Signal has no look-ahead: ER at date t uses price at t and the prior fiscal-year
fundamentals; realized return uses prices AFTER t only to score, never to select.

Realized return is close-to-close (price only; dividends not in the data, so total
return is modestly understated). A name with no print at the next rebalance carries
its last close (0% that period) — mildly optimistic for delistings; flagged below.

  python3 pricing/backtest.py --cadence M --out <dir>
"""
from __future__ import annotations
import argparse, json, glob, os, sys
import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from frameworks import fill_under_slot_cap, RISK_TAGS_SHORT, CORE_N, CORE_SLOT_CAP, GATE1_IRR

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
SCRATCH = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
UPLOADS = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70"
ER_PATH = os.path.join(SCRATCH, "daily_expected_return.parquet")
PRICE_FILES = [
    os.path.join(UPLOADS, "972f0581-daily_volume_price_0.parquet"),
    os.path.join(UPLOADS, "257124b3-daily_volume_price_1.parquet"),
    os.path.join(UPLOADS, "86e54ec3-daily_volume_price_2.parquet"),
    os.path.join(UPLOADS, "13f82c18-daily_volume_price_0630.parquet"),
]
SEV_THRESHOLD = 3   # severity >= 3 (High/Extreme) == binary "material exposure"


# ---------------------------------------------------------------- tag map ----
def build_tagmap(instruments):
    """binary 11-tag vector per RIC: severity_master>=3, else notes>=3, else
    the research cache's native binary tags. Returns {ric: frozenset(short tags)}."""
    sm = json.load(open(os.path.join(PKG, "universe_run", "severity_master.json")))
    sn = json.load(open(os.path.join(PKG, "universe_run", "severity_from_notes.json")))
    cache = {}
    for f in glob.glob(os.path.join(PKG, ".cache", "*.json")):
        try:
            d = json.load(open(f))
        except Exception:
            continue
        t, rt = d.get("ticker"), d.get("risk_tags")
        if t and isinstance(rt, list) and len(rt) == 11:
            cache[t] = rt

    def vec(ric):
        if ric in sm:
            return [1 if sm[ric].get(k, 0) >= SEV_THRESHOLD else 0 for k in RISK_TAGS_SHORT]
        if ric in sn:
            return [1 if sn[ric].get(k, 0) >= SEV_THRESHOLD else 0 for k in RISK_TAGS_SHORT]
        if ric in cache:
            return cache[ric]
        return None

    out = {}
    for ric in instruments:
        v = vec(ric)
        if v is not None:
            out[ric] = frozenset(RISK_TAGS_SHORT[i] for i, b in enumerate(v) if b)
    return out


# ------------------------------------------------------------ price/ER grid --
def load_close():
    fr = []
    for f in PRICE_FILES:
        p = pd.read_parquet(f, columns=["Instrument", "Date", "Close Price"])
        fr.append(p)
    px = pd.concat(fr, ignore_index=True)
    px["Instrument"] = px["Instrument"].astype(str)
    px["Date"] = px["Date"].astype("datetime64[ns]")
    px["close"] = pd.to_numeric(px["Close Price"], errors="coerce")
    px = px[(px["close"] > 0)].dropna(subset=["close"])
    return px[["Instrument", "Date", "close"]]


def rebalance_dates(all_dates, cadence):
    """Last available trading day in each period (M=month, Q=quarter, A=year)."""
    s = pd.Series(all_dates).sort_values().drop_duplicates()
    grp = {"M": s.dt.to_period("M"), "Q": s.dt.to_period("Q"), "A": s.dt.to_period("Y")}[cadence]
    return list(s.groupby(grp).max().values)


def wide_on_grid(long_df, value, grid, instruments):
    """Pivot a long (Instrument,Date,value) frame to Date x Instrument, then
    as-of (ffill) onto the rebalance grid so every grid date has the latest known
    value for each instrument."""
    sub = long_df[long_df["Instrument"].isin(instruments)]
    w = sub.pivot_table(index="Date", columns="Instrument", values=value, aggfunc="last")
    w = w.sort_index().ffill()
    grid = pd.DatetimeIndex(sorted(pd.DatetimeIndex(grid)))
    return w.reindex(grid, method="ffill")


# ------------------------------------------------------------- construction --
def build_core(er_row, tagmap, n=CORE_N, slot_cap=CORE_SLOT_CAP, gate=GATE1_IRR,
               use_slotcap=True, apply_gate=True):
    """Return the list of held RICs for one rebalance date from that date's ER row."""
    er = er_row.dropna()
    if apply_gate:
        er = er[er >= gate]
    ranked = list(er.sort_values(ascending=False).index)
    if not use_slotcap:
        return ranked[:n]
    book, _ = fill_under_slot_cap(ranked, lambda r: tagmap.get(r, frozenset()), n, slot_cap)
    return book


# --------------------------------------------------------------- metrics -----
def perf_stats(returns, periods_per_year):
    r = pd.Series(returns).dropna()
    if len(r) == 0:
        return {}
    eq = (1 + r).cumprod()
    years = len(r) / periods_per_year
    cagr = eq.iloc[-1] ** (1 / years) - 1 if years > 0 and eq.iloc[-1] > 0 else np.nan
    vol = r.std(ddof=1) * np.sqrt(periods_per_year)
    sharpe = (r.mean() * periods_per_year) / vol if vol > 0 else np.nan
    dd = eq / eq.cummax() - 1
    # rolling-1y windows >= 12%
    k = int(round(periods_per_year))
    roll = eq.pct_change(k).dropna() if len(eq) > k else pd.Series(dtype=float)
    pct12 = float((roll >= 0.12).mean()) if len(roll) else np.nan
    return {
        "CAGR": cagr, "Vol": vol, "Sharpe": sharpe, "MaxDD": dd.min(),
        "TotalReturn": eq.iloc[-1] - 1, "Periods": len(r),
        "MeanPeriodRet": r.mean(), "Pct_1y_ge_12pct": pct12,
        "FinalNAV": eq.iloc[-1],
    }


# ----------------------------------------------------------------- runner ----
STRATS = {
    "Core30_framework": dict(use_slotcap=True, apply_gate=True),   # the system
    "Top30_no_slotcap": dict(use_slotcap=False, apply_gate=True),  # gate only, no factor rule
}


def run(cadence="M", out=SCRATCH):
    ppy = {"M": 12, "Q": 4, "A": 1}[cadence]
    print(f"[load] ER panel + prices (cadence={cadence})", flush=True)
    er = pd.read_parquet(ER_PATH)
    er["Instrument"] = er["Instrument"].astype(str)
    er["Date"] = er["Date"].astype("datetime64[ns]")
    instruments = er["Instrument"].unique()
    tagmap = build_tagmap(instruments)
    tagged = set(tagmap)
    print(f"[tags] {len(tagged)} of {len(instruments)} ER securities tagged", flush=True)

    er = er[er["Instrument"].isin(tagged)]
    px = load_close()
    px = px[px["Instrument"].isin(tagged)]

    grid = rebalance_dates(er["Date"].unique(), cadence)
    grid = [d for d in grid if d >= er["Date"].min()]
    print(f"[grid] {len(grid)} rebalance dates {pd.Timestamp(grid[0]).date()} -> {pd.Timestamp(grid[-1]).date()}", flush=True)

    er_w = wide_on_grid(er.rename(columns={"expected_return": "v"}), "v", grid, tagged)
    px_w = wide_on_grid(px.rename(columns={"close": "v"}), "v", grid, tagged)

    grid = list(er_w.index)
    rows = {s: [] for s in STRATS}
    rows["Universe_EW"] = []           # all tagged names with a price, EW (benchmark)
    rows["Gate_EW"] = []               # ALL names with ER>=12%, EW (gate, uncapped count)
    holdings = {s: {} for s in STRATS}
    bucket_log = []

    for i in range(len(grid) - 1):
        t, t1 = grid[i], grid[i + 1]
        er_row = er_w.loc[t]
        px_t, px_t1 = px_w.loc[t], px_w.loc[t1]
        have_px = px_t.dropna().index
        # per-name realized return over [t, t1]
        ret = (px_t1 / px_t - 1.0)

        def book_ret(book):
            b = [r for r in book if r in have_px]
            if not b:
                return np.nan, b
            return float(ret.reindex(b).mean()), b

        for s, kw in STRATS.items():
            # restrict ER row to names that actually have a price at t
            row = er_row.reindex(have_px)
            book = build_core(row, tagmap, **kw)
            pr, b = book_ret(book)
            rows[s].append(pr)
            holdings[s][pd.Timestamp(t).date().isoformat()] = b
            if s == "Core30_framework" and b:
                buckets = {}
                for r in b:
                    for tag in tagmap.get(r, ()):
                        buckets[tag] = buckets.get(tag, 0) + 1
                bucket_log.append((pd.Timestamp(t).date().isoformat(), len(b),
                                   max(buckets.values()) if buckets else 0))
        # benchmarks
        uni = list(have_px)
        rows["Universe_EW"].append(float(ret.reindex(uni).mean()) if uni else np.nan)
        gate_names = list(er_row.reindex(have_px).dropna()[lambda x: x >= GATE1_IRR].index)
        rows["Gate_EW"].append(float(ret.reindex(gate_names).mean()) if gate_names else np.nan)

    # assemble
    idx = pd.DatetimeIndex(grid[:-1])
    retdf = pd.DataFrame({s: rows[s] for s in rows}, index=idx)
    stats = {s: perf_stats(retdf[s], ppy) for s in retdf.columns}
    statdf = pd.DataFrame(stats).T

    os.makedirs(out, exist_ok=True)
    retdf.to_parquet(os.path.join(out, f"bt_returns_{cadence}.parquet"))
    eqdf = (1 + retdf).cumprod()
    eqdf.to_parquet(os.path.join(out, f"bt_equity_{cadence}.parquet"))
    json.dump(holdings, open(os.path.join(out, f"bt_holdings_{cadence}.json"), "w"))
    bl = pd.DataFrame(bucket_log, columns=["date", "n_book", "max_tag_bucket"])

    print("\n==== PERFORMANCE (cadence=%s, %d periods, %s -> %s) ====" %
          (cadence, len(retdf), idx[0].date(), idx[-1].date()))
    show = statdf.copy()
    for c in ["CAGR", "Vol", "MaxDD", "TotalReturn", "MeanPeriodRet", "Pct_1y_ge_12pct"]:
        show[c] = (show[c] * 100).round(2)
    show["Sharpe"] = show["Sharpe"].round(2)
    show["FinalNAV"] = show["FinalNAV"].round(2)
    print(show[["CAGR", "Vol", "Sharpe", "MaxDD", "Pct_1y_ge_12pct", "FinalNAV", "Periods"]].to_string())
    print("\nCore30 diversification: avg book size %.1f, avg max tag bucket %.2f (cap=%d)" %
          (bl["n_book"].mean(), bl["max_tag_bucket"].mean(), CORE_SLOT_CAP))

    statdf.to_csv(os.path.join(out, f"bt_stats_{cadence}.csv"))
    return retdf, statdf, eqdf, bl


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--cadence", default="M", choices=["M", "Q", "A"])
    ap.add_argument("--out", default=SCRATCH)
    a = ap.parse_args()
    run(a.cadence, a.out)
