#!/usr/bin/env python3
"""attribution.py — yearly return attribution for the books: internal growth vs rerating.

Identity (Bogle): a stock's price return over a year = earnings growth x multiple
change.  Using NOPAT ("New Operating Income") as the earnings base and MC/NOPAT as
the multiple:
    1 + r   =  (NOPAT1/NOPAT0)  x  (m1/m0)
    r        =  internal growth  +  rerating   (+ small cross term; exact in logs)
INTERNAL GROWTH (carry) = the business earning more; RERATING = the market paying a
different multiple for the same earnings.

For each calendar year we reconstitute the book at the prior year-end (the same
selection the backtest uses), hold it through the year, and split the realized
return.  Growth/rerating are aggregated in log space over names with positive NOPAT
at both ends (so they compose exactly); coverage is reported.

  python3 pricing/attribution.py
"""
import json, os
import numpy as np, pandas as pd

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"

mc = pd.read_parquet(f"{SCR}/daily_marketcap_ev.parquet", columns=["Instrument", "Date", "market_cap"])
mc["Instrument"] = mc["Instrument"].astype(str); mc["Date"] = pd.to_datetime(mc["Date"]).astype("datetime64[ns]")
nop = pd.read_parquet(f"{SCR}/nopat_yearly.parquet"); nop["Instrument"] = nop["Instrument"].astype(str)
raw = pd.read_parquet(f"{SCR}/daily_expected_return.parquet"); raw["Instrument"] = raw["Instrument"].astype(str)
raw["Date"] = pd.to_datetime(raw["Date"]).astype("datetime64[ns]")
full = pd.read_parquet(f"{SCR}/daily_expected_return_full.parquet"); full["Instrument"] = full["Instrument"].astype(str)
full["Date"] = pd.to_datetime(full["Date"]).astype("datetime64[ns]")
import openpyxl
ws = openpyxl.load_workbook(f"{SCR}/Universe_final.xlsx", data_only=True, read_only=True)["Scored"]
rws = list(ws.iter_rows(values_only=True)); ci = {x: i for i, x in enumerate(rws[0])}
moat = {str(r[ci["Ticker"]]): r[ci["CompanyMoat(v3.2)"]] for r in rws[1:]
        if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]], (int, float))}


def asof_wide(df, val, dates, names=None):
    sub = df if names is None else df[df["Instrument"].isin(names)]
    w = sub.pivot_table(index="Date", columns="Instrument", values=val, aggfunc="last").sort_index().ffill()
    return w.reindex(pd.DatetimeIndex(sorted(dates)), method="ffill")


def yearends(dates):
    s = pd.Series(sorted(set(dates)))
    return list(s.groupby(s.dt.to_period("Y")).max().values)


GRID = [d for d in yearends(mc["Date"].unique()) if pd.Timestamp(d).year >= 2015]
mc_w = asof_wide(mc, "market_cap", GRID)
nop_w = asof_wide(nop.rename(columns={"ped": "Date"}), "nopat", GRID)


def book_at(t, er_df, er_col, min_moat=None, drop_artifact=False, n=30, gate=0.12):
    row = er_df[er_df["Date"] == t]
    if row.empty:
        return []
    s = row.set_index("Instrument")
    er = s[er_col].copy()
    if drop_artifact and "artifact" in s:
        er[s["artifact"].astype(bool)] = np.nan
    er = er.dropna(); er = er[er >= gate]
    if min_moat is not None:
        er = er[[moat.get(t_, 0) > min_moat for t_ in er.index]]
    return list(er.sort_values(ascending=False).index)[:n]


def decompose(label, er_df, er_col, **kw):
    out = []
    er_df = er_df.copy()
    er_df["Date"] = pd.to_datetime(er_df["Date"]).astype("datetime64[ns]")
    # snap ER selection dates to the grid (asof)
    er_w = asof_wide(er_df, er_col, GRID)
    art_w = asof_wide(er_df.assign(a=er_df.get("artifact", False).astype(float)), "a", GRID) if "artifact" in er_df else None
    for i in range(len(GRID) - 1):
        t0, t1 = GRID[i], GRID[i + 1]
        yr = pd.Timestamp(t1).year
        er0 = er_w.loc[t0].dropna()
        er0 = er0[er0 >= kw.get("gate", 0.12)]
        if kw.get("drop_artifact") and art_w is not None:
            er0 = er0[art_w.loc[t0].reindex(er0.index).fillna(0) < 0.5]
        if kw.get("min_moat") is not None:
            er0 = er0[[moat.get(t_, 0) > kw["min_moat"] for t_ in er0.index]]
        book = list(er0.sort_values(ascending=False).index)[:30]
        if not book:
            continue
        m0, m1 = mc_w.loc[t0], mc_w.loc[t1]
        n0, n1 = nop_w.loc[t0], nop_w.loc[t1]
        # decompose on names with positive NOPAT at both ends (so growth+rerating compose
        # exactly in logs); winsorize gross return to [-95%, +500%] to kill microcap/data
        # artifacts that a real equal-weight, frequently-rebalanced book never realises.
        LO, HI = 0.05, 6.0
        lr, lg, lu = [], [], []
        for t_ in book:
            a, b = m0.get(t_), m1.get(t_)
            e0, e1 = n0.get(t_), n1.get(t_)
            if not (a and b and a > 0 and b > 0 and e0 and e1 and e0 > 0 and e1 > 0):
                continue
            gr = min(max(b / a, LO), HI)
            lr.append(np.log(gr)); lg.append(np.log(e1 / e0)); lu.append(np.log(gr) - np.log(e1 / e0))
        if not lr:
            continue
        R = float(np.exp(np.mean(lr)) - 1)
        G = float(np.exp(np.mean(lg)) - 1)
        Uu = float(np.exp(np.mean(lu)) - 1)
        share = round(np.mean(lu) / np.mean(lr) * 100) if abs(R) >= 0.05 else None   # noisy when net~0
        out.append({"Year": yr, "Return_%": round(R*100, 1),
                    "Growth_%": round(G*100, 1), "Rerating_%": round(Uu*100, 1),
                    "Rerate_share_%": share, "n": len(lr), "clean_%": round(len(lr)/len(book)*100)})
    df = pd.DataFrame(out)
    # full-period compounded split (geometric, clean subset)
    g = df.dropna(subset=["Growth_%"])
    comp_g = np.prod([1+x/100 for x in g["Growth_%"]]) - 1
    comp_u = np.prod([1+x/100 for x in g["Rerating_%"]]) - 1
    comp_r = np.prod([1+x/100 for x in df["Return_%"]]) - 1
    df.attrs["totals"] = {"label": label, "compound_return_%": round(comp_r*100),
                          "from_growth_%": round(comp_g*100), "from_rerating_%": round(comp_u*100)}
    return df


carry = pd.read_parquet(f"{SCR}/carry_grid_norm.parquet"); carry["Instrument"] = carry["Instrument"].astype(str)
carry["Date"] = pd.to_datetime(carry["Date"]).astype("datetime64[ns]")

BOOKS = {
    "Full engine, full universe": dict(er_df=full, er_col="expected_return", drop_artifact=True),
    "Full engine, moat>7.8": dict(er_df=full, er_col="expected_return", drop_artifact=True, min_moat=7.8),
    "Carry-selected, full universe": dict(er_df=carry, er_col="expected_return", drop_artifact=True),
    "Carry + moat>7.8": dict(er_df=carry, er_col="expected_return", drop_artifact=True, min_moat=7.8),
}

tables = {}
for label, kw in BOOKS.items():
    df = decompose(label, **kw)
    tables[label] = df
    print(f"\n==== {label} ====")
    print(df.to_string(index=False))
    t = df.attrs["totals"]
    print(f"  FULL PERIOD: total {t['compound_return_%']}%  =  growth {t['from_growth_%']}%  x  rerating {t['from_rerating_%']}%")

# ---- chart: yearly growth vs rerating, one panel per book ----
import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
fig, axes = plt.subplots(3, 1, figsize=(11, 10), sharex=True)
for ax, (label, df) in zip(axes, tables.items()):
    x = df["Year"].values
    ax.bar(x - 0.2, df["Growth_%"], 0.4, label="Internal growth (earnings)", color="#2ca02c")
    ax.bar(x + 0.2, df["Rerating_%"], 0.4, label="Rerating (multiple)", color="#d62728")
    ax.plot(x, df["Return_%"], "ko-", ms=4, lw=1, label="Total return")
    ax.axhline(0, color="k", lw=0.6); ax.grid(True, axis="y", alpha=0.25)
    t = df.attrs["totals"]
    ax.set_title(f"{label}  —  full period: growth {t['from_growth_%']}%  ×  rerating {t['from_rerating_%']}%",
                 fontsize=10)
    ax.set_ylabel("contribution %")
axes[0].legend(loc="upper left", fontsize=8, ncol=3)
axes[-1].set_xlabel("Year")
fig.suptitle("Where the return comes from: internal growth vs rerating (yearly)", fontsize=12)
plt.tight_layout(); fig.savefig(f"{SCR}/attribution.png", dpi=130)
print("wrote attribution.png")

with pd.ExcelWriter(f"{SCR}/Attribution_yearly.xlsx", engine="openpyxl") as xl:
    summ = pd.DataFrame([t.attrs["totals"] for t in tables.values()])
    summ.to_excel(xl, sheet_name="Summary", index=False)
    for label, df in tables.items():
        df.to_excel(xl, sheet_name=label[:31], index=False)
print("\nwrote Attribution_yearly.xlsx")
