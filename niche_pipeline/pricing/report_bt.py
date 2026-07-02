#!/usr/bin/env python3
"""Build the equity-curve chart + summary workbook from backtest outputs."""
import json
import pandas as pd, numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

S = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"

LABELS = {
    "Core30_framework": "Core-30 (full framework: Gate-1 + slot cap)",
    "Top30_no_slotcap": "Top-30 by ER (no slot cap)",
    "Gate_EW": "All ER≥12% equal-weight (gate only)",
    "Universe_EW": "Tagged universe equal-weight (benchmark)",
}
COL = {"Core30_framework": "#1f77b4", "Top30_no_slotcap": "#ff7f0e",
       "Gate_EW": "#2ca02c", "Universe_EW": "#7f7f7f"}

eqM = pd.read_parquet(f"{S}/bt_equity_M.parquet")
retM = pd.read_parquet(f"{S}/bt_returns_M.parquet")
statM = pd.read_csv(f"{S}/bt_stats_M.csv", index_col=0)
statQ = pd.read_csv(f"{S}/bt_stats_Q.csv", index_col=0)

# ---- chart: equity curves (log) + drawdown ----
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8.5), height_ratios=[3, 1.4], sharex=True)
for c in ["Core30_framework", "Top30_no_slotcap", "Gate_EW", "Universe_EW"]:
    ax1.plot(eqM.index, eqM[c], label=LABELS[c], color=COL[c],
             lw=2.2 if c == "Core30_framework" else 1.4)
ax1.set_yscale("log")
ax1.set_ylabel("Growth of 1.0 (log scale)")
ax1.set_title("Constrained Quality Compounder Index — realized performance, 2015–2026\n"
              "(monthly rebalance; Gate-2 downside screen disabled; local-ccy price return, no costs)",
              fontsize=11)
ax1.legend(loc="upper left", fontsize=9)
ax1.grid(True, which="both", alpha=0.25)
ax1.axhline(1.0, color="k", lw=0.6)
# 12% target reference line off the framework start
yrs = np.arange(len(eqM))
ax1.plot(eqM.index, 1.12 ** (yrs / 12.0), ls="--", color="red", lw=1.0,
         label="12% target")
ax1.legend(loc="upper left", fontsize=8.5)

eq = eqM["Core30_framework"]
dd = eq / eq.cummax() - 1
ax2.fill_between(dd.index, dd * 100, 0, color="#1f77b4", alpha=0.5)
ax2.set_ylabel("Core-30\ndrawdown (%)")
ax2.grid(True, alpha=0.25)
plt.tight_layout()
fig.savefig(f"{S}/backtest_equity.png", dpi=130)
print("wrote backtest_equity.png")

# ---- workbook ----
def fmt(df):
    d = df.copy()
    for c in ["CAGR", "Vol", "MaxDD", "TotalReturn", "MeanPeriodRet", "Pct_1y_ge_12pct"]:
        if c in d: d[c] = (d[c] * 100).round(2)
    for c in ["Sharpe", "FinalNAV"]:
        if c in d: d[c] = d[c].round(2)
    return d

order = ["Core30_framework", "Top30_no_slotcap", "Gate_EW", "Universe_EW"]
sM = fmt(statM).reindex(order); sM.index = [LABELS[i] for i in sM.index]
sQ = fmt(statQ).reindex(order); sQ.index = [LABELS[i] for i in sQ.index]

hold = json.load(open(f"{S}/bt_holdings_M.json"))["Core30_framework"]
hrows = []
for d in sorted(hold)[::6]:   # every ~6 months
    hrows.append({"rebalance": d, "n": len(hold[d]), "names": ", ".join(hold[d])})
holddf = pd.DataFrame(hrows)

cols = ["CAGR", "Vol", "Sharpe", "MaxDD", "Pct_1y_ge_12pct", "FinalNAV", "TotalReturn", "Periods"]
notes = pd.DataFrame({"Methodology & caveats": [
    "SYSTEM: Constrained Quality Compounder Index (IPS). Gate-2 (downside hard-stop) DISABLED per instruction.",
    "Universe: 1,574 high-moat researched names carrying the 11 binary risk tags (a stock must be tagged before purchase).",
    "Binary tag = severity >= 3 (High/Extreme) from severity_master, else severity_from_notes >= 3, else research-cache native binary.",
    "Gate 1: forward Expected IRR >= 12%, from the daily ROIIC-engine ER panel (no look-ahead: ER@t uses price@t + prior FY fundamentals).",
    "Build: rank eligible by ER desc; fill 30 equal-weighted names under the 6-of-30 slot cap (no risk tag in >6 of 30 names).",
    "Realized return: close-to-close, LOCAL currency, price only (no dividends, no FX translation, no transaction costs/slippage).",
    "Delisting: a name with no print at next rebalance carries its last close (0% that period) — mildly optimistic for failures.",
    "Benchmarks: Top-30 by ER without the slot cap; all ER>=12% equal-weight (gate only); tagged-universe equal-weight.",
    "Rebalance grids: monthly (126 periods) and quarterly (42 periods, IPS-native). Both shown.",
    "READ: absolute CAGR carries FX/cost noise; the ROBUST signal is the RELATIVE ranking (framework vs benchmarks), invariant to FX.",
]})

with pd.ExcelWriter(f"{S}/Backtest_CQCI.xlsx", engine="openpyxl") as xl:
    sM[cols].to_excel(xl, sheet_name="Stats_Monthly")
    sQ[cols].to_excel(xl, sheet_name="Stats_Quarterly")
    (retM.rename(columns=LABELS)).to_excel(xl, sheet_name="Monthly_Returns")
    (eqM.rename(columns=LABELS)).to_excel(xl, sheet_name="Equity_Curves")
    holddf.to_excel(xl, sheet_name="Core30_Holdings_semiannual", index=False)
    notes.to_excel(xl, sheet_name="Methodology", index=False)
print("wrote Backtest_CQCI.xlsx")
print("\nMONTHLY:\n", sM[cols].to_string())
print("\nQUARTERLY:\n", sQ[cols].to_string())
