#!/usr/bin/env python3
import json
import pandas as pd, numpy as np
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt

S = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"

# new pure-ER bands (monthly) + prior framework strategies (monthly)
erb = pd.read_parquet(f"{S}/bt_erbands_returns_M.parquet")["ER_bands"]
prior = pd.read_parquet(f"{S}/bt_returns_M.parquet")
ret = pd.DataFrame({
    "ER-only + IPS bands (full universe)": erb,
    "Core-30 framework (slot cap, tagged)": prior["Core30_framework"],
    "Tagged universe EW (benchmark)": prior["Universe_EW"],
}).dropna()
eq = (1 + ret).cumprod()
COL = ["#d62728", "#1f77b4", "#7f7f7f"]

fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8.5), height_ratios=[3, 1.4], sharex=True)
for c, col in zip(ret.columns, COL):
    ax1.plot(eq.index, eq[c], label=c, color=col, lw=2.3 if "ER-only" in c else 1.6)
ax1.set_yscale("log"); ax1.set_ylabel("Growth of 1.0 (log)")
ax1.set_title("Pure-ER Core index on IPS tolerance bands vs the risk-tagged framework\n"
              "(monthly checkup; ER>=12% selector; trim@5% / liquidate@2%; price-only, no costs)",
              fontsize=11)
yrs = np.arange(len(eq))
ax1.plot(eq.index, 1.12 ** (yrs / 12.0), ls="--", color="green", lw=1.0, label="12% target")
ax1.legend(loc="upper left", fontsize=9); ax1.grid(True, which="both", alpha=0.25)
e = eq.iloc[:, 0]; dd = e / e.cummax() - 1
ax2.fill_between(dd.index, dd * 100, 0, color="#d62728", alpha=0.5)
ax2.set_ylabel("ER-bands\ndrawdown (%)"); ax2.grid(True, alpha=0.25)
plt.tight_layout(); fig.savefig(f"{S}/backtest_erbands.png", dpi=130)
print("wrote backtest_erbands.png")

# stats workbook
def stats(r, ppy):
    eq = (1 + r).cumprod(); yrs = len(r) / ppy
    dd = eq / eq.cummax() - 1
    roll = eq.pct_change(ppy).dropna()
    return dict(CAGR=eq.iloc[-1]**(1/yrs)-1, Vol=r.std(ddof=1)*np.sqrt(ppy),
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)), MaxDD=dd.min(),
                Pct_1y_ge_12pct=(roll >= 0.12).mean(), FinalNAV=eq.iloc[-1])

rows = {c: stats(ret[c], 12) for c in ret.columns}
st = pd.DataFrame(rows).T
for c in ["CAGR", "Vol", "MaxDD", "Pct_1y_ge_12pct"]:
    st[c] = (st[c] * 100).round(2)
st["Sharpe"] = st["Sharpe"].round(2); st["FinalNAV"] = st["FinalNAV"].round(2)

# turnover summary (from holdings)
def turnover(path, ppy):
    h = json.load(open(path)); ds = sorted(h); bks = [set(h[d]) for d in ds]
    ch = [(len(a-b)+len(b-a))/2 for a, b in zip(bks[:-1], bks[1:])]
    ch = ch[1:]
    ab = np.mean([len(b) for b in bks])
    return dict(avg_changed=np.mean(ch), pct_book=np.mean(ch)/ab*100,
                annual_oneway_pct=np.mean(ch)*ppy/ab*100, avg_book=ab,
                hold_months=ab/(np.mean(ch)*ppy)*12 if np.mean(ch) else np.nan)
def turnover_framework(path, key, ppy):
    h = json.load(open(path))[key]; ds = sorted(h); bks = [set(h[d]) for d in ds]
    ch = [(len(a-b)+len(b-a))/2 for a, b in zip(bks[:-1], bks[1:])][1:]
    ab = np.mean([len(b) for b in bks])
    return dict(avg_changed=np.mean(ch), pct_book=np.mean(ch)/ab*100,
                annual_oneway_pct=np.mean(ch)*ppy/ab*100, avg_book=ab,
                hold_months=ab/(np.mean(ch)*ppy)*12 if np.mean(ch) else np.nan)
tov = pd.DataFrame({
    "ER-only + IPS bands (Q)": turnover(f"{S}/bt_erbands_holdings_Q.json", 4),
    "ER-only + IPS bands (M)": turnover(f"{S}/bt_erbands_holdings_M.json", 12),
    "Core-30 framework (M, reconstitute)": turnover_framework(f"{S}/bt_holdings_M.json", "Core30_framework", 12),
}).T

notes = pd.DataFrame({"Methodology & caveats": [
    "SELECTOR = Expected IRR only. Risk tags IGNORED (tag coverage incomplete) -> universe is ALL 3,121 priced ER securities.",
    "Rest of the IPS retained: 30 names equal-weighted 3.33%; Gate-1 ER>=12%; quarterly (or monthly) checkup with tolerance bands.",
    "Trigger 2 (5% momentum cap): trim a name back to 3.33%; redeploy into highest-ER holdings below target.",
    "Trigger 3 (2% thesis violation): liquidate; replace with the highest-ER name not held. Trigger 1 (risk-tag cap) dropped.",
    "Bands (not the calendar) drive trades -> turnover ~15-17%/yr (~4-5 names/yr), avg holding period ~6-7 years.",
    "No look-ahead in selection; realized return close-to-close, local ccy, price-only, no dividends/FX/costs.",
    "Robust signal is the RELATIVE ranking. Absolute CAGR is an upper bound (no costs; turnover here is low so cost drag is small).",
]})

with pd.ExcelWriter(f"{S}/Backtest_ER_bands.xlsx", engine="openpyxl") as xl:
    st[["CAGR", "Vol", "Sharpe", "MaxDD", "Pct_1y_ge_12pct", "FinalNAV"]].to_excel(xl, sheet_name="Stats_Monthly")
    tov.round(1).to_excel(xl, sheet_name="Turnover")
    ret.to_excel(xl, sheet_name="Monthly_Returns")
    eq.to_excel(xl, sheet_name="Equity_Curves")
    h = json.load(open(f"{S}/bt_erbands_holdings_Q.json"))
    hr = [{"checkup": d, "n": len(h[d]), "names": ", ".join(h[d])} for d in sorted(h)[::4]]
    pd.DataFrame(hr).to_excel(xl, sheet_name="ERbands_Holdings_annual", index=False)
    notes.to_excel(xl, sheet_name="Methodology", index=False)
print("wrote Backtest_ER_bands.xlsx\n")
print(st[["CAGR", "Vol", "Sharpe", "MaxDD", "Pct_1y_ge_12pct", "FinalNAV"]].to_string())
