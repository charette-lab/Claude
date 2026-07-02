#!/usr/bin/env python3
"""put_overlay.py — protective-put overlay as the risk adjuster vs Gate 2.

Instead of screening NAMES (Gate 2, which drags return), overlay a rolling
protective put on the book's systematic factor (the equal-weight market proxy,
to which the Core-30 has beta ~0.84, corr ~0.93). The put offsets systematic
crashes (2020, 2022) while the book keeps every return-generating name.

Modelled honestly: Black-Scholes puts on the market proxy, implied vol = realized
+ a variance-risk-premium markup (so the insurance is appropriately expensive),
notional = book-beta x NAV, rolled, quarterly mark-to-market so it cuts intra-year
drawdown. Compares moneyness/tenor against unhedged and Gate 2.

  python3 pricing/put_overlay.py
"""
import os
import numpy as np, pandas as pd
from scipy.stats import norm

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U = "/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
PRICE = ["972f0581-daily_volume_price_0", "257124b3-daily_volume_price_1",
         "86e54ec3-daily_volume_price_2", "13f82c18-daily_volume_price_0630"]
VRP = 0.03            # implied = realized vol + 3 vol points (variance risk premium)
SKEW = 0.05           # extra implied vol per 10% OTM (real puts trade rich on the skew)
RF = 0.02


def stats(nav):
    nav = np.asarray(nav, float); r = nav[1:]/nav[:-1]-1; y = len(r)/4
    eq = nav/nav[0]; dd = eq/np.maximum.accumulate(eq)-1
    return dict(CAGR=eq[-1]**(1/y)-1, MaxDD=dd.min(), Sharpe=np.mean(r)*4/(np.std(r, ddof=1)*np.sqrt(4)),
                Calmar=(eq[-1]**(1/y)-1)/abs(dd.min()))


def bsput(S, K, s, r, t):
    if t <= 1e-9:
        return max(K-S, 0.0)
    d1 = (np.log(S/K)+(r+0.5*s*s)*t)/(s*np.sqrt(t))
    return K*np.exp(-r*t)*norm.cdf(-(d1-s*np.sqrt(t))) - S*norm.cdf(-d1)


def market_proxy(book_index):
    px = pd.concat([pd.read_parquet(U+f+".parquet", columns=["Instrument", "Date", "Close Price"]) for f in PRICE])
    px["Instrument"] = px["Instrument"].astype(str); px["Date"] = pd.to_datetime(px["Date"])
    px["c"] = pd.to_numeric(px["Close Price"], errors="coerce")
    W = px[px.c > 0].pivot_table(index="Date", columns="Instrument", values="c", aggfunc="last").sort_index()
    qe = pd.Series(W.index).groupby(pd.Series(W.index).dt.to_period("Q")).max().values
    Wg = W.reindex(pd.DatetimeIndex(sorted(qe)), method="ffill")
    return (Wg/Wg.shift(1)-1).mean(axis=1).reindex(book_index)


def overlay(b, u, beta, sig, money, tn):
    si = sig+VRP+SKEW*(1-money)/0.10          # implied vol: realized + VRP + OTM skew
    S = np.cumprod(np.r_[1.0, 1+u]); E = 1.0; pv = 0.0; con = None; nav = [1.0]
    for q in range(len(b)):
        if q % tn == 0:
            if con:
                E += pv; pv = 0.0; con = None
            N = beta*E; K = money*S[q]
            cost = bsput(S[q], K, si, RF, tn/4.0)/S[q]*N
            E -= cost; pv = cost; con = (K, q, N, S[q])
        E *= (1+b[q])
        if con:
            K, q0, N, S0 = con; tau = max((q0+tn-(q+1))/4.0, 0.0)
            pv = bsput(S[q+1], K, si, RF, tau)/S0*N
        nav.append(E+pv)
    return np.array(nav)


def msci_quarterly(book_index):
    """REAL MSCI World IMI: monthly returns (uploaded) compounded into the book's quarters."""
    m = pd.read_parquet(SCR+"/msci_imi_monthly.parquet").set_index("date")["ret"]
    lvl = (1+m).cumprod().reindex(book_index, method="ffill")
    return lvl/lvl.shift(1)-1


def main():
    book = pd.read_parquet(SCR+"/bt_core30_returns_Q.parquet").iloc[:, 0]; book.index = pd.to_datetime(book.index)
    mret = msci_quarterly(book.index) if os.path.exists(SCR+"/msci_imi_monthly.parquet") else market_proxy(book.index)
    df = pd.concat([book.rename("b"), mret.rename("u")], axis=1).dropna()
    b, u = df["b"].values, df["u"].values
    beta = np.cov(b, u)[0, 1]/np.var(u); sig = np.std(u, ddof=1)*np.sqrt(4)
    print(f"book beta to market {beta:.2f} | market vol {sig*100:.1f}% | corr {np.corrcoef(b,u)[0,1]:.2f}\n")
    rows = {"Unhedged Core-30": np.cumprod(np.r_[1.0, 1+b])}
    g2 = pd.read_parquet(SCR+"/bt_core30_returns_Q_g2.parquet").iloc[:, 0].reindex(df.index).dropna().values
    rows["+ Gate 2 (name screen)"] = np.cumprod(np.r_[1.0, 1+g2])
    for money in (0.95, 0.90, 0.85):
        for tn, l in ((4, "annual"), (2, "semi")):
            rows[f"+ put {int((1-money)*100)}% OTM {l}"] = overlay(b, u, beta, sig, money, tn)
    print(f"{'strategy':32s} {'CAGR':>6} {'MaxDD':>7} {'Sharpe':>6} {'Calmar':>6}")
    out = {}
    for k, nav in rows.items():
        s = stats(nav); out[k] = s
        print(f"{k:32s} {s['CAGR']*100:5.1f}% {s['MaxDD']*100:6.1f}% {s['Sharpe']:6.2f} {s['Calmar']:6.2f}")
    # chart: unhedged vs Gate2 vs the recommended 9% OTM semi
    import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
    pick = {"Unhedged Core-30": "#7f7f7f", "+ Gate 2 (name screen)": "#d62728", "+ put 9% OTM semi": "#1f77b4"}
    idx = pd.DatetimeIndex([book.index[0]-pd.offsets.QuarterEnd()]).append(df.index)
    fig, (a1, a2) = plt.subplots(2, 1, figsize=(11, 8), height_ratios=[3, 1.5], sharex=True)
    for k, c in pick.items():
        nav = rows[k]; a1.plot(idx[:len(nav)], nav, label=k, color=c, lw=2.2 if "put" in k else 1.6)
        eq = nav/nav[0]; dd = eq/np.maximum.accumulate(eq)-1
        a2.plot(idx[:len(nav)], dd*100, color=c, lw=1.6)
    a1.set_yscale("log"); a1.legend(loc="upper left"); a1.grid(True, which="both", alpha=0.25); a1.set_ylabel("Growth of 1.0 (log)")
    a1.set_title("Risk adjuster: protective-put overlay vs Gate 2 (name screen) vs unhedged\n"
                 "put on REAL MSCI World IMI (book corr ~0.76), implied vol = realized + VRP + OTM skew", fontsize=10.5)
    a2.grid(True, alpha=0.25); a2.set_ylabel("drawdown (%)")
    plt.tight_layout(); fig.savefig(SCR+"/put_overlay.png", dpi=130); print("\nwrote put_overlay.png")


if __name__ == "__main__":
    main()
