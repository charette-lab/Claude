#!/usr/bin/env python3
"""deal_and_complement.py — characterize the actual deal engine, then find the
complement that actually works on the REAL monthly book.

Part 1: deal-level stats on the two provided tables (2006-2014, 2015-2025) — is the
return broad or one-deal? P&L concentration, hit rate, ex-best robustness, alpha breadth.

Part 2: the book's bad months are SYSTEMATIC, so a correlated long-equity sleeve can't
cut the drawdown — only a tail hedge can. Tests a put overlay on the ACTUAL monthly book
returns (MSCI World IMI, 2015+), net of the real fee structure.

  python3 pricing/deal_and_complement.py
"""
import numpy as np, pandas as pd
from scipy.stats import norm
import fee_analysis as F   # reuse monthly_series, fee_run, stats
SCR = F.SCR

# ---------- deal tables: (name, y0, y1, salesCAGR, ebitCAGR, IRR, indexCAGR, invested, returned, MOIC) ----------
D_recent = [
 ("Bufab",2015,2025,13.3,18.4,34.9,7.7,54,107,2.0),("DistIt",2015,2025,-1.8,0.8,6.6,7.8,77,94,1.2),
 ("Ngex",2016,2025,None,None,68.9,9.6,8,51,6.2),("Alcadon",2016,2025,22.3,13.2,73.8,9.2,38,64,1.7),
 ("Tempest",2018,2025,13.4,4.1,12.5,9.9,20,23,1.2),("Zalaris",2018,2025,11.5,33.3,25.5,9.9,50,73,1.5),
 ("Zutec",2018,2025,24.7,31.9,37.6,8.5,37,103,2.8),("Roko",2019,2025,69.6,65.8,28.4,10.7,26,134,5.1),
 ("Catella",2020,2025,8.5,86.5,12.0,9.9,71,87,1.2),("AI",2020,2025,None,None,3002.7,8.3,113,612,5.4),
 ("Renold",2024,2025,1.9,13.4,219.8,9.0,35,56,1.6),("Actic",2017,2023,-4.0,7.3,-26.9,6.5,81,34,0.4),
 ("Haldex",2019,2022,-0.9,13.2,35.9,2.5,22,36,1.6),("Robit",2015,2020,15.4,-13.7,13.7,6.2,54,73,1.4),
 ("Kitron",2016,2018,16.7,25.5,39.7,23.9,43,63,1.5),("Sandvik",2015,2017,-1.1,19.7,48.7,5.2,17,27,1.6),
 ("Transcom",2015,2016,-2.8,6.3,80.2,-0.2,21,36,1.7)]
D_hist = [
 ("Bilia",2007,2014,4.2,43.6,11.2,6.6,21,62,2.9),("Concentric",2011,2014,2.1,24.1,81.0,13.3,93,364,3.9),
 ("eWork",2007,2014,24.2,13.2,29.7,8.2,30,133,4.4),("Haldex",2009,2014,6.9,41.6,93.9,21.6,85,467,5.5),
 ("Lindab",2012,2014,1.1,13.8,32.3,17.1,448,702,1.6),("Note",2010,2014,-5.3,12.9,31.6,12.7,37,186,5.1),
 ("Transcom",2011,2014,4.5,32.6,12.8,12.4,128,170,1.3),("Global",2010,2014,-15.9,-7.6,16.1,13.4,142,214,1.5),
 ("Ferronordic",2010,2014,42.0,125.2,42.8,13.9,50,102,2.0),("Klarna",2007,2014,105.9,238.3,107.9,4.2,30,836,28.2),
 ("Acne",2007,2014,18.8,40.3,29.4,7.7,70,386,5.5),("Usports",2011,2013,0.0,0.0,-51.4,7.9,96,30,0.3),
 ("Catena",2006,2007,2.8,8.0,722.4,26.7,88,270,3.1),("Tigerholm",2006,2006,0.0,0.0,111.1,16.5,115,157,1.4),
 ("Avanza",2006,2014,12.4,11.0,33.7,12.3,313,2614,8.4),("Fabege",2006,2011,-3.3,-5.4,25.3,13.0,849,1942,2.3),
 ("JohnsonPump",2006,2006,7.2,145.0,51.9,28.9,71,213,3.0),("Klovern",2007,2011,19.8,16.3,17.7,9.7,110,227,2.1),
 ("Nobia",2007,2011,2.1,-8.2,8.4,12.5,490,668,1.4),("SkiStar",2007,2011,6.6,11.0,20.3,10.9,99,534,5.4),
 ("Brokk",2003,2006,24.0,158.0,73.8,27.8,120,630,5.3),("HQ",2007,2010,-11.5,-33.6,-5.4,13.5,684,393,0.6)]

def deal_stats(deals, tag):
    inv = np.array([d[7] for d in deals], float); ret = np.array([d[8] for d in deals], float)
    moic = np.array([d[9] for d in deals], float); irr = np.array([d[5] for d in deals], float)
    outp = np.array([d[5]-d[6] for d in deals], float)   # IRR - index
    pnl = ret - inv
    order = np.argsort(-pnl)
    tot_inv, tot_ret, tot_pnl = inv.sum(), ret.sum(), pnl.sum()
    top1 = pnl[order[0]]/tot_pnl; top3 = pnl[order[:3]].sum()/tot_pnl; top5 = pnl[order[:5]].sum()/tot_pnl
    # ex the single best P&L deal
    keep = order[1:]
    exbest_moic = ret[keep].sum()/inv[keep].sum()
    print(f"\n=== {tag}  ({len(deals)} deals) ===")
    print(f"  total invested {tot_inv:.0f}  returned {tot_ret:.0f}  -> weighted MOIC {tot_ret/tot_inv:.2f}x")
    print(f"  hit rate  MOIC>1: {(moic>1).mean()*100:.0f}%   |   beat index (IRR>idx): {(outp>0).mean()*100:.0f}%")
    print(f"  MOIC   median {np.median(moic):.1f}x   mean {moic.mean():.1f}x   min {moic.min():.1f}x   max {moic.max():.1f}x")
    print(f"  IRR    median {np.median(irr):.1f}%   outperformance median {np.median(outp):.1f}%")
    print(f"  P&L concentration: best deal {top1*100:.0f}%  top-3 {top3*100:.0f}%  top-5 {top5*100:.0f}%  of total profit")
    print(f"  robustness: drop the single best P&L deal -> weighted MOIC still {exbest_moic:.2f}x on {inv[keep].sum():.0f} invested")
    return dict(inv=tot_inv, ret=tot_ret)

print("#"*80); print("PART 1 — is the return engine broad, or one deal?"); print("#"*80)
deal_stats(D_recent, "2015-2025 book")
deal_stats(D_hist,   "2006-2014 book")
alld = D_recent + D_hist
inv = np.array([d[7] for d in alld], float); ret = np.array([d[8] for d in alld], float)
print(f"\n  BOTH books combined: {len(alld)} deals, {inv.sum():.0f} invested -> {ret.sum():.0f} returned "
      f"({ret.sum()/inv.sum():.2f}x); {( (ret-inv)>0).mean()*100:.0f}% of deals profitable.")

# ---------- Part 2: the complement, tested on the ACTUAL monthly book ----------
print("\n"+"#"*80); print("PART 2 — the book's drawdowns are systematic; test the hedge on the real book"); print("#"*80)
m = F.monthly_series()
msci = pd.read_parquet(SCR+"/msci_imi_monthly.parquet").set_index("date")["ret"]
msci.index = pd.to_datetime(msci.index)
df = pd.concat([m.rename("b"), msci.rename("u")], axis=1).dropna()
b, u = df["b"].values, df["u"].values
beta = np.cov(b, u)[0,1]/np.var(u); corr = np.corrcoef(b, u)[0,1]
# downside correlation: only months the market fell
dn = u < 0
beta_dn = np.cov(b[dn], u[dn])[0,1]/np.var(u[dn]); corr_dn = np.corrcoef(b[dn], u[dn])[0,1]
sig = np.std(u, ddof=1)*np.sqrt(12)
print(f"\n  book vs MSCI World IMI, monthly {df.index.min():%Y-%m}..{df.index.max():%Y-%m} ({len(df)} months)")
print(f"  full-sample beta {beta:.2f}  corr {corr:.2f}   |   DOWN-market beta {beta_dn:.2f}  corr {corr_dn:.2f}")
print("  -> a long-equity 'diversifier' shares these down months; only a hedge is negatively convex to them.")

RF, VRP, SKEW = 0.02, 0.03, 0.05
def bsput(S,K,s,r,t):
    if t<=1e-9: return max(K-S,0.0)
    d1=(np.log(S/K)+(r+0.5*s*s)*t)/(s*np.sqrt(t)); return K*np.exp(-r*t)*norm.cdf(-(d1-s*np.sqrt(t)))-S*norm.cdf(-d1)
def overlay_m(b,u,beta,sig,money,tn):
    si=sig+VRP+SKEW*(1-money)/0.10; S=np.cumprod(np.r_[1.0,1+u]); E=1.0; pv=0.0; con=None; nav=[1.0]
    for q in range(len(b)):
        if q%tn==0:
            if con: E+=pv; pv=0.0; con=None
            N=beta*E; K=money*S[q]; cost=bsput(S[q],K,si,RF,tn/12.)/S[q]*N; E-=cost; pv=cost; con=(K,q,N,S[q])
        E*=(1+b[q])
        if con: K,q0,N,S0=con; tau=max((q0+tn-(q+1))/12.,0.0); pv=bsput(S[q+1],K,si,RF,tau)/S0*N
        nav.append(E+pv)
    return np.array(nav)

def mstats(nav):
    r=nav[1:]/nav[:-1]-1; y=len(r)/12; eq=nav/nav[0]; dd=eq/np.maximum.accumulate(eq)-1; dnr=r[r<0]
    return dict(CAGR=eq[-1]**(1/y)-1, MaxDD=dd.min(), Vol=np.std(r,ddof=1)*np.sqrt(12),
                Sortino=np.mean(r)*12/(np.std(dnr,ddof=1)*np.sqrt(12)), Calmar=(eq[-1]**(1/y)-1)/abs(dd.min()))

rows = {"Book unhedged (actual monthly)": np.cumprod(np.r_[1.0,1+b])}
for money,l in ((0.95,"5% OTM"),(0.90,"10% OTM")):
    rows[f"Book + put {l} (semi-annual)"] = overlay_m(b,u,beta,sig,money,6)
print(f"\n  {'strategy':34s} {'CAGR':>6} {'Vol':>6} {'MaxDD':>7} {'Sortino':>7} {'Calmar':>6}")
for k,nav in rows.items():
    s=mstats(nav); print(f"  {k:34s} {s['CAGR']*100:5.1f}% {s['Vol']*100:5.1f}% {s['MaxDD']*100:6.1f}% {s['Sortino']:7.2f} {s['Calmar']:6.2f}")

# net-of-fee HWM years: unhedged vs 10% OTM hedged, monthly annual crystallization
def hwm_years(nav):
    r = pd.Series(nav[1:]/nav[:-1]-1, index=df.index)
    net, led = F.fee_run(((k, r[k]) for k in r.index), 12, lambda k: k.year)
    return led["yrs_under"], len(set(df.index.year))
u0,ny = hwm_years(rows["Book unhedged (actual monthly)"]); u1,_ = hwm_years(rows["Book + put 10% OTM (semi-annual)"])
print(f"\n  net-of-fee years below high-water mark (2015-2026): unhedged {u0}/{ny}  vs  +put {u1}/{ny}")

import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
idx = pd.DatetimeIndex([df.index[0]-pd.offsets.MonthEnd()]).append(df.index)
col={"Book unhedged (actual monthly)":"#d62728","Book + put 5% OTM (semi-annual)":"#1f77b4","Book + put 10% OTM (semi-annual)":"#2ca02c"}
fig,(a1,a2)=plt.subplots(2,1,figsize=(11,8.5),height_ratios=[3,1.5],sharex=True)
for k,nav in rows.items():
    a1.plot(idx[:len(nav)],nav,label=k,color=col[k],lw=2.2 if "unhedged" in k else 1.8)
    eq=nav/nav[0]; dd=eq/np.maximum.accumulate(eq)-1; a2.plot(idx[:len(nav)],dd*100,color=col[k],lw=1.6)
a1.set_yscale("log"); a1.legend(loc="upper left"); a1.grid(True,which="both",alpha=0.25); a1.set_ylabel("Growth of 1.0 (log)")
a1.set_title("The complement that works: MSCI World IMI put overlay on the ACTUAL monthly book\n"
             f"book down-market beta {beta_dn:.2f} to the index — a long sleeve shares the crash; the put offsets it", fontsize=10.5)
a2.grid(True,alpha=0.25); a2.set_ylabel("drawdown (%)")
plt.tight_layout(); fig.savefig(SCR+"/deal_complement.png",dpi=130); print("\nwrote deal_complement.png")
