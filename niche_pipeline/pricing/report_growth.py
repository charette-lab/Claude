#!/usr/bin/env python3
import json, pandas as pd, numpy as np, openpyxl
import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
S="/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
def col(p): return pd.read_parquet(f"{S}/{p}")["ER_bands"]
ret=pd.DataFrame({
    "Internal-growth + moat>7.8":       col("bt_erbands_returns_M_moat7.8_growth.parquet"),
    "Internal-growth, full universe":   col("bt_erbands_returns_M_growth.parquet"),
    "Total-ER + moat>7.8 (rerating-incl)":col("bt_erbands_returns_M_moat7.8_full.parquet"),
    "Universe EW (benchmark)":          pd.read_parquet(f"{S}/bt_returns_M.parquet")["Universe_EW"],
}).dropna()
eq=(1+ret).cumprod(); COL=["#2ca02c","#98df8a","#1f77b4","#7f7f7f"]
fig,(ax1,ax2)=plt.subplots(2,1,figsize=(11,8.5),height_ratios=[3,1.4],sharex=True)
for c,cl in zip(ret.columns,COL): ax1.plot(eq.index,eq[c],label=c,color=cl,lw=2.4 if "growth + moat" in c else 1.6)
ax1.set_yscale("log"); ax1.set_ylabel("Growth of 1.0 (log)")
ax1.set_title("Select for INTERNAL GROWTH (gate er1>=12%, rank by engine carry) vs total-ER\n"
              "(full engine; monthly checkup; trim@5% / liquidate@2%; price-only, no costs)",fontsize=11)
yrs=np.arange(len(eq)); ax1.plot(eq.index,1.12**(yrs/12.),ls="--",color="black",lw=0.9,label="12% target")
ax1.legend(loc="upper left",fontsize=9); ax1.grid(True,which="both",alpha=0.25)
e=eq.iloc[:,0]; dd=e/e.cummax()-1
ax2.fill_between(dd.index,dd*100,0,color="#2ca02c",alpha=0.5); ax2.set_ylabel("growth+moat\nDD (%)"); ax2.grid(True,alpha=0.25)
plt.tight_layout(); fig.savefig(f"{S}/backtest_growth.png",dpi=130); print("wrote backtest_growth.png")

def stats(r,ppy=12):
    eq=(1+r).cumprod(); y=len(r)/ppy; dd=eq/eq.cummax()-1; roll=eq.pct_change(ppy).dropna()
    return dict(CAGR=eq.iloc[-1]**(1/y)-1,Vol=r.std(ddof=1)*np.sqrt(ppy),
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)),MaxDD=dd.min(),
                Pct_1y_ge_12pct=(roll>=0.12).mean(),FinalNAV=eq.iloc[-1])
st=pd.DataFrame({c:stats(ret[c]) for c in ret.columns}).T
for c in ["CAGR","Vol","MaxDD","Pct_1y_ge_12pct"]: st[c]=(st[c]*100).round(2)
st["Sharpe"]=st["Sharpe"].round(2); st["FinalNAV"]=st["FinalNAV"].round(2)

g=pd.read_parquet(f"{S}/growth_select.parquet"); g["Instrument"]=g["Instrument"].astype(str); g["Date"]=pd.to_datetime(g["Date"])
last=g.sort_values("Date").groupby("Instrument").last()
ws=openpyxl.load_workbook(f"{S}/Universe_final.xlsx",data_only=True,read_only=True)["Scored"]
rows=list(ws.iter_rows(values_only=True)); ci={x:i for i,x in enumerate(rows[0])}
nm={str(r[ci['Ticker']]):r[ci['Company']] for r in rows[1:] if r[ci['Ticker']] and 'Company' in ci}
mo={str(r[ci['Ticker']]):r[ci['CompanyMoat(v3.2)']] for r in rows[1:] if r[ci['Ticker']]}
h=json.load(open(f"{S}/bt_erbands_holdings_Q_moat7.8_growth.json")); d=sorted(h)[-1]
def row(t):
    r=last.loc[t] if t in last.index else None
    e=r.er1*100 if r is not None else np.nan; c=r.carry_raw*100 if r is not None else np.nan
    elig = bool(r is not None and r.er1>=0.12 and r.carry_raw<=0.50 and not bool(r.artifact))
    return {"Ticker":t,"Company":nm.get(t,""),"Moat":mo.get(t),"ER_total_%":round(e,1),
            "Internal_growth_%":round(min(c,999.9),1),"Rerate_%":round(e-c,1) if r is not None else None,
            "eligible_now":elig}
hold=pd.DataFrame([row(t) for t in h[d]]).sort_values(["eligible_now","Internal_growth_%"],ascending=[False,False])

with pd.ExcelWriter(f"{S}/Backtest_internal_growth.xlsx",engine="openpyxl") as xl:
    st[["CAGR","Vol","Sharpe","MaxDD","Pct_1y_ge_12pct","FinalNAV"]].to_excel(xl,sheet_name="Stats_Monthly")
    hold.to_excel(xl,sheet_name=f"Growth_moat_book_{d}",index=False)
    ret.to_excel(xl,sheet_name="Monthly_Returns"); eq.to_excel(xl,sheet_name="Equity_Curves")
print("wrote Backtest_internal_growth.xlsx\n")
print(st[["CAGR","Vol","Sharpe","MaxDD","Pct_1y_ge_12pct","FinalNAV"]].to_string())
print(f"\nInternal-growth + moat>7.8 book ({d}):"); print(hold.to_string(index=False))
