#!/usr/bin/env python3
import json, pandas as pd, numpy as np, openpyxl
import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
S="/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"

def col(p,c="ER_bands"): return pd.read_parquet(f"{S}/{p}")[c]
ret=pd.DataFrame({
    "Full engine, moat>7.8":        col("bt_erbands_returns_M_moat7.8_full.parquet"),
    "Full engine, full universe":   col("bt_erbands_returns_M_full.parquet"),
    "Raw ER, full universe":        col("bt_erbands_returns_M.parquet"),
    "Universe EW (benchmark)":      pd.read_parquet(f"{S}/bt_returns_M.parquet")["Universe_EW"],
}).dropna()
eq=(1+ret).cumprod(); COL=["#2ca02c","#9467bd","#d62728","#7f7f7f"]
fig,(ax1,ax2)=plt.subplots(2,1,figsize=(11,8.5),height_ratios=[3,1.4],sharex=True)
for c,cl in zip(ret.columns,COL): ax1.plot(eq.index,eq[c],label=c,color=cl,lw=2.3 if "Full engine, moat" in c else 1.6)
ax1.set_yscale("log"); ax1.set_ylabel("Growth of 1.0 (log)")
ax1.set_title("Full valuation engine (supply/demand-normalized ER + artifact screen) vs raw ER\n"
              "(monthly checkup; ER>=12%; trim@5% / liquidate@2%; price-only, no costs)",fontsize=11)
yrs=np.arange(len(eq)); ax1.plot(eq.index,1.12**(yrs/12.),ls="--",color="black",lw=0.9,label="12% target")
ax1.legend(loc="upper left",fontsize=9); ax1.grid(True,which="both",alpha=0.25)
e=eq.iloc[:,0]; dd=e/e.cummax()-1
ax2.fill_between(dd.index,dd*100,0,color="#2ca02c",alpha=0.5); ax2.set_ylabel("Full-engine\nmoat>7.8 DD (%)"); ax2.grid(True,alpha=0.25)
plt.tight_layout(); fig.savefig(f"{S}/backtest_full.png",dpi=130); print("wrote backtest_full.png")

def stats(r,ppy=12):
    eq=(1+r).cumprod(); y=len(r)/ppy; dd=eq/eq.cummax()-1; roll=eq.pct_change(ppy).dropna()
    return dict(CAGR=eq.iloc[-1]**(1/y)-1,Vol=r.std(ddof=1)*np.sqrt(ppy),
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)),MaxDD=dd.min(),
                Pct_1y_ge_12pct=(roll>=0.12).mean(),FinalNAV=eq.iloc[-1])
st=pd.DataFrame({c:stats(ret[c]) for c in ret.columns}).T
for c in ["CAGR","Vol","MaxDD","Pct_1y_ge_12pct"]: st[c]=(st[c]*100).round(2)
st["Sharpe"]=st["Sharpe"].round(2); st["FinalNAV"]=st["FinalNAV"].round(2)

# current full-engine moat>7.8 holdings, annotated
full=pd.read_parquet(f"{S}/daily_expected_return_full.parquet"); full["Instrument"]=full["Instrument"].astype(str)
latest=full.sort_values("Date").groupby("Instrument").last()
ws=openpyxl.load_workbook(f"{S}/Universe_final.xlsx",data_only=True,read_only=True)["Scored"]
rows=list(ws.iter_rows(values_only=True)); ci={x:i for i,x in enumerate(rows[0])}
moatm={str(r[ci['Ticker']]):r[ci['CompanyMoat(v3.2)']] for r in rows[1:] if r[ci['Ticker']]}
namem={str(r[ci['Ticker']]):r[ci['Company']] for r in rows[1:] if r[ci['Ticker']] and 'Company' in ci}
h=json.load(open(f"{S}/bt_erbands_holdings_Q_moat7.8_full.json")); last=sorted(h)[-1]
def annot(t):
    r=latest.loc[t] if t in latest.index else None
    return {"Ticker":t,"Company":namem.get(t,""),"Moat":moatm.get(t),
            "ER_raw_%":round(r.er_current*100,1) if r is not None else None,
            "ER_adj_%":round(r.expected_return*100,1) if r is not None else None}
hold=pd.DataFrame([annot(t) for t in h[last]]).sort_values("ER_adj_%",ascending=False)

# what the engine did: biggest downgrades + artifact-excluded, among moat>7.8 names
hi=[t for t,m in moatm.items() if isinstance(m,(int,float)) and m>7.8 and t in latest.index]
eng=latest.loc[hi].copy()
eng["downgrade_pp"]=((eng.er_current-eng.expected_return)*100).round(1)
eng["Company"]=[namem.get(t,"") for t in eng.index]
eng["Moat"]=[moatm.get(t) for t in eng.index]
eng["ER_raw_%"]=(eng.er_current*100).round(1); eng["ER_adj_%"]=(eng.expected_return*100).round(1)
topdown=eng.sort_values("downgrade_pp",ascending=False).head(20)[["Company","Moat","ER_raw_%","ER_adj_%","downgrade_pp","artifact"]]
arti=eng[eng.artifact].sort_values("ER_raw_%",ascending=False)[["Company","Moat","ER_raw_%","ER_adj_%"]].head(25)

with pd.ExcelWriter(f"{S}/Backtest_full_engine.xlsx",engine="openpyxl") as xl:
    st[["CAGR","Vol","Sharpe","MaxDD","Pct_1y_ge_12pct","FinalNAV"]].to_excel(xl,sheet_name="Stats_Monthly")
    hold.to_excel(xl,sheet_name=f"FullEngine_book_{last}",index=False)
    topdown.to_excel(xl,sheet_name="Biggest_downgrades")
    arti.to_excel(xl,sheet_name="Artifact_excluded")
    ret.to_excel(xl,sheet_name="Monthly_Returns"); eq.to_excel(xl,sheet_name="Equity_Curves")
print("wrote Backtest_full_engine.xlsx\n")
print(st[["CAGR","Vol","Sharpe","MaxDD","Pct_1y_ge_12pct","FinalNAV"]].to_string())
print(f"\nCurrent full-engine moat>7.8 book ({last}):"); print(hold.to_string(index=False))
print("\nArtifact-excluded (moat>7.8 names):"); print(arti.to_string())
