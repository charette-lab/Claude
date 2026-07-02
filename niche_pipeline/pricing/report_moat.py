#!/usr/bin/env python3
import json, pandas as pd, numpy as np, openpyxl
import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
S="/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"

# returns
base=pd.read_parquet(f"{S}/bt_erbands_returns_M.parquet")["ER_bands"]
moat=pd.read_parquet(f"{S}/bt_erbands_returns_M_moat7.8.parquet")["ER_bands"]
prior=pd.read_parquet(f"{S}/bt_returns_M.parquet")
ret=pd.DataFrame({
    "ER-only + bands, moat>7.8": moat,
    "ER-only + bands (full universe)": base,
    "Tagged universe EW (benchmark)": prior["Universe_EW"],
}).dropna()
eq=(1+ret).cumprod(); COL=["#9467bd","#d62728","#7f7f7f"]
fig,(ax1,ax2)=plt.subplots(2,1,figsize=(11,8.5),height_ratios=[3,1.4],sharex=True)
for c,col in zip(ret.columns,COL):
    ax1.plot(eq.index,eq[c],label=c,color=col,lw=2.3 if "moat" in c else 1.6)
ax1.set_yscale("log"); ax1.set_ylabel("Growth of 1.0 (log)")
ax1.set_title("ER selector + IPS bands: top-tier moat (>7.8) vs full universe\n"
              "(monthly checkup; ER>=12%; trim@5% / liquidate@2%; price-only, no costs)",fontsize=11)
yrs=np.arange(len(eq)); ax1.plot(eq.index,1.12**(yrs/12.),ls="--",color="green",lw=1,label="12% target")
ax1.legend(loc="upper left",fontsize=9); ax1.grid(True,which="both",alpha=0.25)
e=eq.iloc[:,0]; dd=e/e.cummax()-1
ax2.fill_between(dd.index,dd*100,0,color="#9467bd",alpha=0.5); ax2.set_ylabel("moat>7.8\ndrawdown (%)"); ax2.grid(True,alpha=0.25)
plt.tight_layout(); fig.savefig(f"{S}/backtest_moat.png",dpi=130); print("wrote backtest_moat.png")

def stats(r,ppy=12):
    eq=(1+r).cumprod(); y=len(r)/ppy; dd=eq/eq.cummax()-1; roll=eq.pct_change(ppy).dropna()
    return dict(CAGR=eq.iloc[-1]**(1/y)-1,Vol=r.std(ddof=1)*np.sqrt(ppy),
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)),MaxDD=dd.min(),
                Pct_1y_ge_12pct=(roll>=0.12).mean(),FinalNAV=eq.iloc[-1])
st=pd.DataFrame({c:stats(ret[c]) for c in ret.columns}).T
for c in ["CAGR","Vol","MaxDD","Pct_1y_ge_12pct"]: st[c]=(st[c]*100).round(2)
st["Sharpe"]=st["Sharpe"].round(2); st["FinalNAV"]=st["FinalNAV"].round(2)

# CURRENT holdings of the moat>7.8 book, annotated with moat + latest ER + name
h=json.load(open(f"{S}/bt_erbands_holdings_Q_moat7.8.json"))
last=sorted(h)[-1]; names=h[last]
ws=openpyxl.load_workbook(f"{S}/Universe_final.xlsx",data_only=True,read_only=True)["Scored"]
rows=list(ws.iter_rows(values_only=True)); ci={x:i for i,x in enumerate(rows[0])}
moatm={str(r[ci['Ticker']]):r[ci['CompanyMoat(v3.2)']] for r in rows[1:] if r[ci['Ticker']]}
namem={str(r[ci['Ticker']]):r[ci['Company']] for r in rows[1:] if r[ci['Ticker']] and 'Company' in ci}
er=pd.read_parquet(f"{S}/daily_expected_return.parquet"); er['Instrument']=er['Instrument'].astype(str)
erl=er[er.Date<=pd.Timestamp(last)].sort_values('Date').groupby('Instrument').last()['expected_return']
hold=pd.DataFrame([{"Ticker":t,"Company":namem.get(t,""),"Moat":moatm.get(t),
                    "ER_%":round(float(erl.get(t,np.nan))*100,1)} for t in names]).sort_values("ER_%",ascending=False)

with pd.ExcelWriter(f"{S}/Backtest_moat78.xlsx",engine="openpyxl") as xl:
    st[["CAGR","Vol","Sharpe","MaxDD","Pct_1y_ge_12pct","FinalNAV"]].to_excel(xl,sheet_name="Stats_Monthly")
    hold.to_excel(xl,sheet_name=f"Holdings_{last}",index=False)
    ret.to_excel(xl,sheet_name="Monthly_Returns")
    eq.to_excel(xl,sheet_name="Equity_Curves")
print("wrote Backtest_moat78.xlsx\n")
print(st[["CAGR","Vol","Sharpe","MaxDD","Pct_1y_ge_12pct","FinalNAV"]].to_string())
print(f"\nCurrent moat>7.8 book ({last}), {len(names)} names:")
print(hold.to_string(index=False))
