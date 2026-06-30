#!/usr/bin/env python3
import json, pandas as pd, numpy as np, openpyxl
import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
S="/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
def col(p): return pd.read_parquet(f"{S}/{p}")["ER_bands"]
ret=pd.DataFrame({
    "Carry + moat>7.8 (internal compounding)": col("bt_erbands_returns_M_moat7.8_carry.parquet"),
    "Carry, full universe":                    col("bt_erbands_returns_M_carry.parquet"),
    "Full engine ER, full universe":           col("bt_erbands_returns_M_full.parquet"),
    "Universe EW (benchmark)":                 pd.read_parquet(f"{S}/bt_returns_M.parquet")["Universe_EW"],
}).dropna()
eq=(1+ret).cumprod(); COL=["#1f77b4","#17becf","#d62728","#7f7f7f"]
fig,(ax1,ax2)=plt.subplots(2,1,figsize=(11,8.5),height_ratios=[3,1.4],sharex=True)
for c,cl in zip(ret.columns,COL): ax1.plot(eq.index,eq[c],label=c,color=cl,lw=2.4 if "Carry + moat" in c else 1.6)
ax1.set_yscale("log"); ax1.set_ylabel("Growth of 1.0 (log)")
ax1.set_title("Carry selector (internal compounding) vs total-ER selector\n"
              "(monthly checkup; gate carry>=12%; trim@5% / liquidate@2%; price-only, no costs)",fontsize=11)
yrs=np.arange(len(eq)); ax1.plot(eq.index,1.12**(yrs/12.),ls="--",color="black",lw=0.9,label="12% target")
ax1.legend(loc="upper left",fontsize=9); ax1.grid(True,which="both",alpha=0.25)
e=eq.iloc[:,0]; dd=e/e.cummax()-1
ax2.fill_between(dd.index,dd*100,0,color="#1f77b4",alpha=0.5); ax2.set_ylabel("Carry+moat\nDD (%)"); ax2.grid(True,alpha=0.25)
plt.tight_layout(); fig.savefig(f"{S}/backtest_carry.png",dpi=130); print("wrote backtest_carry.png")

def stats(r,ppy=12):
    eq=(1+r).cumprod(); y=len(r)/ppy; dd=eq/eq.cummax()-1; roll=eq.pct_change(ppy).dropna()
    return dict(CAGR=eq.iloc[-1]**(1/y)-1,Vol=r.std(ddof=1)*np.sqrt(ppy),
                Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)),MaxDD=dd.min(),
                Pct_1y_ge_12pct=(roll>=0.12).mean(),FinalNAV=eq.iloc[-1])
st=pd.DataFrame({c:stats(ret[c]) for c in ret.columns}).T
for c in ["CAGR","Vol","MaxDD","Pct_1y_ge_12pct"]: st[c]=(st[c]*100).round(2)
st["Sharpe"]=st["Sharpe"].round(2); st["FinalNAV"]=st["FinalNAV"].round(2)

# current Carry+moat>7.8 book: show carry vs total ER vs moat
cn=pd.read_parquet(f"{S}/carry_grid_norm.parquet"); cn["Instrument"]=cn["Instrument"].astype(str)
cn["Date"]=pd.to_datetime(cn["Date"]); cl=cn.sort_values("Date").groupby("Instrument").last()
full=pd.read_parquet(f"{S}/daily_expected_return_full.parquet"); full["Instrument"]=full["Instrument"].astype(str)
full["Date"]=pd.to_datetime(full["Date"]); fl=full.sort_values("Date").groupby("Instrument").last()
ws=openpyxl.load_workbook(f"{S}/Universe_final.xlsx",data_only=True,read_only=True)["Scored"]
rows=list(ws.iter_rows(values_only=True)); ci={x:i for i,x in enumerate(rows[0])}
moatm={str(r[ci['Ticker']]):r[ci['CompanyMoat(v3.2)']] for r in rows[1:] if r[ci['Ticker']]}
namem={str(r[ci['Ticker']]):r[ci['Company']] for r in rows[1:] if r[ci['Ticker']] and 'Company' in ci}
h=json.load(open(f"{S}/bt_erbands_holdings_Q_moat7.8_carry.json")); last=sorted(h)[-1]
hold=pd.DataFrame([{"Ticker":t,"Company":namem.get(t,""),"Moat":moatm.get(t),
    "Carry_%":round(cl.expected_return.get(t,np.nan)*100,1),
    "Total_ER_%":round(fl.expected_return.get(t,np.nan)*100,1)} for t in h[last]]).sort_values("Carry_%",ascending=False)

with pd.ExcelWriter(f"{S}/Backtest_carry.xlsx",engine="openpyxl") as xl:
    st[["CAGR","Vol","Sharpe","MaxDD","Pct_1y_ge_12pct","FinalNAV"]].to_excel(xl,sheet_name="Stats_Monthly")
    hold.to_excel(xl,sheet_name=f"Carry_moat_book_{last}",index=False)
    ret.to_excel(xl,sheet_name="Monthly_Returns"); eq.to_excel(xl,sheet_name="Equity_Curves")
print("wrote Backtest_carry.xlsx\n")
print(st[["CAGR","Vol","Sharpe","MaxDD","Pct_1y_ge_12pct","FinalNAV"]].to_string())
print(f"\nCurrent Carry + moat>7.8 book ({last}):"); print(hold.to_string(index=False))
