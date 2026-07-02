#!/usr/bin/env python3
import json, glob, os
import numpy as np, pandas as pd, openpyxl
import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
S="/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U="/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"

def stats(r,ppy=4):
    r=pd.Series(r).dropna(); eq=(1+r).cumprod(); y=len(r)/ppy; dd=eq/eq.cummax()-1; roll=eq.pct_change(ppy).dropna()
    return dict(CAGR=eq.iloc[-1]**(1/y)-1,Vol=r.std(ddof=1)*np.sqrt(ppy),Sharpe=r.mean()*ppy/(r.std(ddof=1)*np.sqrt(ppy)),
                MaxDD=dd.min(),Calmar=(eq.iloc[-1]**(1/y)-1)/abs(dd.min()),Pct1y=(roll>=0.12).mean(),FinalNAV=eq.iloc[-1])

mine=pd.read_parquet(f"{S}/bt_core30_returns_Q.parquet")["Core30_resilient"]
fw=pd.read_parquet(f"{S}/bt_returns_Q.parquet")
cols={"My Core-30 (resilient, risk-adjusted)":mine,
      "Framework slot-cap (full engine)":fw["Core30_framework"],
      "Universe EW (benchmark)":fw["Universe_EW"]}
mp=f"{S}/bt_erbands_returns_Q_moat7.8_full.parquet"
if os.path.exists(mp): cols["Total-ER + moat>7.8"]=pd.read_parquet(mp)["ER_bands"]
ret=pd.DataFrame(cols).dropna(); eq=(1+ret).cumprod()
st=pd.DataFrame({c:stats(ret[c]) for c in ret.columns}).T
for c in ["CAGR","Vol","MaxDD","Pct1y"]: st[c]=(st[c]*100).round(1)
for c in ["Sharpe","Calmar","FinalNAV"]: st[c]=st[c].round(2)

COL=["#1f77b4","#ff7f0e","#7f7f7f","#2ca02c"]
fig,(ax1,ax2)=plt.subplots(2,1,figsize=(11,8.5),height_ratios=[3,1.4],sharex=True)
for c,cl in zip(ret.columns,COL[:len(ret.columns)]): ax1.plot(eq.index,eq[c],label=c,color=cl,lw=2.4 if "My Core" in c else 1.5)
ax1.set_yscale("log"); ax1.set_ylabel("Growth of 1.0 (log)")
ax1.set_title("My Core-30 'Resilient Quality Compounder' vs benchmarks (quarterly, full engine)\n"
              "risk-adjusted ER selection + inverse-vol weighting + factor/country caps + severity-4 screen",fontsize=10.5)
yrs=np.arange(len(eq)); ax1.plot(eq.index,1.12**(yrs/4.),ls="--",color="black",lw=0.9,label="12% target")
ax1.legend(loc="upper left",fontsize=9); ax1.grid(True,which="both",alpha=0.25)
e=eq["My Core-30 (resilient, risk-adjusted)"]; dd=e/e.cummax()-1
ax2.fill_between(dd.index,dd*100,0,color="#1f77b4",alpha=0.5); ax2.set_ylabel("My Core-30\ndrawdown (%)"); ax2.grid(True,alpha=0.25)
plt.tight_layout(); fig.savefig(f"{S}/backtest_core30.png",dpi=130); print("wrote backtest_core30.png")

# current book with inverse-vol weights
holds=json.load(open(f"{S}/bt_core30_holds_Q.json")); d=sorted(holds)[-1]; book=holds[d]
px=pd.concat([pd.read_parquet(f,columns=["Instrument","Date","Close Price"]) for f in
   [U+"972f0581-daily_volume_price_0.parquet",U+"257124b3-daily_volume_price_1.parquet",
    U+"86e54ec3-daily_volume_price_2.parquet",U+"13f82c18-daily_volume_price_0630.parquet"]])
px["Instrument"]=px["Instrument"].astype(str); px=px[px.Instrument.isin(book)]
px["Date"]=pd.to_datetime(px["Date"]); px["c"]=pd.to_numeric(px["Close Price"],errors="coerce")
W=px.pivot_table(index="Date",columns="Instrument",values="c",aggfunc="last").sort_index()
vol=(W.pct_change().rolling(126,min_periods=60).std()*np.sqrt(252)).reindex([pd.Timestamp(d)],method="ffill").iloc[0]
dec=pd.read_parquet(f"{S}/daily_return_decomposition.parquet")[["Instrument","Date","er_total"]]
dec["Instrument"]=dec["Instrument"].astype(str); dec["Date"]=pd.to_datetime(dec["Date"])
erl=dec[dec.Date<=pd.Timestamp(d)].sort_values("Date").groupby("Instrument").last()["er_total"]
ws=openpyxl.load_workbook(f"{S}/Universe_final.xlsx",data_only=True,read_only=True)["Scored"]
rows=list(ws.iter_rows(values_only=True)); ci={x:i for i,x in enumerate(rows[0])}
nm={str(r[ci['Ticker']]):r[ci['Company']] for r in rows[1:] if r[ci['Ticker']] and 'Company' in ci}
mo={str(r[ci['Ticker']]):r[ci['CompanyMoat(v3.2)']] for r in rows[1:] if r[ci['Ticker']]}
co={str(r[ci['Ticker']]):r[ci['Country']] for r in rows[1:] if r[ci['Ticker']]}
iv=np.array([1.0/float(vol[t]) if t in vol and pd.notna(vol[t]) and vol[t]>0 else np.nan for t in book])
w=iv/np.nansum(iv); w=np.clip(w,0.01,0.075); w=w/np.nansum(w)
hb=pd.DataFrame([{"Ticker":t,"Company":nm.get(t,"")[:32],"Country":co.get(t),"Moat":mo.get(t),
   "ER_total_%":round(float(erl.get(t,np.nan))*100,1),"Vol_%":round(float(vol[t])*100,1) if t in vol and pd.notna(vol[t]) else None,
   "Weight_%":round(float(w[i])*100,1)} for i,t in enumerate(book)]).sort_values("Weight_%",ascending=False)

with pd.ExcelWriter(f"{S}/Core30_resilient.xlsx",engine="openpyxl") as xl:
    st[["CAGR","Vol","Sharpe","MaxDD","Calmar","Pct1y","FinalNAV"]].to_excel(xl,sheet_name="Stats_Quarterly")
    hb.to_excel(xl,sheet_name=f"Book_{d}",index=False)
    ret.to_excel(xl,sheet_name="Quarterly_Returns"); eq.to_excel(xl,sheet_name="Equity_Curves")
print("wrote Core30_resilient.xlsx\n")
print(st[["CAGR","Vol","Sharpe","MaxDD","Calmar","Pct1y","FinalNAV"]].to_string())
print(f"\nCurrent book ({d}), {len(book)} names, inverse-vol weighted:")
print(hb.to_string(index=False))
