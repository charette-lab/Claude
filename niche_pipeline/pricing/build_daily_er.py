import sys, pandas as pd, numpy as np
sys.path.insert(0,"/home/user/Claude/niche_pipeline")
import panel30, aip, openpyxl
SCR="/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U="/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB=[U+"fb2aec33-30_file_1.xlsb",U+"a810c35f-30_file_2.xlsb",U+"62545fc2-30_file_3.xlsb"]
TEST = "--test" in sys.argv

# researched moat (constant per security)
ws=openpyxl.load_workbook(SCR+"/Universe_final.xlsx",data_only=True,read_only=True)["Scored"]
rows=list(ws.iter_rows(values_only=True)); ci={x:i for i,x in enumerate(rows[0])}
moat={r[ci["Ticker"]]: r[ci["CompanyMoat(v3.2)"]] for r in rows[1:]
      if r[ci["Ticker"]] and isinstance(r[ci["CompanyMoat(v3.2)"]],(int,float))}
ind={r[ci["Ticker"]]: r[ci["Industry"]] for r in rows[1:] if r[ci["Ticker"]]}
ctry={r[ci["Ticker"]]: r[ci["Country"]] for r in rows[1:] if r[ci["Ticker"]]}

print("loading panel...",flush=True)
by,idx=panel30.load(XLSB)
g=lambda r,k:(r[idx[k]] if k in idx and idx[k]<len(r) else None)
def num(v): return v if isinstance(v,(int,float)) else None
# per (security, ped) fundamentals
fin_rows=[]
for t,rs in by.items():
    t=str(t); m=moat.get(t)
    if m is None: continue
    for r in rs:
        ped=g(r,"Period_End_Date"); mc=num(g(r,"Market Capitalization")); ev=num(g(r,"EV"))
        nopat=num(g(r,"New Operating Income")); roiic=num(g(r,"ROICm_total - 7 years"))
        rr=num(g(r,"average_C - 7 year")); sales=num(g(r,"Sales"))
        tax=num(g(r,"Income_Tax_Rate___Pct"))
        ltd=num(g(r,"Debt - Long-Term - Total")) or 0.0
        std=num(g(r,"Short-Term Debt & Current Portion of Long-Term Debt")) or 0.0
        if ped is None or mc is None or ev is None or nopat is None or roiic is None or rr is None: continue
        fin_rows.append({"Instrument":t,"ped":ped,"moat":m,"nopat":nopat,"roiic":roiic,"rr":rr,
                         "sales":sales,"tax":(tax/100.0 if tax is not None else 0.25),
                         "gross":ltd+std,"netdebt":ev-mc,"mc_ye":mc,"ind":ind.get(t),"ctry":ctry.get(t)})
F=pd.DataFrame(fin_rows); F["ped"]=pd.to_datetime(F["ped"],errors="coerce").astype("datetime64[ns]")
F=F.dropna(subset=["ped"]).sort_values("ped")
print(f"per-(security,year) fins: {len(F)} across {F.Instrument.nunique()} securities",flush=True)

# daily market cap
daily=pd.read_parquet(SCR+"/daily_marketcap_ev.parquet",columns=["Instrument","Date","market_cap"])
daily["Date"]=pd.to_datetime(daily["Date"]).astype("datetime64[ns]"); daily["Instrument"]=daily["Instrument"].astype(str)
if TEST:
    keep=list(F.Instrument.unique())[:8]
    daily=daily[daily.Instrument.isin(keep)]; F=F[F.Instrument.isin(keep)]
    print("TEST mode, securities:",keep)

# attach the applicable fiscal-year valuation (most recent ped <= date) to each daily row
daily=daily.sort_values("Date")
m=pd.merge_asof(daily, F.sort_values("ped"), left_on="Date", right_on="ped", by="Instrument", direction="backward")
m=m.dropna(subset=["ped","mc_ye"])
print(f"daily rows with a valuation anchor: {len(m)}",flush=True)

def er_at(finbase, mc):
    f=dict(finbase); f["Market Cap"]=mc
    v=aip.value_and_return(f, re=0.07, re2=0.12)
    return (v.get("er1") if v else None)

# per (Instrument, ped) group: sample er over the group's mc range, interpolate to daily
out=[]
NP=20
for (t,ped),grp in m.groupby(["Instrument","ped"]):
    r0=grp.iloc[0]
    base={"Company Name":t,"Instrument":t,"GICS Industry Group Name":r0["ind"],
          "Country of Headquarters":r0["ctry"],"New Operating Income":r0["nopat"],
          "ROICm 7":r0["roiic"],"RR 7":r0["rr"],"Moat Score":r0["moat"],"Gross debt":r0["gross"],
          "Income Tax Rate - Instrument":r0["tax"],"Net debt":r0["netdebt"],"Sales":r0["sales"]}
    mcs=grp["market_cap"].values
    lo,hi=np.nanmin(mcs),np.nanmax(mcs)
    if not (lo>0): 
        continue
    grid=np.unique(np.linspace(lo,hi,NP)) if hi>lo else np.array([lo])
    ers=[er_at(base,x) for x in grid]
    gd=[(x,e) for x,e in zip(grid,ers) if e is not None]
    if not gd: continue
    gx=np.array([x for x,_ in gd]); ge=np.array([e for _,e in gd])
    er_daily=np.interp(mcs, gx, ge)
    out.append(pd.DataFrame({"Instrument":t,"Date":grp["Date"].values,
                             "market_cap":mcs,"expected_return":er_daily}))
res=pd.concat(out,ignore_index=True).sort_values(["Instrument","Date"])
res.to_parquet(SCR+"/daily_expected_return.parquet",index=False)
print(f"WROTE {len(res)} rows / {res.Instrument.nunique()} securities",flush=True)
# validate: interp vs direct engine at the SAME daily price
print("\nVALIDATION (interp vs direct engine at identical price):")
for t in res.Instrument.unique()[:5]:
    ff=F[F.Instrument==t].iloc[-1]
    base={"Company Name":t,"Instrument":t,"GICS Industry Group Name":ff["ind"],"Country of Headquarters":ff["ctry"],
          "New Operating Income":ff["nopat"],"ROICm 7":ff["roiic"],"RR 7":ff["rr"],"Moat Score":ff["moat"],
          "Gross debt":ff["gross"],"Income Tax Rate - Instrument":ff["tax"],"Net debt":ff["netdebt"],"Sales":ff["sales"]}
    row=res[res.Instrument==t].iloc[-1]
    mc=float(row["market_cap"]); direct=er_at(base,mc)
    interp_s="%.2f%%"%(row["expected_return"]*100)
    direct_s=("%.2f%%"%(direct*100)) if direct is not None else "NA"
    print("  %s: interp er=%s  direct@same-MC=%s  (MC=%.0f)"%(t,interp_s,direct_s,mc))
print("DONE_ER")
