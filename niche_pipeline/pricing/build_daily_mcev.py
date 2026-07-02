import pandas as pd, pyarrow.dataset as ds, sys
sys.path.insert(0,"/home/user/Claude/niche_pipeline")
import panel30
SCR="/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"
U="/root/.claude/uploads/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/"
XLSB=[U+"fb2aec33-30_file_1.xlsb",U+"a810c35f-30_file_2.xlsb",U+"62545fc2-30_file_3.xlsb"]
PARQ=[U+"972f0581-daily_volume_price_0.parquet",U+"257124b3-daily_volume_price_1.parquet",U+"86e54ec3-daily_volume_price_2.parquet",
      U+"13f82c18-daily_volume_price_0630.parquet",
      # 30-year history (1995-2015), 7 batches
      U+"cf27bad3-two_batches_daily_volume_price_0.parquet",U+"c3621e8f-two_batches_daily_volume_price_1.parquet",
      U+"35e9a52e-two_batches_daily_volume_price_2.parquet",U+"e397aa9f-two_batches_daily_volume_price_3.parquet",
      U+"3cfd2814-two_batches_daily_volume_price_4.parquet",U+"25b0cf78-two_batches_daily_volume_price_5.parquet",
      U+"95c8416b-two_batches_daily_volume_price_6.parquet"]

print("loading yearly panel...",flush=True)
by,idx=panel30.load(XLSB)
mi=idx["Market Capitalization"]; ei=idx["EV"]; di=idx["Period_End_Date"]
rows=[(str(t),r[di],r[mi],r[ei]) for t,rs in by.items() for r in rs
      if isinstance(r[mi],(int,float)) and isinstance(r[ei],(int,float)) and r[di] is not None]
yr=pd.DataFrame(rows,columns=["Instrument","ped","mc","ev"])
yr["ped"]=pd.to_datetime(yr["ped"],errors="coerce").astype("datetime64[ns]")
yr["Instrument"]=yr["Instrument"].astype(str)
yr=yr.dropna(subset=["ped","mc","ev"])
yr=yr[(yr["mc"]>0)]

print("loading daily parquet...",flush=True)
daily=ds.dataset(PARQ,format="parquet").to_table(columns=["Instrument","Date","Close Price"]).to_pandas()
daily=daily.rename(columns={"Close Price":"close"})
daily["Date"]=pd.to_datetime(daily["Date"],errors="coerce").astype("datetime64[ns]")
daily["Instrument"]=daily["Instrument"].astype(str)
daily=daily.dropna(subset=["Date","close","Instrument"])
daily=daily[daily["close"]>0].drop_duplicates(["Instrument","Date"])
print(f"daily rows {len(daily)} | yearly obs {len(yr)}",flush=True)

# shares = MC / close at period-end (nearest trading day within 10d)
dd=daily.rename(columns={"Date":"d","close":"c"})[["Instrument","d","c"]].dropna().sort_values("d")
yr2=pd.merge_asof(yr.sort_values("ped"),dd,left_on="ped",right_on="d",by="Instrument",
                  direction="nearest",tolerance=pd.Timedelta("10D"))
yr2["shares"]=yr2["mc"]/yr2["c"]; yr2["net_debt"]=yr2["ev"]-yr2["mc"]
yr2=yr2.dropna(subset=["shares"])
step=yr2[["Instrument","ped","shares","net_debt"]].dropna(subset=["ped"]).sort_values("ped")

out=pd.merge_asof(daily.sort_values("Date"),step,left_on="Date",right_on="ped",by="Instrument",direction="backward")
out=out.sort_values(["Instrument","Date"])
out[["shares","net_debt"]]=out.groupby("Instrument")[["shares","net_debt"]].bfill()
out=out.dropna(subset=["shares"])
out["market_cap"]=out["shares"]*out["close"]
out["enterprise_value"]=out["market_cap"]+out["net_debt"]
out=out[["Instrument","Date","close","shares","net_debt","market_cap","enterprise_value"]]
out.to_parquet(SCR+"/daily_marketcap_ev.parquet",index=False)
print(f"WROTE {len(out)} rows / {out.Instrument.nunique()} instruments",flush=True)
ye=yr2[yr2.Instrument=="6737.T"][["ped","mc","ev","shares","net_debt"]].tail(3)
print("VALIDATION 6737.T yearly anchors:\n"+ye.to_string(index=False))
v=out[out.Instrument=="6737.T"]
m=v[(v.Date>="2025-03-26")&(v.Date<="2025-04-02")][["Date","close","shares","market_cap","enterprise_value"]]
print("daily near 2025-03-31 (yearly mc=610986828 ev=523430976):\n"+m.to_string(index=False))
print("DONE_BUILD")
