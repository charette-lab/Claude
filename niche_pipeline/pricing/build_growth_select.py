#!/usr/bin/env python3
"""build_growth_select.py — the internal-growth selection signal, per spec.

Three steps, exactly as specified:
  1. Expected return as usual = full-engine er1 (5yr price->intrinsic, terminal
     value included, supply/demand-normalized).  Gate 1 = er1 >= 12%.
  2. Decompose er1 into the engine's own carry (internal growth) and rerate
     (revaluation).  carry = daily_expected_return_full er1 - rerate; here carry is
     the engine's er1_carry from carry_grid_norm (full engine).
  3. Select for internal-growth VALUE = rank the Gate-1 survivors by carry.

So the RANK metric is the carry; eligibility requires the usual total-ER gate and
no artifact, plus the engine's plausibility ceiling on carry (<= MAX_PLAUSIBLE_IRR)
to drop IRR-solver-ceiling artifacts. Names that fail any of these get carry=NaN
(excluded). Output month-end grid: [Instrument, Date, market_cap, expected_return
(=carry rank metric, gated), er1, carry_raw, artifact].

  python3 pricing/build_growth_select.py
"""
import os
import numpy as np, pandas as pd
import sys
sys.path.insert(0, "/home/user/Claude/niche_pipeline")
from frameworks import GATE1_IRR, MAX_PLAUSIBLE_IRR

SCR = "/tmp/claude-0/-home-user-Claude/58dd72fb-993d-5f0c-ab76-69d17b8d5d70/scratchpad"

tot = pd.read_parquet(SCR+"/daily_expected_return_full.parquet")[["Instrument", "Date", "expected_return", "artifact"]]
tot = tot.rename(columns={"expected_return": "er1"})
tot["Instrument"] = tot["Instrument"].astype(str); tot["Date"] = pd.to_datetime(tot["Date"]).astype("datetime64[ns]")
car = pd.read_parquet(SCR+"/carry_grid_norm.parquet")[["Instrument", "Date", "market_cap", "expected_return"]]
car = car.rename(columns={"expected_return": "carry"})
car["Instrument"] = car["Instrument"].astype(str); car["Date"] = pd.to_datetime(car["Date"]).astype("datetime64[ns]")

# attach the most-recent total er1 + artifact to each month-end carry obs
car = car.sort_values("Date"); tot = tot.sort_values("Date")
m = pd.merge_asof(car, tot, on="Date", by="Instrument", direction="backward")
m["artifact"] = m["artifact"].fillna(False).astype(bool)

# Step 3 eligibility: usual Gate 1 on er1, no artifact, plausible carry; rank = carry
eligible = (m["er1"] >= GATE1_IRR) & (~m["artifact"]) & (m["carry"] <= MAX_PLAUSIBLE_IRR) & m["er1"].notna() & m["carry"].notna()
m["carry_raw"] = m["carry"]
# ineligible -> a non-NaN sentinel (-9.99) so the backtest's as-of forward-fill PROPAGATES
# the ineligibility instead of resurrecting the last eligible carry across a NaN gap. The
# growth-mode gate (-1.0) sits below any real carry (IRR floor -0.95) and above the sentinel.
m["expected_return"] = np.where(eligible, m["carry"], -9.99)    # rank metric, pre-gated
out = m[["Instrument", "Date", "market_cap", "expected_return", "er1", "carry_raw", "artifact"]].sort_values(["Instrument", "Date"])
out.to_parquet(SCR+"/growth_select.parquet", index=False)
print(f"WROTE {len(out)} rows / {out.Instrument.nunique()} securities -> growth_select.parquet")
print(f"eligible (Gate1 er1>=12%, plausible carry, non-artifact) obs: {int(eligible.sum())} "
      f"({eligible.mean()*100:.0f}% of month-end obs)")
print(f"of eligible: median carry {m.loc[eligible,'carry'].median()*100:.1f}%  "
      f"median er1 {m.loc[eligible,'er1'].median()*100:.1f}%")
print("DONE_GROWTHSEL")
