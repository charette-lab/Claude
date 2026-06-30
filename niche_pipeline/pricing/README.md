# Pricing — daily market cap, EV, and expected return

Builds daily series from the yearly panel + the daily price/volume parquet.

- `build_daily_mcev.py` — daily market cap & enterprise value.
  shares(year)=yearly MktCap/close@period-end; net_debt(year)=yearly EV-MktCap;
  daily MktCap=shares*daily close; daily EV=daily MktCap+net_debt (step-held per fiscal year).
- `build_daily_er.py` — daily expected return. Values each security with the aip
  engine at each fiscal year-end (panel fundamentals + researched moat, held
  constant), carries that valuation forward to the next year, and computes the
  daily ER as the engine's er@7% with that day's market cap as entry price.
  ER is monotonic in price, so sampled at 20 price points per (security,year) and
  interpolated to daily — verified exact (<=0.01%) vs direct engine calls.

NOTE: input/output paths are session-specific (uploads + scratchpad) — adjust
before re-running. Outputs are large parquet files (not committed).
Assumption: moat is held constant (panel has no per-year moat); all fundamentals
are fixed at the year-end valuation, only price varies daily.
