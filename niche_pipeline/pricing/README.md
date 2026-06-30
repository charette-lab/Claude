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

- `backtest.py` — evaluates the Constrained Quality Compounder Index through time
  (Gate-2 downside hard-stop DISABLED). Investable universe = the high-moat names
  that carry the 11 binary risk tags (severity>=3 from severity_master / notes,
  else research-cache native binary). At each rebalance date: keep names with
  Expected IRR >= 12% (Gate 1), rank by ER, fill 30 equal-weighted names under the
  6-of-30 slot cap (no risk tag in >6 names). Realized return is close-to-close,
  local currency, price-only. Benchmarks: Top-30 by ER without the slot cap, all
  ER>=12% equal-weight, and the tagged-universe equal-weight. `report_bt.py` builds
  the equity chart + `Backtest_CQCI.xlsx`.

  Findings (2015-2026, monthly rebalance): the full framework returns ~20.9% CAGR
  (Sharpe 1.24, max DD -22.6%) vs ~14.9% for the universe benchmark and ~16.8% for
  the gate alone — it clears the 12% design target with a wide margin (71% of
  rolling-1y windows >= 12%). The slot cap trades ~4 pts of CAGR for ~4 pts less
  drawdown and the best monthly Sharpe: it is doing risk control, not return
  enhancement. Quarterly (IPS-native) cadence confirms the ranking.

  CAVEATS: local-currency price returns (no dividends, no FX translation, no
  transaction costs); delisted names carry last close (mildly optimistic). The
  robust signal is the RELATIVE ranking (framework vs benchmarks), invariant to FX.

- `backtest_er_bands.py` — pure-ER variant: the SELECTOR is Expected IRR alone
  (risk tags ignored, since tag coverage is incomplete), so the universe is ALL
  3,121 priced ER securities. The rest of the IPS is retained — 30 EW names,
  Gate-1 ER>=12%, and the Phase-4 Quarterly Checkup with Tolerance Bands: trim a
  name back to 3.33% when it drifts to >=5% (Trigger 2), liquidate and replace
  with the highest-ER non-held name when it bleeds to <=2% (Trigger 3). Trigger 1
  (the risk-tag ceiling) is dropped. Because the BANDS, not the calendar, drive
  trades, turnover collapses to ~15-17%/yr (~4-5 names/yr, ~6-7yr holding period)
  vs ~150%/yr for full reconstitution. Realized ~25-27% CAGR (Sharpe ~1.4, max DD
  -21%) — it beats both the slot-cap framework (~21%) and the benchmark (~15%):
  dropping the diversification constraint lets the highest-ER names compound, and
  the bands add a let-winners-run / cut-losers tilt. `report_erbands.py` builds
  the chart + `Backtest_ER_bands.xlsx`. Same return caveats as above (here the low
  turnover makes the no-cost assumption far less material).

NOTE: input/output paths are session-specific (uploads + scratchpad) — adjust
before re-running. Outputs are large parquet files (not committed).
Assumption: moat is held constant (panel has no per-year moat); all fundamentals
are fixed at the year-end valuation, only price varies daily.
