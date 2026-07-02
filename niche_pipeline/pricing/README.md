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

- `build_daily_extra.py` — values the names that live in the 132-col hist
  part-file (`hist_20260629_2.xlsx`) rather than the 30-year `.xlsb` panel (e.g.
  Workday and ~150 others that were absent from the panel, so the engine never
  valued them). Same rich schema (spaced column names), so it extracts per-year
  fundamentals identically, pairs them with the Universe_final moats + the 0630
  daily prices, builds daily MC/EV + ER on the same methodology, and APPENDS to
  the panels. Net effect: ER universe 3,121 -> 3,260, MC/EV 3,186 -> 3,338.

- `build_daily_er_full.py` — daily ER from the FULL engine. Adds the two layers
  the complete pipeline applies on top of the raw fade ER: (1) SUPPLY+DEMAND
  normalization (`overearning.two_stage_return`) — a capital-cycle scarcity rent
  is faded to the reproduction equilibrium and a genAI-compressed software moat
  fades faster, applied as a per-(security,fiscal-year) downgrade dn = er_current
  - er_adj, point-in-time (signals use only history up to each year-end); (2) an
  UNREASONABLE-RETURN screen (`frameworks.er_is_artifact`, quality-aware via
  roic_star + history verdict) flagging extrapolation artifacts. Output
  `daily_expected_return_full.parquet` [er_current, expected_return(=er_adj),
  artifact]. Run `backtest_er_bands.py --full` to select on er_adj and exclude
  artifacts. Effect (monthly): NVIDIA er 180%->131% and artifact-excluded, semicap
  scarcity rents (Lasertec, Advantest, Lam) faded; the book swaps high-octane names
  for steady compounders (Amazon, Intuitive Surgical, Heico, Safran). Realized
  ~22% CAGR / Sharpe 1.2 vs ~28% / 1.5 for the raw book — the engine's conservatism
  costs return in a 2016-26 boom that rewarded the over-earning it guards against,
  but still clears 12% in ~72-80% of rolling years and beats the ~16% benchmark.

- `build_carry_grid.py` / `build_carry_norm.py` — CARRY signal: the engine's
  internal-compounding return (`er1_carry` — the IRR if the price never re-rates,
  exiting at the entry multiple), at month-end dates, then supply/demand-normalized
  (the normalization barely moves carry — ~0.02pp — it hits the rerate/terminal, not
  the near-term carry). `backtest_er_bands.py --carry` selects on it (gate carry>=12%,
  rank by carry). This builds the book on INTERNAL COMPOUNDING rather than re-rating.
  Result: ~12.7% CAGR full universe / ~18% with moat>7.8 (Sharpe 1.08, DD -22%,
  ~2 names/yr turnover) — below the total-ER books (22-28%), which quantifies how
  much of the period's return was re-rating the carry book forgoes. The carry book
  tilts to high-cash-return infrastructure (regulated utilities, water, gas, airports)
  — the businesses whose return genuinely comes from internal cash compounding.

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

- `backtest_core30.py` / `report_core30.py` — my own "Resilient Quality Compounder"
  Core-30, built to maximise return AND minimise drawdown. Eligibility: full-engine
  er_total>=12% & non-artifact, moat>=6.5, carries risk tags with NO Extreme
  (severity-4) risk, >=60d price history. RANK by RISK-ADJUSTED return (er_total/vol).
  Select 30 under a 6-of-30 risk-tag slot cap + <=10/country (correlation caps).
  Weight INVERSE-VOLATILITY, clipped [1%,7.5%]. Quarterly reconstitution; sell on
  loss of eligibility. Result (Q): ~17% CAGR, Sharpe 1.00, max DD -19.5%, Calmar 0.87
  — the lowest drawdown and best return-per-drawdown of any book built (vs framework
  slot-cap 22.5%/DD-28.5%/Calmar 0.79 and the -26 to -29% drawdowns elsewhere).
