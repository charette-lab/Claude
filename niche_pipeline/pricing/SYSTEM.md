# Investment system — high-returning AND robust

A ground-up design using only the researched universe, built on what this session's
backtests actually demonstrated (not theory). Every choice below is tied to an
empirical result we produced.

## The five lessons the data taught us (the design constraints)

1. **Return in 2016–26 was overwhelmingly RE-RATING, not internal growth.** Across all
   `er_total ≥ 12%` non-artifact observations the median carry-share was **~11%** — ~89%
   of the expected return on offer is the market closing a valuation gap. A system that
   depends on that is fragile. → **Robustness = anchor the book in QUALITY**, the only
   thing that converts return toward durable earnings growth. The `moat > 7.8` full-engine
   book was the *only* one with genuinely positive realized earnings growth (+85%).

2. **Quality raises both return and robustness.** Full-engine `er_total` + `moat` screen:
   ~22% CAGR, Sharpe 1.24, growth-backed. Dropping quality for raw ER got a higher headline
   (~28%) but a more concentrated, rerating-dependent, fragile book.

3. **Simple diversification beats optimization.** The 6/30 factor slot cap + country cap +
   inverse-vol weighting cut drawdown to −19.5% (Calmar 0.87). A full **CVaR / co-tail
   optimizer did WORSE out-of-sample** (−21%, Calmar 0.80) — estimation error dominates.
   → **Use robust heuristics, never covariance/CVaR optimization.**

4. **Hedge the tail EXTERNALLY, don't degrade selection.** Gate 2 (a name-level downside
   screen) cost ~3 pts of return for almost no drawdown benefit (Calmar 0.75). A
   semi-annual OTM **put on MSCI World IMI** cut drawdown by a third-to-half for ~3.5%
   and lifted Calmar to ~1.03 while keeping every name. → **separate the return engine
   from the risk control.**

5. **Let winners run, cut losers, low turnover.** Tolerance bands (~15%/yr turnover) matched
   or beat full reconstitution (~150%/yr). Sell-on-fail (drop names that fall below the
   gate) helped the broad book. → **bands + blank-slate exit, quarterly.**

## Architecture — Core + Satellite + external hedge

### 1. THE CORE — robust compounder index (~75% of capital)
The durable spine. Its job is growth-backed compounding, not maximum return.
- **Universe:** the researched high-moat names carrying the 11 risk tags.
- **Gate (all required):** full-engine `er_total ≥ 12%` (5-yr price→intrinsic, terminal
  incl., supply/demand-normalized) **AND not artifact** AND **moat ≥ 7.0** (the robustness
  anchor) AND no Extreme (severity-4) risk.
- **Rank:** `er_total` with a mild volatility penalty (`er_total / vol^0.5`) — keeps most of
  the return of raw-ER selection but leans off the most volatile names.
- **Build 25–30 names** under the **6-of-30 risk-tag slot cap** + **≤ 6 per country**
  (correlation control).
- **Weight inverse-volatility**, capped **[1.5%, 6%]** (drawdown control; DeMiguel-style
  robust, not optimized).
- **Maintain:** quarterly tolerance bands — trim a name back at 5%, liquidate at 2% or when
  it drops below the 12% gate (blank-slate exit). ~15–20% turnover/yr.
- *Expected (from the tested components): ~18–22% CAGR, Sharpe ~1.2, max DD ~−20% before
  the hedge; returns materially more earnings-growth-backed than the raw-ER book.*

### 2. THE SATELLITE — concentrated conviction (~25%)
Where the extra return comes from. Funded by selling the Core pro-rata (the IPS mechanism).
- **8–12 names**, each clearing a HIGHER bar: `er_total ≥ 18%`, **history-CONFIRMED** quality
  (moat vs. through-cycle ROIC*), non-artifact, **AND carry > 0** — a hard robustness filter:
  even the concentrated bets must have *positive internal compounding*, so we never bet big
  on a pure rerating/value-trap.
- **Return-interpolated sizing 3–8%** per name (the IPS satellite rule) — size by upside.
- Same slot-cap discipline across Core+Satellite combined (no risk tag > 20% of total).

### 3. THE HEDGE — external tail control (overlay, ~beta-sized)
- A standing **semi-annual ~8–10% OTM put on MSCI World IMI**, notional = book beta (~0.8)
  × NAV. Optionally a **collar** (sell an ~8% OTM call to fund it) to cut the ~3.5% premium
  toward ~1% in exchange for capping some upside.
- This is the drawdown control — far more efficient than screening names (proven this session).

## Why this is BOTH high-returning and robust
- **Return:** the Satellite concentrates in the best quality-confirmed ideas (the raw-ER
  work showed ~28% is achievable at the top end); the Core compounds at ~20%; no Gate-2
  return drag.
- **Robust:** the moat floor makes the return earnings-growth-backed (not a pure rerating
  bet); the carry>0 satellite filter blocks value traps; slot/country/name caps + severity
  screen kill correlated and idiosyncratic wipeouts; inverse-vol Core lowers drawdown; the
  index put/collar caps the systematic tail cheaply; bands keep turnover and cost low; and
  we deliberately avoid covariance/CVaR optimization (it overfits).

## Deliberate NON-choices (things the data said to avoid)
- ❌ Selecting on raw ER without a quality floor (fragile, rerating-dependent).
- ❌ Selecting on carry/internal-growth alone (tilts to cheap cyclicals; underperforms).
- ❌ Gate-2-style downside name-screens (cost return, poor drawdown payoff).
- ❌ CVaR / covariance weight optimization (overfits; loses to inverse-vol).
- ❌ Calendar reconstitution (150%/yr turnover for no benefit).

## Honest caveats
- All backtests are local-currency, price-only (no dividends/FX/costs); the *relative*
  rankings are robust, the absolute CAGRs are optimistic.
- The 2016–26 window had two sharp-but-quick crashes (put-friendly) and a rerating-friendly
  regime; a prolonged bear or a valuation de-rating would compress the rerating component —
  which is exactly why the quality anchor and the hedge exist.
- Tag/severity coverage is partial; the Core universe is the tagged subset.

## BUILT & TESTED — the empirical verdict (backtest_system.py)

The integrated Core+Satellite+hedge was built and run end-to-end. The result
overturned part of my own design: **complexity did not help.**

| System (quarterly, full engine) | CAGR | Sharpe | Max DD | Calmar |
|---|---|---|---|---|
| **RECOMMENDED — Resilient Core-30 (robust spine, whole book)** | 17.7% | 1.03 | −19.5% | **0.91** |
| **+ MSCI World IMI put (drawdown-averse variant)** | 13.6% | 1.01 | **−10.8%** | **1.26** |
| Satellite ONLY (standalone, return-interpolated) | 20.0% | 0.85 | −29.7% | 0.67 |
| Core + Satellite 75/25 (proper blend) | 18.0% | 0.98 | −21.8% | 0.83 |
| High-return alternative — Framework slot-cap | 21.8% | 1.05 | −28.5% | 0.77 |
| Universe EW | 14.4% | 0.82 | −26.2% | 0.55 |

- **CORRECTION on the Satellite** (measured standalone after the fact). It was *not* a
  failure — that verdict was a build artifact (the integrated core sleeve was cannibalized
  of its best names). Standalone the satellite compounds at **20% CAGR** (Sharpe 0.85, DD
  −29.7%, Calmar 0.67) — the highest-returning single sleeve — and the return-interpolated
  sizing added +3.5 pts over equal-weighting the same names. A PROPER 75/25 blend (two books
  held separately) returns **18.0%** vs the core's 17.1% — the satellite is **accretive to
  return** — but it's high-vol and 0.88-correlated to the core, so it costs a little
  risk-adjusted quality (Sharpe 1.00→0.98, Calmar 0.87→0.83). A ~10% tilt is Calmar-neutral.
  → **Include the satellite sized to mandate:** 0% for max Sharpe/Calmar; ~10% for near-free
  extra return; 20–25% for return-seeking (then the MSCI put earns its keep on the drawdown).
- **The winning system is therefore the SIMPLE one:** the **Resilient Core-30** (quality
  floor + risk-adjusted er/vol selection + factor/country caps + severity screen +
  inverse-vol weighting + bands) as the whole book — **17.7% CAGR, Sharpe 1.03,
  Calmar 0.91** — with the **MSCI World IMI put as an OPTIONAL overlay** for the
  drawdown-averse (Calmar **1.26**, max DD **−10.8%**, ~4% premium).
- **If maximum return is the mandate** (and −28% drawdowns are acceptable), the Framework
  slot-cap tilt delivers ~22% — the return/robustness trade-off is explicit, not free.

**Meta-lesson (the whole session in one line):** every added layer — Satellite,
CVaR optimization, Gate 2, integrated hedge — *failed to beat* the simple robust core.
Return and robustness came from **quality + simple diversification + inverse-vol + an
external hedge**, not from machinery. Keep the engine sophisticated; keep the portfolio
construction boring.
