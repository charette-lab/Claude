---
name: aip-core30
description: >-
  AIP Core-30 — build the "Constrained Quality Compounder Index" from an
  aip-value Excel screener. Values every company with the aip-value engine
  (per-company WACC, lever-glide, distributed-cash equity IRR), applies the risk
  framework (market-cap floor, moat floor, TAM-exhaustion screen), ranks by an
  equity-return hurdle, and selects N equal-weighted names under the 20% risk-tag
  rule (no single risk tag held by more than ~6 of 30 stocks). Use when the user
  attaches the moat/ROIC screener and asks to build a portfolio, the core index,
  a 30-stock book, rank-and-select names under the risk framework / tag limits,
  or apply the Investment Policy / Constrained Quality Compounder framework.
---

# AIP Core-30 — the Constrained Quality Compounder Index

Turns an aip-value screener into a finished, risk-constrained portfolio. It reuses
the **aip-value** engine for every valuation (so the expected returns are exactly
what `aip-value` / `rank.py` produce), then layers the Investment Policy framework
on top: screen → tag → construct.

## Workflow

1. **Get the Excel file** — the same one-company-per-row screener `aip-value` uses,
   which must additionally carry the **11 risk-tag columns** (scored 1–4) and an
   `RR 3` column for the TAM screen.
2. **Refresh the risk-free rates** (the WACC depends on them) — web-search the
   current 10y govt yields for the countries in the file and pass them in
   (`USD` = UST − ~40bp swap):
   ```bash
   python3 .claude/skills/aip-core30/core30.py "<file.xlsx>" \
       --country-base "USD=0.0406,JPY=0.0265,EUR=0.0302,GBP=0.0485,SEK=0.0279,KRW=0.0412,CHF=0.0040,CAD=0.0344" \
       --csv core30.csv
   # (personal install: python3 ~/.claude/skills/aip-core30/core30.py ...)
   ```
3. **Present the portfolio memo:** the ranked 30, equal weight, ER at both
   hurdles, the **tag-exposure table** (proving the 20% rule holds), the names
   *bumped* by the tag caps, country mix, and the exclusion counts.

Requires the **aip-value** skill installed alongside (it imports `roiic_dcf.py`
from `../aip-value/` or `~/.claude/skills/aip-value/`).

## How it maps to the framework

| Framework phase | What the script does |
|---|---|
| **P1 Universe** | Excludes financials; `--mcap-floor` (default $500M); `--moat-min` (default 7.0 on `Moat Score`). |
| **P1 TAM exhaustion** | Runs the **reinvestment-deceleration** leg only: excludes if `RR 3 < (1 − --tam-reinv-drop)·RR 7` (default 25% drop). The **give-up ratio** and **top-line-gravity** legs need FCF/payout/revenue-history columns the sheet lacks — *not applied* (flagged). |
| **P2 11 risk tags** | Read from the sheet's tag columns. They're scored **1–4, not binary** (1 = safe, higher = more exposure); **material exposure = score ≥ `--tag-material` (default 4)**. |
| **P3 Construction** | Rank by ER at `--rank-re` (default **12%**), greedily take `--n` (30) names s.t. no tag is held by more than `--tag-cap` (6) stocks — the **20% rule** (6 × 3.33% ≈ 20%). Equal-weighted at `1/n`. |

## Inputs and valuation

Every company is valued by the **aip-value** engine at two equity hurdles:
`--rank-re` (rank by this) and `--alt-re` (shown alongside). The expected return is
the **distributed-cash equity IRR** with the **lever-glide** de-risking on by
default (`--no-lever-glide` to turn it off). Drivers, WACC, country rates, and the
CAGR-fallback flag (`c`) are all inherited from `aip-value` — see that skill's
`reference.md`.

## Key options

| Flag | Default | Meaning |
|------|---------|---------|
| `--rank-re` / `--alt-re` | 0.12 / 0.07 | equity hurdle to rank by / to also display |
| `--mcap-floor` | 5e8 | minimum market cap (data-thin guard) |
| `--moat-min` | 7.0 | minimum Moat Score |
| `--n` | 30 | portfolio size |
| `--tag-cap` | 6 | max stocks sharing a material tag (the 20% rule) |
| `--tag-material` | 4 | tag score ≥ this = material exposure |
| `--er-gate` | none | also require ER at the ranking hurdle `>` this (e.g. `0` = positive at 12%) |
| `--country-base` / `--country-crp` | cached | refresh risk-free bases / country risk premiums |
| `--csv` | — | write the portfolio to a CSV |

## Important behaviours & caveats

- **A hard `--er-gate 0` at the 12% hurdle cannot fill 30.** Only ~19 names in a
  moat-7 universe are positive at 12% (≈17 after the tag caps; ≈25 with no size
  floor). The framework's "exactly 30 equal-weight" and a hard positive-at-12%
  gate are mutually exclusive — use the gate for a concentrated book, or rank
  without it (default) for a full 30.
- **`--tag-material 3` is usually infeasible for 30** — this quality universe is
  so correlated on macro/regulatory/geographic/FX that requiring ≤6 with a
  "moderate-or-higher" exposure caps the book in the low-20s. Default `4` (the
  top "high-exposure" bucket) is what makes a 30-stock book buildable.
- The **20% rule is enforced, not country diversification** — the book can tilt
  heavily to one country (it caps the 11 risk tags, which only partly proxy
  geography via the FX / Geographic tags).
- Ranking *without* a gate means the tail of the list can be mildly negative at
  the ranking hurdle (they're the least-expensive of the remaining quality names).
- This builds the **Core index** only — the framework's Satellite book (12% IRR
  gate + ≤20% downside-probability gate + return-based sizing) needs a
  downside-floor model that is not part of this skill.

Flag clearly that this is an analytical framework, not investment advice.
