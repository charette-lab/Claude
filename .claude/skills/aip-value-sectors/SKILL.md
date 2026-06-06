---
name: aip-value-sectors
description: >-
  AIP Value Sectors — industry-aware valuation on the rich moat-7 panel dataset.
  Routes every company to the model that matches its value driver: operating
  compounders use the aip-value ROIIC DCF; Energy/Materials use a cycle-normalised
  DCF with a reproduction-cost asset floor; Utilities/Telecom use a regulated
  asset-base DCF (excess return faded to WACC, value ~ RAB); Real Estate uses a
  net-asset-value model. Financials (Banks, Insurance, Financial Services) are
  flagged and skipped — they need a residual-income/ROE model. Use when the user
  attaches the wide 132-column panel file (margins, debt/lease/pension components,
  intangible capitalization, reproduction cost, ROICm windows) and asks to value
  the whole universe with the right model per industry, or value cyclicals,
  utilities, or REITs specifically.
---

# AIP Value Sectors — the right model for each industry

Operating-company DCF is wrong for a utility, a miner, or a property developer —
their value drivers differ. This skill classifies each company by GICS group and
routes it to the matching model, **reusing the aip-value engine** (`roiic_dcf.py`)
for the DCF primitives (WACC, lever-glide, value_company, equity IRR).

## Workflow

1. **Get the wide panel file** — the 2011–2026, one-row-per-company-year dataset
   with the full financial detail (not the slim screener). The latest row per
   `Instrument` is used; **net debt = EV − Market Cap**.
2. **Refresh risk-free rates** (WACC depends on them) and run:
   ```bash
   python3 .claude/skills/aip-value-sectors/sector_value.py "<file.xlsx>" --re 0.09 \
       --country-base "USD=0.0406,JPY=0.0265,EUR=0.0302,GBP=0.0485,SEK=0.0279,KRW=0.0412,CHF=0.0040,CAD=0.0344" \
       --csv sectors.csv
   # one archetype only:  --only cyclical | regulated | nav | operating
   ```
3. **Present** the per-archetype tables (Model/EV, expected return, model note),
   the archetype counts, and the flagged misclassifications.

Requires the **aip-value** skill installed alongside.

## The models

| Archetype | GICS groups | Model | Key mechanics |
|---|---|---|---|
| **operating** | Capital Goods, Software, Semis, Pharma, Health Care, Transportation, Consumer, Media, Comm/Prof Svcs | aip-value ROIIC DCF on **spot** ROICm7 | intangibles already capitalized in `New Operating Income`; g = ROIIC×RR, fade to CFROI base, lever-glide |
| **cyclical** | Energy, Materials | **cycle-normalised** ROIIC DCF + **reproduction-cost floor** | NOPAT = 7-yr-sum/7 (mid-cycle); ROIC = `ROICm_total - 21 years`; value floored at `Gross Reproduction Cost × clamp(ROIC/WACC, 0.3, 1)` (asset value, discounted if sub-WACC) |
| **regulated** | Utilities, Telecommunications | **excess return faded to WACC** → value ≈ RAB | starting ROIC capped at 1.5×WACC, fade target = WACC, utility leverage; reports `model/RAB` (RAB = Invested Capital) |
| **nav** | Real Estate Mgmt & Development | **net asset value** | NAV equity = book equity + max(0, reproduction cost − net PPE); expected return = price→NAV convergence + dividend yield |
| **financial** | Banks, Insurance, Financial Services | **skipped** | needs a residual-income/ROE model (not built) |

## Important behaviours & caveats

- **No Moat Score column** — this file is pre-filtered to moat≥7, so a single tier
  is assumed (`--moat-assume`, default 7.5 → ~20-yr competitive life). Override per
  run; it sets the hold/fade lengths and persistence for the DCF models.
- **Operating-model outliers are input-quality flags, not buys.** The panel's
  *marginal* ROICm and Reinvestment Rate are noisier than the slim screener's
  `ROICm 7`/`RR 7`; a handful of operating names print absurd Model/EV (the
  high-ROIC extrapolation artifact). Treat anything above ~+50% ER as "check the
  inputs," consistent with aip-value.
- **Three GICS groups are heterogeneous and need name-level sub-classification:**
  - *Energy* mixes commodity producers (true cyclical) with asset-light
    services/tech (e.g. LNG containment licensing) that behave like operating
    compounders — the latter show `no repro` and very high returns.
  - *Real Estate* mixes property developers (true NAV) with capital-light real-
    estate **data/SaaS** (CoStar, Altus) — the NAV model **auto-flags these
    `[asset-light: likely OPERATING]`**; revalue them as operating.
  - *Financial Services* mixes exchanges/payments/data (operating) with lenders
    (financial) — the whole group is skipped here; the operating slice should be
    run through aip-value separately.
- **Regulated value ≈ RAB by construction** — `model/RAB` near 1× is the expected
  outcome; developed-market utilities trading well above RAB screen expensive.
- **Reproduction-cost floor** uses the inflation-adjusted asset replacement value
  (`Gross Reproduction Cost`); it only floors when present and is discounted by
  relative profitability so sub-WACC asset bases aren't over-credited.

Builds the **non-financial** models only. Analytical framework output, not
investment advice.
