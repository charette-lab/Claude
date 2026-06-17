---
name: aip-resource-mine
description: >-
  AIP Resource-Mine — NAV / DCF and resource-option valuation for mining
  developers and explorers, which the earnings-based skills (aip-value,
  aip-value-sectors) cannot value because there are no cash flows yet. Builds a
  mine NPV bottom-up from geology + engineering: grades, recoveries,
  payabilities, smelter TC/RC + concentrate transport, NSR royalties, build
  delay, sustaining capex, tax/depreciation and partial ownership, then runs
  cut-off / metal-price / CAPEX sensitivities and a per-share NAV. Adds an
  implied-probability back-out from the market price for early-stage option
  plays. Use when given a 43-101 / technical report or PEA/PFS for a copper,
  gold or other resource project and asked for an NAV, in-situ value, P/NAV,
  CAPEX/grade/price sensitivity, or the implied odds of exploration success.
---

# AIP Resource-Mine — NAV / DCF for mining developers

A pre-production miner has no earnings, so the ROIIC DCF (aip-value) and the
sector models don't apply. This skill values a deposit **bottom-up from its
geology and engineering**:

```
per-tonne NSR  ->  after-tax mine cash flow  ->  NPV (NAV)  ->  P/NAV & per-share
```

It models the levers that actually move a mine NAV — grade × recovery ×
payability per metal, smelter **TC/RC + concentrate transport**, **NSR
royalties**, the **build delay**, sustaining capex, tax with straight-line
depreciation, and **partial ownership** — and stress-tests cut-off grade, metal
price and CAPEX.

## Workflow

1. **Extract the deposit parameters** from the 43-101 / PEA / PFS into a JSON
   file. Get the template:
   ```bash
   python3 .claude/skills/aip-resource-mine/mine_dcf.py --example > deposit.json
   ```
   Fill in ownership, shares, FX, economics (tpd, discount, tax, build years,
   life cap, sustaining, opex/t, capex, NSR royalty, base cut-off, price deck),
   per-metal recovery / payable / refining, concentrate (Cu grade, moisture,
   TC, transport), and the grade-tonnage table by cut-off (Indicated + Inferred).
2. **Value it** (single case):
   ```bash
   python3 mine_dcf.py --deposit deposit.json [--cu 4.50 --capex 3.5e9 --tpd 100000]
   ```
3. **Sensitivity grid** (cut-off × copper × CAPEX):
   ```bash
   python3 mine_dcf.py --deposit deposit.json --sweep
   ```
4. **Implied probability of success** (early-stage option, 2-state vs market):
   ```bash
   python3 mine_dcf.py --deposit deposit.json --implied-prob --price 19.87 --floor 5
   ```

## What it computes

- **Per-tonne NSR** = Σ metal[grade × recovery × payable × (price − refining)]
  − smelter TC − concentrate transport − NSR royalty.
- **Operating margin/tonne** = NSR − opex/t; **annual op CF** = margin × Mtpa.
- **NPV** = −CAPEX (spread over the build years) + Σ after-tax CF over the mine
  life (mineable tonnes ÷ throughput, capped), discounted at the project rate.
  After-tax CF credits the depreciation tax shield and deducts sustaining capex.
- **Attributable NAV** = NPV × ownership; **per-share** in the report currency
  (NPV × FX ÷ shares).
- **Sweep** tabulates NAV and per-share across each cut-off case, a copper grid,
  and a CAPEX grid (e.g. greenfield vs. a brownfield-synergy CAPEX).

## How to read it

- It's a **pre-production developer**, so the standalone NPV is the *intrinsic*
  NAV — apply a **P/NAV of ~0.3–0.7×** (rising toward 1.0× as it de-risks toward
  construction/production) for a market-value read.
- A **negative NPV at greenfield CAPEX is a real signal**, not an error — many
  large low-grade porphyries are sub-economic standalone and only work via a
  **brownfield-synergy CAPEX**, a **high-grade starter pit**, or much higher
  metal prices. Use the CAPEX sweep to find the breakeven and the synergy case;
  if the standalone NPV is negative, carry the asset at an **in-situ / option
  value** (a few cents $/lb), not the DCF.
- For an **early-stage discovery with no resource**, don't force a DCF — value it
  as the **hard floor (liquid + any defined resource's in-situ value) + an option
  on the undefined deposit**, and use `--implied-prob` to back the market-implied
  odds of success out of the share price (2-state: success payoff vs. floor).
  Remember the size of the find drives the success payoff — sensitise on it.

## Important gaps to source per project (not in a 43-101)
A maiden technical report gives geology/resources but usually **not** the
economic build, so confirm and override:
1. **CAPEX** — pre-production build (shafts/declines, mill, tailings); the single
   biggest NAV driver. Distinguish **greenfield vs. brownfield-synergy**.
2. **Production schedule** — throughput (tpd) and the **ramp-up** (block caves
   ramp over 5–8 yrs; the model uses flat production after the build — conservative
   where ramp is slow, optimistic on early cash if ramp is fast).
3. **Tax/royalty detail** — jurisdiction corporate tax + mining royalty (the model
   takes one effective `tax` rate and the NSR royalty; refine per country, e.g.
   Chile's mining royalty, Argentina's RIGI regime).

Analytical framework output, not investment advice.
