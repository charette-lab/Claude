# Reproduction Value of Assets — adding the Greenwald asset leg to the framework

Synthesized from four Columbia "Value Investing in Technology" cases (Meta, Adobe, Salesforce,
PayPal). Purpose: add **reproduction value** (what a competitor would pay to rebuild the business,
*including the intangibles GAAP leaves off the balance sheet*) as a third valuation leg alongside
the engine's operating value, and derive **franchise value = EPV − reproduction value** as an
independent, accounting-based moat gauge that cross-checks CompanyMoat / CoreMoat.

## The three-layer stack (Greenwald protocol)
1. **Asset Value (AV) = reproduction cost** of the balance sheet + off-book intangibles.
2. **Earnings Power Value (EPV)** = sustainable no-growth NOPAT / WACC.
3. Compare: **EPV > AV ⇒ barriers to entry (a franchise/moat exists)**; EPV ≤ AV ⇒ commodity /
   value-trap. Then **Market cap vs EPV** = what you pay for growth. Confirm the moat with
   **marginal ROIC > WACC**.

Case results (EPV vs AV vs Market, $bn): Meta 396–508 / 238 / 442 · Adobe 89 / 37 / 203 ·
Salesforce 93 / 75 / 188 · PayPal 68 / 37 / 92. Every one: EPV > AV ⇒ real moat.

## Reproduction-value build (exact rules, consistent across all four cases)
| Component | Rule |
|---|---|
| **Tangibles** | Start from **book equity**; no line-item haircuts (book ≈ reproduction proxy). Cash/AR/inventory/PP&E at book. |
| **Goodwill** | Remove acquisition goodwill to avoid double-count with capitalized R&D — but **keep** goodwill for separately-saleable products/brands (Salesforce kept Tableau+Mulesoft; PayPal kept Paidy+Honey; Meta removed only speculative Reality Labs). |
| **Product portfolio = capitalized R&D** | Perpetual inventory, **20%/yr declining balance**: `P = Σ_k (0.8)^k · R&D_{t−k}`, ~11–17 yrs of history. Strip speculative-segment R&D (Meta ex-Reality Labs, ramp 0→25%). |
| **Customer/user base = capitalized S&M** | Flow model `C_t = (1−a)·C_{t−1} + β_t·i_t` (i = S&M spend, a = churn 5–10%, β = customers per $). Value = Σ β_t·ΔC_t (vintaged cost) **or** CAC×base (full replacement). Brand is subsumed here (cross-check vs Interbrand/Kantar). |
| **Workforce** | headcount × avg wage × **20%** (headhunter/replacement cost). Usually small. |
| **AV of equity** | book equity − speculative goodwill + product portfolio + customer base + workforce. |

## EPV build — the piece that fixes "reported earnings are artificially low"
Heavy S&M/R&D depress reported margins (Salesforce reported 2.1% op margin!). Restate: **expense
only the maintenance portion, add back the growth portion** of R&D and S&M.
- **Maintenance R&D = 20% × product-portfolio stock**; growth R&D = total − maintenance → add back.
- **Maintenance S&M = churn × base / β = a·C/β**; growth S&M = total − maintenance → add back.
- Adjusted NOPAT = (op income + growth add-backs) × (1−tax); **EPV = Adj NOPAT / WACC**; then
  + cash − debt = EPV of equity. WACC used: 8–11%.
- Optional cessation overlay: `r* = (p + r)/(1 − p)`, p = probability the business disappears.

## Marginal-ROIC moat confirmation
`marginal ROIC = ΔAdjNOPAT / Σ growth-capex`, where growth-capex = (capex − depreciation) +
(R&D − maint) + (S&M − maint). Meta 14%, Adobe 37%, SF 15%, PP >WACC — all > WACC ⇒ growth
creates value. "**Growth is only worth paying for behind a moat.**"

## How this maps onto OUR framework (what already exists vs the gap)
The `hist` panel ALREADY carries the machinery:
- `Gross Reproduction Cost` — inflation/age-adjusted PP&E reproduction (tangible).
- `Operating Working Capital`, `Base Physical Capital & Leases` — tangible operating base.
- `R&D Capital Base` (built via `R&D Decay Rate`) — **= the CBS product portfolio** (20% decay).
- `SG&A Capital Base` (via `SG&A Cap Rate` / `SG&A Decay Rate`) — a capitalized-SG&A intangible.
- The `overearning` engine's growth/maintenance split ≈ the CBS EPV add-back.
- The Scored sheet has `OpValue` (engine operating value ≈ EPV+growth) and `ImpliedMoat`/`MoatGap`.

**The gaps (why the naive assembly does NOT yet validate — corr(moat, franchise)≈0):**
1. **No customer-base (S&M/CAC) reproduction.** In every CBS case this is the *largest* intangible
   (Meta $44bn, SF $36bn). The panel's `SG&A Capital Base` is a generic SG&A capitalization, not the
   β/churn customer-flow reproduction — and it runs only ~2% of invested capital here, far too small.
2. **No growth/maintenance EPV restatement surfaced** as a standalone EPV (only the engine's internal use).
3. **PP&E reproduction missing for ~89/396 names**, and negative working-capital float drives some
   assembled RVs negative — needs the book-equity-anchored CBS build, not raw component summation.
4. This universe is tangible-heavy (industrials/utilities/consumer), so the tech-style intangible
   reproduction matters for a minority of names — but for those (software, payments, platforms) it is
   decisive and currently mis-valued by a pure tangible/IC proxy.

## Proposed integration (the build)
1. **`reproduction_value` module**: per name, from the `hist` R&D and S&M/SG&A history, build
   (a) capitalized R&D (20% perpetual inventory), (b) customer base via the β/churn CAC model
   (proxy customer count from revenue where not disclosed), (c) workforce, (d) tangible = Gross
   Reproduction Cost (fill missing from Base Physical Capital) + working capital; anchor on book equity.
2. **EPV module**: growth/maintenance add-back → Adj NOPAT / WACC (reuse `overearning`), + cash − debt.
3. **Franchise value = EPV − AV**, `FranchiseShare = Franchise/EPV`, `AV/EV` asset floor.
4. **Cross-checks**: correlate FranchiseShare with CompanyMoat (validation); flag disagreements
   (high moat but low franchise = suspect moat; low moat but high franchise = hidden moat).
5. **Buried-core**: reproduction value of the CORE segment vs the group = asset backing for the
   activist thesis (how much of the rescue is asset-covered).
6. Surface `ReproValue`, `EPV`, `FranchiseValue`, `FranchiseShare`, `AV/EV` as new Scored columns.

## Refinements from the Module quick-reference guides (Meta/PayPal/Booking)
- **Workforce = 20–25% of the wage bill** (headhunter replacement cost).
- **Attrition cancels maintenance in the customer/TPV flow**, so *invested capital to grow the base =
  cumulative S&M spend* (`IC_growth = Σ S&M`); maintenance S&M = churn×base/β offsets attrition exactly.
- **Marginal ROIC strips ORGANIC growth**: credit only ΔNOPAT attributable to growth *capital*
  (subtract organic-growth NOPAT), ÷ accumulated growth capex; compare to WACC. Meta 33%, PayPal 20.5%.
- **Level vs growth-rate shocks** (Gordon `V/E=(1−k)/(r−g)`): a *level* shock moves value proportionally
  (multiple constant); a *growth-rate* shock moves value AND the multiple. Maps onto our carry/rerate split.
- **Crisis lens (Booking)**: when AV can't be built, use EPV alone as a *conservative floor*; buy when
  price dips below conservative EPV (e.g. Booking Q1-2020). "Buy the product/service, not the company."

## KEY FINDING — franchise value is ORTHOGONAL to the moat score (do not use it as a moat validator)
Empirically (352 names): `FranchiseShare = 1 − AV/EPV` ranks almost perfectly with **ROIC
(Spearman 0.94)** and is **uncorrelated with CompanyMoat (0.03)** / CoreMoat (0.10). Algebraically
`FranchiseShare ≈ 1 − WACC/ROIC` — it is a capitalization of the **current excess-return spread**, a
LEVEL, whereas the moat score is DURABILITY. They are different axes. So reproduction/franchise value
adds a **new orthogonal dimension**, it does not confirm the moat. Correct uses:
1. **Asset floor** `AV/EV` — the downside leg, now for the whole book (not just cyclicals).
2. **The franchise × moat map** (4 quadrants): compounder (earns+durable) · fading (earns, not durable,
   mean-reversion risk) · **BURIED CORE (durable, under-earning)** · commodity.
3. **Buried-core / activist screen (the prize)**: `CoreMoat ≥ 7.8` **and** franchise value suppressed
   (group NOPAT/ROIC ≈ 0 because bad subs crush earnings) = a durable franchise not currently earning
   its excess. 98 such names in the book; e.g. REHN.S (CoreMoat 8.7, group ROIC ~0), MVVGn.DE
   (CoreMoat 8.8, MoatGap 2.1). Free the core → NOPAT rises → EPV rises → franchise value materialises.
   This operationalises the activist thesis with reproduction value: buy suppressed franchise behind a
   durable core, get paid as it un-buries.

Status: method fully captured (all 6 modules); `reproduction_value.py` computes AV, EPV, FranchiseValue,
FranchiseShare, AV/EV, the franchise×moat quadrants and the buried-core screen from existing panel
components (IC + capitalized R&D/SG&A). Remaining for a full CBS build: (a) age/inflation-adjusted PP&E
tangible reproduction, (b) the S&M customer-base (β/churn) reproduction, (c) the EPV growth/maintenance
restatement — decisive for asset-light tech names, minor for this tangible-heavy quality universe.
