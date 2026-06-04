# Three-Stage ROIIC Valuation — Full Methodology

Anchored in one reality: **growth requires capital**, and as competitive
advantage fades, returns on capital regress toward the cost of capital.

## Variable glossary

| Notation | Definition |
|----------|------------|
| `NOPAT` | Net Operating Profit After Tax |
| `CF` | Free Cash Flow |
| `r` | Cost of Capital (discount rate / WACC) |
| `ROIIC` | Return on Incremental Invested Capital |
| `RR` | Reinvestment Rate (proportion of NOPAT reinvested) |
| `g` | Growth rate of operating income |
| `n` | Duration (years) of a stage |

## Step 1 — The core growth engine (immutable law)

`g`, `ROIIC`, and `RR` are permanently linked; you cannot set them independently:

```
g = ROIIC × RR        =>        RR = g / ROIIC
```

When transitioning to a stage with different growth or returns, re-solve for the
capital requirement with this identity.

## Step 2 — Stage 1: Competitive Advantage Period (CAP), n1 years

```
RR_1 = g1 / ROIIC_1
CF_1 = NOPAT_0 × (1 + g1) × (1 − RR_1)
PV_1 = CF_1 / (r − g1) × [ 1 − ((1+g1)/(1+r))^n1 ]
```

(If `g1 = r`, the growing-annuity PV is the limit `n1 × CF_1 / (1+r)`.)

## Step 3 — Stage 2: Fade Period, n2 years

```
ROIIC_2 = (ROIIC_1 + r) / 2                 # midpoint toward cost of capital
RR_2    = g2 / ROIIC_2
CF_Stage2_Yr1 = [ NOPAT_0 × (1+g1)^n1 × (1+g2) ] × (1 − RR_2)
PV_2 = ( CF_Stage2_Yr1 / (r − g2) × [ 1 − ((1+g2)/(1+r))^n2 ] ) × 1/(1+r)^n1
```

## Step 4 — Stage 3: Terminal value

Economic theory: in perpetuity the advantage is competed away, so
`ROIIC_term = r` and incremental growth creates no excess value (it still sets
the maintenance capital required).

```
RR_term      = g_term / r
CF_term_Yr1  = [ NOPAT_0 × (1+g1)^n1 × (1+g2)^n2 × (1+g_term) ] × (1 − RR_term)
TV_n         = CF_term_Yr1 / (r − g_term)         # Gordon growth, at end of Stage 2
PV_TV        = TV_n / (1+r)^(n1+n2)
```

`g_term` must be strictly below `r`, or the Gordon denominator is zero.

## Step 5 — Total operating value

```
Total Operating Value = PV_1 + PV_2 + PV_TV
```

To reach equity value, subtract net debt (add net cash); divide by diluted
shares for per-share intrinsic value; compare to market EV / price.

## Step 6 — Expected return (IRR over the holding horizon)

Translate the static valuation into an annualized expected return over an
`n`-year horizon (default `n = 5`).

1. **Target Enterprise Value** — the modelled operating value is assumed to be
   realised as the future EV at exit:
   ```
   EV_target = Total Operating Value
   ```
2. **Project Year-n net debt (the cash sweep)** — over the horizon, the
   forecast free cash flows accumulate as cash or pay down debt, reducing net
   debt dollar-for-dollar. Subtract any dividends/buybacks, which leave the firm
   instead of de-levering:
   ```
   ND_n = ND_0 − ( Σ_{t=1..n} CF_t − Dividends/Buybacks )
   ```
   `CF_t` is the stage-appropriate free cash flow (Stage 1 within the CAP).
3. **Target equity value:**
   ```
   EqV_target = EV_target − ND_n
   ```
4. **Expected equity return (IRR):** with current market cap `EqV_0`,
   ```
   Expected Return = (EqV_target / EqV_0)^(1/n) − 1
   ```

**Optional — unlevered return (sanity check):**
```
Unlevered Return = (EV_target / EV_0)^(1/n) − 1
```
If the equity return is large but the unlevered return is minimal, the thesis
rests on financial engineering / leverage rather than business improvement.

The calculator pulls `ND_0` from "Net debt", `EqV_0` from "Market Cap", and
`EV_0` from "EV"; set the horizon with `--horizon` and any distributions with
`--payout-total`.

---

## Fixed mapping for the attached screener

| Model input | Source |
|-------------|--------|
| `NOPAT_0` | column "New Operating Income" |
| `ROIIC_1` | column "ROICm 7" |
| `RR_1` | column "RR 7" |
| `r` | 12% (required return) |
| `n1` | 1/3 of the "Moat Score" competitive life (override `--n1`) |
| `n2` | 2/3 of the "Moat Score" competitive life (override `--n2`) |
| `g1` | `ROIIC_1 × RR_1` |
| `ROIIC_2` | `(ROIIC_1 + r)/2` |
| `g2` | `(g1 + g_term) / 2` (growth fades to midpoint of g1 and terminal) |
| `RR_2` | `g2 / ROIIC_2` (reinvestment *derived*, per Step 3) |
| `ROIIC_term` | `r` |
| `g_term` | 2.5% default (`< r`; set with `--gterm`) |
| `RR_term` | `g_term / r` |

### Competitive period (n1) — Moat Life Estimation by Score

The Moat Score (the weighted eight-category Final Score) sets the *total*
competitive period, which is split 1/3 into Stage 1 (CAP, `n1`) and 2/3 into
Stage 2 (Fade, `n2`). Bands for total life: `< 6.0` → short-term (< 10y);
`6.0–7.5` → medium (10–20y); `> 7.5` → permanent (50y, "50+", per the Lindy
Effect on basic-need demand). Within the two bounded bands the value is linearly
interpolated; above 7.5 it is fixed at 50. The eight categories and weights:

```
Final Score = Criticality x0.20 + Premium x0.15 + Hegemony x0.15
            + Ecosystem x0.10 + Lifecycle x0.10 + Substitution x0.10
            + Demand x0.10 + Allocation x0.10
```

### Resolved ambiguity (terminal stage)

The original spec said `g_term = required return = r = 7%` **and**
`RR_term = RR_2 × 1.5`. Together these are inconsistent: `g_term = r` makes the
Gordon denominator `(r − g_term) = 0` (undefined), and an independently chosen
`RR_term` breaks the `g = ROIIC × RR` identity. Per the user's decision, the
skill uses a **terminal growth below r** (default 2.5%) with the framework's
`RR_term = g_term / r`. The `RR_2 × 1.5` terminal override is not used.

> Analytical framework output, not investment advice.
