# AIP Value — ROIIC Persistence-Fade DCF (methodology)

Anchored in two empirical realities: **growth requires capital** (`g = ROIIC ×
RR`), and **excess returns mean-revert** — `ROIIC − WACC` decays toward zero at a
speed set by the durability of the moat. This avoids extrapolating a
cyclically-inflated return (the classic trap with a semiconductor or commodity
firm posting a 40–60% ROIIC at a demand peak).

## Variable glossary

| Notation | Definition |
|----------|------------|
| `NOPAT` | Net Operating Profit After Tax |
| `r` | Cost of capital / WACC / required return (discount rate) |
| `ROIIC` | Return on Incremental Invested Capital |
| `RR` | Reinvestment Rate (proportion of NOPAT reinvested) |
| `g` | Growth rate of operating income, `g = ROIIC × RR` |
| `φ` | Persistence factor — the annual decay of the excess return |
| `N` | CAP — the explicit fade period, in years |

## Step 1 — Map the Moat Score to CAP and persistence

The Moat Score (the weighted eight-category Final Score) sets the **empirical CAP
duration** and the **persistence factor**:

```
Final Score = Criticality x0.20 + Premium x0.15 + Hegemony x0.15
            + Ecosystem x0.10 + Lifecycle x0.10 + Substitution x0.10
            + Demand x0.10 + Allocation x0.10
```

| Moat Score | Tier | CAP (N) | Persistence φ | Reality |
|------------|------|---------|---------------|---------|
| `> 7.5` | Superior / Wide | 10–20y | 0.85–0.95 | Network effects, switching costs, regulatory monopoly slow — but don't stop — decay. Base rate overtakes the inside view ~year 8–9. |
| `6.0–7.5` | Narrow / Standard | 5–10y | 0.70–0.80 | Brand, some pricing power, minor scale; top-quartile returns regress within a decade. |
| `< 6.0` | No moat / cyclical | 1–5y | 0.50–0.60 | Commodity/hardware at a cyclical shortage; capacity is added and the anomaly is arbitraged away fast. |

Interpolate linearly within each band by score. **Why even superior moats fade:**
the *size penalty* (reinvesting ever-larger incremental capital at 40–50% becomes
impossible as the denominator scales toward the TAM) and *creative destruction*
(moats are bypassed by asymmetric innovation — e.g. on-prem hardware → cloud).

## Step 2 — Decay the excess return over the CAP

For each year `t = 1 … N`, ROIIC mean-reverts to the **industry base rate** (the
long-run level from the Base Rate Book), not all the way to WACC:

```
ROIIC_t = base + (ROIIC_0 − base) · φ^t   # = ROIIC_0·φ^t + (1−φ^t)·base
g_t     = ROIIC_t · RR                    # RR held at the current rate (RR 7)
NOPAT_t = NOPAT_(t−1) · (1 + g_t)
FCF_t   = NOPAT_t · (1 − RR)
```

`ROIIC_0` is the starting marginal return (ROICm 7), which may be cyclically
elevated; the geometric decay reverts it to the industry `base` rate. Holding
`RR` constant means growth falls *with* returns — the desired behaviour at a
cyclical peak. The `base` rate is looked up by GICS industry (override
`--base-rate`).

## Step 3 — Discount the explicit period

```
PV_explicit = Σ_{t=1..N} FCF_t / (1 + r)^t
```

## Step 4 — Terminal value

After `N` years ROIIC has settled at the industry **base rate**. The terminal is
the value-driver perpetuity with a safe terminal growth
`g_eff = min(g_term, ~base, ~r)` (keeps it `< r` and reinvestment `≤ 100%`):

```
RR_term = g_eff / base
TV      = NOPAT_N · (1 + g_eff) · (1 − RR_term) / (r − g_eff)
PV_TV   = TV / (1 + r)^N
```

If `base = r` this reduces to the value-neutral `NOPAT/r` form; `base > r` is
perpetual value creation, `base < r` mild value destruction. (`base ≤ 0` falls
back to `TV = NOPAT_N / r`.)

## Step 5 — Total operating value

```
Total Operating Value = PV_explicit + PV_TV
```

To equity: subtract net debt (add net cash); divide by diluted shares for a
per-share value; compare to market EV / price.

## Step 6 — Expected return (IRR over the holding horizon)

```
EV_target  = Total Operating Value                 # realised at exit
ND_n       = ND_0 − ( Σ_{t=1..n} FCF_t − dividends/buybacks )   # cash sweep
EqV_target = EV_target − ND_n
Expected Return = (EqV_target / EqV_0)^(1/n) − 1   # vs current market cap
Unlevered Return = (EV_target / EV_0)^(1/n) − 1    # sanity check
```

If the equity return is large but the unlevered return is minimal, the thesis
rests on leverage rather than business improvement.

## Fixed mapping for the attached screener

| Model input | Source |
|-------------|--------|
| `NOPAT_0` | "New Operating Income" |
| `ROIIC_0` | "ROICm 7" |
| `RR` | "RR 7" (current rate, held constant) |
| `r` | 12% |
| `base` (ROIIC base rate) | industry table on "GICS Industry Group Name" (override `--base-rate`) — **placeholders pending the Base Rate Book** |
| `CAP`, `φ` | from "Moat Score" (table above; override `--cap` / `--persistence`) |
| `g_term` | 2.5% (`< r`) |

### Notes on resolved design choices
- **Persistence vs. discount rate naming.** The source framework calls the
  persistence factor `r`; here `r` is the discount rate, so persistence is `φ`.
- **RR held constant.** Reinvestment uses the current rate (RR 7) throughout; as
  ROIIC fades, growth fades with it. (Earlier 3-stage variants instead faded
  growth and derived RR; the persistence model supersedes them.)
- **Terminal growth.** `g_term < r` keeps the Gordon term finite; with
  `ROIIC_term = r` the terminal value is the value-neutral `NOPAT/r` form.

> Analytical framework output, not investment advice.
