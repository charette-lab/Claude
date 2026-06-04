---
name: aip-value
description: >-
  AIP Value — value a company with a ROIIC persistence-fade DCF run against an
  attached Excel screener file. Growth must be funded (g = ROIIC x RR), and a
  company's excess return (ROIIC − WACC) decays geometrically toward the cost of
  capital at a speed set by moat durability — so a cyclically-inflated starting
  ROIIC is mean-reverted rather than held flat. Use when the user attaches an
  Excel/.xlsx file and asks to value a company, run a DCF / intrinsic value /
  ROIIC valuation, compute a fair value, expected return/IRR, or upside, or
  compare a stock's price to its operating value.
---

# AIP Value — ROIIC Persistence-Fade DCF

A capital-discipline DCF. Every growth rate must be funded: `g = ROIIC × RR`.
The key idea: **excess returns don't persist** — `ROIIC − WACC` decays
geometrically each year by a **persistence factor** set by moat quality, so a
high starting ROIIC (e.g. a semiconductor firm at a cyclical demand peak) reverts
toward the cost of capital instead of being extrapolated. Operating value =
PV(explicit fade-period FCF) + PV(terminal).

## Workflow

1. **Get the Excel file.** The user attaches an `.xlsx` screener with one
   company per row and headers in the first row. Note its path.
2. **Read `reference.md`** for the full formulas and the parameter mapping. Use
   the calculator — don't re-derive the math by hand.
3. **Run the calculator** (reads the sheet, finds the company, pulls the drivers,
   prints the fade schedule and valuation):
   ```bash
   python3 .claude/skills/aip-value/roiic_dcf.py "<file.xlsx>" "<company>"
   # (personal install: python3 ~/.claude/skills/aip-value/roiic_dcf.py ...)
   ```
   - List companies first if unsure of the exact name: append `--list`.
   - Defaults: `--r 0.12 --gterm 0.025 --horizon 5`; CAP and persistence come
     from the Moat Score. Override with `--cap`, `--persistence`, etc.
4. **Present a valuation memo:** total operating value, explicit-vs-terminal PV
   split, implied equity / per-share value, expected return (IRR), and the gap to
   market. Call out the 2–3 inputs the result is most sensitive to.

## Inputs and engine

| Item | Source / rule |
|------|---------------|
| `NOPAT_0` | Excel column **"New Operating Income"** |
| `ROIIC_0` (starting) | Excel column **"ROICm 7"** |
| `RR_0` (starting) | Excel column **"RR 7"**; fades in Phase 2 to `RR_target = g_term/base` |
| `r` (WACC / required return) | **12%** (`--r`) |
| `base` (CFROI) | real sector CFROI median (GICS column), **no inflation** — mirrors our growth-capitalized, lower ROICm; override `--base-rate` / `--base-inflation` |
| `n1` (hold) / `n2` (fade) | n1 = 3y fixed; n2 = Moat-Score life − 3; override `--n1` / `--n2` |
| `persistence` (φ) | from the **Moat Score** tier; override `--persistence` |
| `g_term` | **2.5%** default — must be `< r` (`--gterm`) |

**Two-phase engine.** Phase 1 holds ROICm7 and RR_0; Phase 2 mean-reverts **both
ROIIC and RR** to their sustainable levels (a rational firm reinvests less as
returns normalise):
```
Phase 1, t = 1…n1 :  ROIIC_t = ROIIC_0 ;  RR_t = RR_0
Phase 2, k = 1…n2 :  ROIIC = base + (ROIIC_0 − base)·φ^k        # fade to CFROI base
                     RR    = RR_target + (RR_0 − RR_target)·φ^k # RR_target = g_term/base
g_t  = ROIIC_t · RR_t
FCF_t = NOPAT_t · (1 − RR_t) ;  PV_explicit = Σ FCF_t / (1+r)^t
```
Optional `--gcap` hard-caps annual growth (trims RR) as a `g ≤ r` backstop;
off by default. ROIIC and RR both fading guarantees the explicit period connects
smoothly to the terminal (`g → g_term`).

`n1 = 3y` (fixed hold) and `n2 = moat-life − 3`, so `n1 + n2` = the moat-score
competitive life (`< 6 → <10y`, `6–7.5 → 10–20y`, `> 7.5 → 50y`). `φ` from the
moat tier (`--persistence`).
**Terminal** (ROIIC has settled at the base rate; value-driver perpetuity with a
safe terminal growth g_eff = min(g_term, ~base, ~r)):
```
TV    = NOPAT_N · (1 + g_eff) · (1 − g_eff/base) / (r − g_eff)
PV_TV = TV / (1+r)^N
Total Operating Value = PV_explicit + PV_TV
```

## ROIIC base rate (industry)

The return reverts toward an industry **ROIIC base rate** (from the Base Rate
Book), looked up from the sheet's "GICS Industry Group Name" column via the
`INDUSTRY_BASE_RATE` table in `roiic_dcf.py`. Override per run with
`--base-rate`. **The table values are placeholders until populated from the Base
Rate Book.** A base rate above WACC implies perpetual value creation in the
terminal; below WACC, mild value destruction.

## CAP & persistence from the Moat Score (empirical CAP durations)

The **"Moat Score"** column (the weighted eight-category Final Score: Criticality
0.20, Premium 0.15, Hegemony 0.15, Ecosystem 0.10, Lifecycle 0.10, Substitution
0.10, Demand 0.10, Allocation 0.10) sets both how long and how slowly the excess
return fades:

| Moat Score | Moat tier | CAP (N) | Persistence φ |
|------------|-----------|---------|---------------|
| `> 7.5` | Superior / Wide (network effects, switching costs, regulatory monopoly) | 10–20y | 0.85–0.95 |
| `6.0 – 7.5` | Narrow / Standard (brand, some pricing power, minor scale) | 5–10y | 0.70–0.80 |
| `< 6.0` | No moat / cyclical peak (commodity, hardware shortage) | 1–5y | 0.50–0.60 |

Values are interpolated within each band by score. A higher φ means the excess
return persists longer (with φ≈0.9 the "inside view" of high returns is roughly
half-eroded by year 7, matching the empirical CAP literature). Even superior
moats must fade because of the **size penalty** (reinvesting ever-larger capital
at 40–50% becomes impossible as the firm scales to its TAM) and **creative
destruction** (asymmetric innovation, e.g. on-prem → cloud). Always use the
current **"Moat Score"**; "Moat Score - 7" is the score 7 years ago (progression
only). No score → defaults to Narrow (CAP 8y, φ 0.75).

> Notation: the source framework writes the persistence factor as `r`; here `r`
> is the discount rate, so the persistence factor is `φ` (`--persistence`).

## Expected return (IRR)

The calculator annualizes the return over an `n`-year horizon (default 5):
`EV_target` = total operating value; a **cash sweep** de-levers net debt by the
forecast FCF (`ND_n = ND_0 − ΣFCF_t`, less dividends/buybacks via
`--payout-total`); `EqV_target = EV_target − ND_n`; then
`Expected Return = (EqV_target / EqV_0)^(1/n) − 1` vs. current market cap. It also
reports the **unlevered return** `(EV_target / EV_0)^(1/n) − 1`.

## Output

Deliver a concise memo: drivers pulled from the sheet, the moat tier / CAP / φ,
the fade schedule, total operating value, implied equity & per-share value, the
expected return, and the gap to market price — plus the key sensitivities
(usually `ROICm 7`, `RR 7`, the moat tier, `r`). Flag clearly that this is an
analytical framework, not investment advice.
