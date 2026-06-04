---
name: aip-value
description: >-
  AIP Value — value a company with a three-stage, ROIIC-driven DCF run against an attached
  Excel screener file. The model is anchored in the law that growth requires
  capital (g = ROIIC x RR) and that competitive advantage fades — returns
  regress to the cost of capital over a Competitive Advantage Period, a Fade
  Period, and a Terminal stage. Use when the user attaches an Excel/.xlsx file
  and asks to value a company, run a DCF / intrinsic value / ROIIC valuation,
  compute a fair value or expected upside, or compare a stock's price to its
  operating value.
---

# Three-Stage ROIIC DCF

A capital-discipline DCF. Every growth rate must be funded: `g = ROIIC × RR`.
As the moat erodes, the return on incremental invested capital fades toward the
cost of capital, so later-stage growth requires more reinvestment and creates
less value. Total operating value is the sum of three discounted stages.

## Workflow

1. **Get the Excel file.** The user attaches an `.xlsx` screener with one
   company per row and headers in the first row. Note its path.
2. **Read `reference.md`** for the full formulas and the fixed parameter
   mapping. Do not re-derive the math by hand — use the calculator.
3. **Run the calculator** (it reads the sheet, finds the company, pulls the
   three drivers, and prints every stage):
   ```bash
   python3 .claude/skills/aip-value/roiic_dcf.py "<file.xlsx>" "<company>"
   # (personal install: python3 ~/.claude/skills/aip-value/roiic_dcf.py ...)
   ```
   - List companies first if unsure of the exact name: append `--list` (or run
     with no company argument).
   - Defaults: `--r 0.12 --n1 15 --n2 10 --gterm 0.025`. Override any of them
     for sensitivities.
4. **Present a valuation memo:** total operating value, the PV split across the
   three stages, the implied equity value / per-share value, and the upside or
   downside vs. the market (EV, price, market cap). Call out the 2–3 inputs the
   result is most sensitive to.

## Fixed inputs and stage logic

| Item | Source / rule |
|------|---------------|
| `NOPAT_0` | Excel column **"New Operating Income"** |
| `ROIIC_1` | Excel column **"ROICm 7"** |
| `RR_1` | Excel column **"RR 7"** |
| `r` (cost of capital / required return) | **12%** |
| `n1` (CAP) | **1/3 of the Moat-Score competitive life** (see below); override `--n1` |
| `n2` (Fade) | **2/3 of the Moat-Score competitive life**; override `--n2` |
| `g1` | `ROIIC_1 × RR_1` |
| `ROIIC_2` | `(ROIIC_1 + r) / 2` |
| `g2` | `(g1 + g_term) / 2` (growth fades to midpoint of g1 and terminal) |
| `RR_2` | `g2 / ROIIC_2` (reinvestment *derived* — framework Step 3 / Damodaran) |
| `ROIIC_term` | `= r` (growth is value-neutral) |
| `g_term` | **2.5%** default — must be `< r`; configurable via `--gterm` |
| `RR_term` | `g_term / r` |

**Note on the terminal stage:** the source spec set `g_term = r = 7%`, which
makes the Gordon denominator `(r − g_term)` zero. Per the user's decision the
skill instead uses a terminal growth **below** `r` (default 2.5%) with
`RR_term = g_term / r`, keeping the terminal value finite and the
`g = ROIIC × RR` law intact. The earlier `RR_term = RR_2 × 1.5` idea is
therefore not used. If the user wants a different terminal growth, pass
`--gterm`.

## Competitive period (n1, n2) from the Moat Score

The **"Moat Score"** column sets the *total* competitive period (the output of
the eight weighted categories: Criticality 0.20, Premium 0.15, Hegemony 0.15,
Ecosystem 0.10, Lifecycle 0.10, Substitution 0.10, Demand 0.10, Allocation
0.10). That total life is then split **1/3 into Stage 1 (CAP)** and **2/3 into
Stage 2 (Fade)**:

| Moat Score | Verdict | Total life | n1 (1/3) | n2 (2/3) |
|------------|---------|-----------|----------|----------|
| `< 6.0` | Pass — vulnerable to disruption | < 10y | ~life/3 | ~2·life/3 |
| `6.0 – 7.5` | Watchlist — check substitution | 10–20y | ~life/3 | ~2·life/3 |
| `> 7.5` | Compounder — permanent moat (Lindy) | 50y | ~17y | ~33y |

Within the bounded bands the total life is interpolated linearly; above 7.5 it
is 50y. Override either leg with `--n1` / `--n2`. Always
use the current **"Moat Score"** column (the default). "Moat Score - 7" is the
score as of 7 years ago — useful only for gauging moat *progression*, not for
valuation. `--col-moat` exists only for files with a differently named column.

## The three stages

- **Stage 1 — CAP.** High `ROIIC_1`, growth `g1` for `n1` years. PV of a growing
  annuity of free cash flow `CF_1 = NOPAT_0 (1+g1)(1−RR_1)`.
- **Stage 2 — Fade.** `ROIIC_2` is the midpoint of `ROIIC_1` and `r`; first cash
  flow grows base NOPAT through all of Stage 1 then one Stage-2 year. PV of the
  growing annuity, discounted back over `n1`.
- **Stage 3 — Terminal.** `ROIIC_term = r` so growth adds no value; Gordon-growth
  perpetuity on `CF_term_Yr1`, discounted back over `n1 + n2`.

`Total Operating Value = PV_1 + PV_2 + PV_TV`.

## Expected return (IRR)

The calculator also annualizes the return over an `n`-year horizon (default 5):
`EV_target` = total operating value; a **cash sweep** de-levers net debt by the
forecast free cash flows (`ND_n = ND_0 − ΣCF_t`, less any dividends/buybacks via
`--payout-total`); `EqV_target = EV_target − ND_n`; then
`Expected Return = (EqV_target / EqV_0)^(1/n) − 1` against current market cap.
It also reports the **unlevered return** `(EV_target / EV_0)^(1/n) − 1` — if the
equity IRR is large but the unlevered return is minimal, the thesis leans on
leverage rather than business improvement. Set the horizon with `--horizon`.

## Output

Deliver a concise memo: the drivers pulled from the sheet, the three-stage PV
breakdown, total operating value, implied equity and per-share value, and the
gap to the current market price — plus the key sensitivities (usually `ROICm 7`,
`RR 7`, `r`, and `g_term`). Flag clearly that this is an analytical framework,
not investment advice.
