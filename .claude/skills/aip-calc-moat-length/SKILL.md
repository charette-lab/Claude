---
name: aip-calc-moat-length
description: >-
  Reverse of the aip-value DCF: given a company's current market cap, solve for
  the moat length (competitive-advantage period, n1 + n2 in years) that the price
  implies at YOUR required return. Holds the discount rate at the required return
  and finds the moat life where the model's operating value equals the market
  enterprise value — an expectations-investing / reverse-DCF read. Use when the
  user attaches the Excel screener and asks what moat length / competitive
  advantage period / CAP the market is pricing in, how many years of excess
  returns are baked into the price, or to reverse-engineer the aip-value model.
---

# AIP — Implied Moat Length (reverse aip-value)

The aip-value skill goes **moat length → value**. This goes **value → moat
length**: it asks *"if I require return r, how many years of competitive
advantage must this company sustain for today's price to be fair?"*

It reuses the **exact aip-value engine** (imported from `../aip-value/`), so all
mechanics are identical — two-phase fade (3y hold of ROICm7, then ROIIC and RR
both mean-revert at the Moat-Score persistence φ toward the CFROI base), and the
value-driver terminal. The only change: instead of taking n1+n2 from the Moat
Score, it **solves for n2** (with n1 fixed at 3y) so that:

```
value_company(…, n1=3, n2, φ, base)  ==  Target EV
Target EV = Market Cap + Net debt        (net cash reduces EV)
```

discounting at **r = the required return** (`--r`, default 12%).

## Workflow
1. Get the attached `.xlsx`. Note its path (`--list` to see names).
2. Run:
   ```bash
   python3 .claude/skills/aip-calc-moat-length/moat_length.py "<file.xlsx>" "<company>" --r 0.12
   # (personal install: python3 ~/.claude/skills/aip-calc-moat-length/moat_length.py ...)
   ```
3. Report the **implied moat length (n1+n2)** and compare it to the Moat-Score
   framework length: market-implied **> score** → price bakes in a longer/stronger
   moat than the score supports (rich); **< score** → cheap.

## Equity hurdle → per-company WACC (`--re`)

Like aip-value, pass **`--re` (required equity return)** instead of a flat `--r`
and the discount rate becomes each company's **WACC** (synthetic credit rating
from interest coverage + country risk-free base). The implied moat length is
then "what CAP justifies the price if I require `R_e` on *equity*". **Refresh the
risk-free bases live** and pass them, e.g.
`--re 0.07 --country-base "EUR=0.0303,USD=0.0405,JPY=0.0267,KRW=0.0412"`; cached
fallbacks apply offline. Banks/financials: ignore (coverage meaningless).

## Inputs (same as aip-value)
`NOPAT_0` ← New Operating Income · `ROIIC_0` ← ROICm 7 · `RR_0` ← RR 7 ·
`base` ← sector CFROI median (GICS) · `φ` ← Moat-Score tier · `g_term` 2.5% ·
`n1` = 3y hold (override `--n1`). Required return via `--r`.

## Reading the output, including the edge cases
- **Interior solution** (most informative): e.g. "implied moat = 6.3y vs score 50y → cheap".
- **"price not justified at any length"** appears when the **required return exceeds
  the sector CFROI base** (e.g. r = 12% vs an 8.5% base): the firm can never
  out-earn the hurdle in the long run, so value asymptotes *below* the price for
  every moat length. Economic read: the stock is priced for a **lower discount
  rate** than you require. Re-run with a lower `--r` (nearer the base rates) to get
  an interior implied length.
- **"≤ 3y"** means hold-only value already exceeds the price — the market prices in
  essentially no competitive advantage.

Flag clearly that this is an analytical framework, not investment advice.
