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

**Whole-sheet ranking in one command** — use `rank.py` to value every company,
ranked by expected return, with the per-company WACC, rating, and implied moat:
```bash
python3 .claude/skills/aip-value/rank.py "<file.xlsx>" --re 0.07 [--re2 0.12] \
        --country-base "EUR=0.0303,USD=0.0405,JPY=0.0267,KRW=0.0412,SEK=0.0273"
```
`--re2` adds a second equity-hurdle ER column; banks are marked `FIN*`. Add
`--lever-glide` to apply the de-risking WACC glide (below) across the whole sheet.

## Inputs and engine

| Item | Source / rule |
|------|---------------|
| `NOPAT_0` | Excel column **"New Operating Income"** |
| `ROIIC_0` (starting) | Excel column **"ROICm 7"** |
| `RR_0` (starting) | Excel column **"RR 7"**; held flat as the structural reinvestment rate, but *lifted* when the sales-growth floor binds |
| `r` (WACC / required return) | **12%** (`--r`) |
| `base` (CFROI) | real sector CFROI median (GICS column), **no inflation** — mirrors our growth-capitalized, lower ROICm; override `--base-rate` / `--base-inflation` |
| `n1` (hold) / `n2` (fade) | n1 = 3y fixed; n2 = Moat-Score life − 3; override `--n1` / `--n2` |
| `persistence` (φ) | from the **Moat Score** tier; override `--persistence` |
| `g_term` | **2.5%** default — must be `< r` (`--gterm`) |

**Reinvestment-driven engine.** Growth is whatever the firm's *own* reinvestment
can fund — `g = ROIIC × RR` — not an externally-imposed industry rate. `RR_0` is
the structural reinvestment rate, **held flat**; as ROIIC fades the growth falls
of its own accord. A **sales-growth floor** (the Mauboussin size×industry
**median** base rate, recomputed as sales compound) stops `g` — and so the
implied RR — going artificially low. When the floor binds, RR is *lifted* to fund
it:
```
Phase 1, t = 1…n1 :  ROIIC_t = ROIIC_0
Phase 2, k = 1…n2 :  ROIIC_t = base + (ROIIC_0 − base)·φ^k        # fade to CFROI base
g_t  = max( ROIIC_t · RR_0 , sales_base_median(size_t) )          # funded, but ≥ floor
RR_t = min( g_t / ROIIC_t , 1 )                                   # lifted when floored
FCF_t = NOPAT_t · (1 − RR_t) ;  PV_explicit = Σ FCF_t / (1+r)^t
```
Growth is **not** capped in the moat period — a moat legitimately compounds above
`r`, and the explicit period is finite so it can't be infinite. Only the
**terminal** enforces `g_eff < r` (that is what prevents infinite value in
perpetuity). **Note:** within the explicit period, where the faded ROIIC sits
*below* WACC the floor forces *value-destroying* reinvestment (RR rises to fund
growth earning < cost of capital). That is intentional and conservative — it
stops an artificially-low RR inflating FCF. The floor is **not** applied in the
terminal, where it would force that value destruction forever (see below).

`n1 = 3y` (fixed hold) and `n2 = moat-life − 3`, so `n1 + n2` = the moat-score
competitive life (`< 6 → <10y`, `6–7.5 → 10–20y`, `> 7.5 → 50y`). `φ` from the
moat tier (`--persistence`).
**Terminal — competitive equilibrium.** By the end of the CAP the return on
**new** invested capital (RONIC) has competed down to the **cost of capital**, so
terminal growth is **value-neutral**: it runs at a GDP-like rate `g_term` but adds
nothing. `RONIC = WACC` ⇒ `RR_term = g_term/WACC` and the terminal collapses to
`NOPAT_N·(1+g)/WACC`:
```
g_eff = min( g_term , 0.99·r ) ;  RONIC = r ;  RR_term = g_eff/r
TV    = NOPAT_N · (1 + g_eff) · (1 − RR_term) / (r − g_eff)  ==  NOPAT_N·(1+g)/r
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

## Cost of capital — equity hurdle → per-company WACC (`--re`)

By default `--r` is a flat firm discount rate. Pass **`--re` (required *equity*
return)** instead and the model discounts each company at its own **WACC**:
```
R_d  = country risk-free base + synthetic credit spread   (Damodaran method)
       synthetic rating from interest coverage = EBIT / (gross debt × R_d),
       EBIT = New Operating Income / (1 − tax), iterated; small-firm table < $2bn
WACC = (MktCap/EV)·R_e + (NetDebt/EV)·R_d           (capped to [4%, 12%])
```
- **Country risk-free** comes from a currency table keyed on "Country of
  Headquarters" (Europe→EUR/SEK/GBP/CHF/…, US→USD swap, Japan→JPY, Korea→KRW).
- The **credit-spread-by-rating curve is one global table** (default risk is
  global); the only country dependence is the risk-free base. Sovereign add-on
  ≈ 0 for US/Europe/Japan/Korea (all IG sovereigns); a `COUNTRY_RISK_PREMIUM` table adds a sovereign spread for EM (India, China, Brazil, Turkey…) on a USD-denominated valuation — override with `--country-crp`.

### De-risking / lever-up over the fade (`--lever-glide`)
Static WACC assumes the capital structure and business risk never change. In
reality, as returns fade the firm de-risks and supports more debt. With
**`--lever-glide`** (requires `--re`) the discount rate becomes **time-varying**:
it holds today's (tax-shielded) WACC through the `n1` hold, then glides linearly
over the `n2` fade to a **mature target-leverage WACC**:
```
target net debt = L × EBIT          (L = sector net-debt/EBIT, TARGET_NETDEBT_EBIT)
mature R_d  = re-rated at that debt: coverage = EBIT/(debt·R_d) = 1/(L·R_d)
mature WACC = re/(1−k),  k = L·(R_d − re/(1−tax))     (debt tax shield; re fixed)
            (closed form from the value-neutral terminal wD = L·WACC/(1−tax))
discounting uses cumulative factors Π(1+WACCₛ); the terminal uses the mature WACC
```
Key properties: the **equity hurdle `re` is held fixed** (it's your opportunity
cost), so WACC falls only through the **tax shield** + cheaper-debt weight at the
mature structure — the MM-clean reading of "they lever up more." **Excess cash is
assumed distributed** (the firm holds target leverage rather than de-levering),
which is what justifies the constant mature structure. The effect is **self-
limiting by debt capacity**: capital-intensive names that carry 3× (industrials,
staples) get a real WACC cut (Volvo 9.98%→8.72%, ~+11% value); asset-light names
that only support 1.5× (software) get ~nothing (Vitec ≈ flat). Override the target
with `--target-lev`. The terminal stays value-neutral (RONIC = mature WACC).
Because the terminal is ~35% of value and scales as `1/(WACC−g)`, this is a
powerful lever — kept bounded by the sector leverage table and the re-rated Rd.

### Refresh rates live each run
The cached `CURRENCY_BASE` (10y govt, ~early-Jun-2026) and `RATING_*` spread
tables (~2024 ICE BofA level) **go stale**. When invoking a valuation, first
**web-search the current 10y government yields** for the countries in the file
(USD: subtract ~40bp for the swap) and, ideally, the current ICE BofA OAS
levels, then pass them in:
```
--re 0.07 --country-base "SEK=0.0273,EUR=0.0303,USD=0.0405,JPY=0.0267,KRW=0.0412"
```
If offline, the script falls back to the cached values (printed in the WACC line
so the user sees what was used). **Exclude banks/financials** — interest
coverage is meaningless when gross debt is customer funding (the output flags
them).

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


## Sales-growth floor (Base Rate Book)

By default each year's growth is **floored** at the **Mauboussin size × industry
*median* sales-growth base rate** (Exhibit 3, "The Impact of Intangibles on Base
Rates", 2021 — 10y median CAGR, keyed on the `Sales` column and GICS industry,
recomputed as sales compound into larger size buckets). The floor expresses that
a firm needs to reinvest *at least* enough to keep its sales growing with its
industry — so the implied RR can't be artificially low (which would inflate FCF).
A **floor, not a cap**: a superior company is never pulled *down* to the industry
average; the floor only lifts `g` when reinvestment-driven growth (`ROIIC·RR_0`)
has faded below it. Disable with `--no-sales-floor` (then `g = ROIIC·RR_0` with a
`g_term` floor in the terminal only). The book covers *growth* only — long-run
**margin** base rates are a separate, not-yet-sourced input.
