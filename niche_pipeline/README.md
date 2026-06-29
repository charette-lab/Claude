# Niche Compounder Pipeline

One program that takes a list of companies in Excel and runs the full research
stack, producing every output our frameworks call for — and **filling in the
moat and risk scores when they are missing**.

For each company it produces:

| Output | Framework |
|---|---|
| **Company moat** (blended, 10 chapters + band) | Unified Niche Compounder Methodology v3.2 |
| **Core business** + **core moat** (score the durable core only) | Niche Compounder v3.2, scored on the core |
| **Keep Score / Detachability / Action** for the core | Engaged Ownership / Battle-Ground Moat v3.0 |
| **Ownership block test** (can you force change, or will a holder veto?) | Japan 2/3 special-resolution test |
| **AIP valuation** — operating value, ER@7%, ER@12% | aip-value ROIIC persistence-fade DCF |
| **Implied moat length** (years of excess return the price implies) | aip-calc-moat-length reverse DCF |
| **11 binary risk tags** + the two **entry gates** | Risk Framework / Constrained Quality Compounder |
| **Implied-moat regime** — is the return cheapness or a discount-rate artifact? | reverse DCF vs warranted CAP |
| **Satellite Book** — gated, 20%-tag-capped, return-sized | Risk Framework portfolio construction |
| **Core Index** — gated, equal-weighted, 6/30 tag rule | Constrained Quality Compounder Index |

## How it works

- `frameworks.py` — exact, auditable encodings of all three scoring systems
  (NCC v3.2 weights & bands, EO v3.0 Keep-Score/Detachability/action grid, the
  11 risk tags, the two gates, the TAM screen, the 20% tag rule and the
  return-interpolated sizing engine). Pure functions, no I/O.
- `aip.py` — thin wrapper over the proven `aip-value` engine (`roiic_dcf.py`):
  per-company WACC, lever-glide, expected return at two hurdles, and the
  reverse-DCF implied moat length. The engine is reused, not re-derived.
- `capital.py` — the locked capital & quality definitions as code (see
  `CAPITAL_SPEC.md`): base-consistent invested capital (intangible capital base
  in, 10% of cash retained) and the through-cycle quality anchor **ROIC\***.
  Single source of truth, imported by `history.py` and `overearning.py`.
- `overearning.py` — revenue over-earning normalization. NOI applies a
  7yr-averaged margin to *current* revenue, so a name over-earning on a transient
  supply/demand imbalance is priced off an inflated base. Detects the revenue
  spike, gates it on margin corroboration (a real scarcity rent shows margin
  *and* revenue elevated; a structural step lifts revenue at flat margin), sizes
  the durable kept fraction from the reproduction barrier (physical + intangible
  capital base) and the capacity response (is supply being built?), and fades the
  transient remainder over the industry lead time. Wired into `pipeline.run()`
  when `--history` is supplied (disable with `--no-overearning`).
- `analyst.py` — the qualitative research backend. One structured Claude call
  per company (Anthropic Messages API + the `web_search` server tool) returns
  the ten moat chapter scores (company **and** core), the core identification,
  the shareholder register and the eleven risk tags as strict JSON. Results are
  cached to `.cache/` so re-runs are free.
- `pipeline.py` — the orchestrator. Reads the screener, runs every deterministic
  module, calls the analyst **only for fields the sheet does not already
  contain**, then writes the output workbook (`Scored` + `Satellite Book`).
- `decompose.py` — splits each name's expected return into **carry** (internal
  ROIIC compounding) vs **re-rating** (price closing to intrinsic), and builds a
  carry-ranked alternative Core index. See *Return decomposition* below.

## Usage

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-...          # enables the research backend

# Full run (researches missing moat/risk scores via web search):
python3 pipeline.py screener.xlsx -o scored_output.xlsx \
        --re 0.07 --re2 0.12 --country-base "JPY=0.0265,USD=0.041"

# Deterministic only (valuation + gates from financials; uses any scores
# already present in the sheet, researches nothing):
python3 pipeline.py screener.xlsx -o scored_output.xlsx --no-research

# Score just the first N rows:
python3 pipeline.py screener.xlsx --limit 25
```

### Input
Any Excel screener using the standard column names (`Instrument`,
`Company Name`, `GICS Industry Group Name`, `Country of Headquarters`,
`New Operating Income`, `ROICm 7`, `RR 7`, `Moat Score`, `Market Cap`,
`Net debt`, `Gross debt`, `Income Tax Rate - Instrument`, `Sales`, and
`Value per share without growth/share price` for Gate 2).

If the sheet already carries `NCC_<Chapter>` / `Core_<Chapter>` columns or the
eleven named risk-tag columns, those are used as-is and **not** re-researched —
that is the "fill in the missing information" behaviour.

### Output
- **Scored** — one row per company: company moat, core moat, Keep/Detachability/
  action, ownership verdict, AIP value, ER@7%/ER@12%, implied moat, the
  implied-moat regime (`ValuationCushion` / `PricedFor` / `ReturnFromDiscount`),
  both gates, and the risk tags.
- **Satellite Book** — the high-conviction sleeve: names that clear the gauntlet,
  sized by return interpolation (5–20%) with the 20% tag rule enforced greedily.
- **Core Index** — the diversified, **equal-weighted** book: the gauntlet-clearers
  (history-vetoed) ranked by return and filled to `--core-n` (default 30) under the
  20% tag rule as a slot cap (no tag in more than 6 of 30). `--core-strict` enforces
  the cap on **raw** tags (the stringent budget) — in a quality universe this binds
  hard and usually yields fewer than `--core-n` names, which is the real signal that
  the universe cannot diversify the correlated factors any further.
- **Risk Exposure** — the Core Index's raw factor counts (each tag's share of the
  book), with a note on why segmented aggregates can exceed 20%.

## The restructuring lens (core vs whole company)

Selection runs on **two additive paths** — the spin-out is optional upside, never
a prerequisite:

- **AS-IS winner.** The whole-company return already clears the 12% hurdle (and
  Gate 2, no artifact, with a Watchlist+ moat in the company *or* the core). These
  qualify **regardless of ownership** — you are not relying on a transaction, so a
  founder-controlled HARD-BLOCK does not disqualify them (e.g. Fagerhult, ~48%
  Latour-controlled, earns ~14% as-is and belongs in the book).
- **FREED-CORE.** The as-is return does *not* clear, but a strong, **separable**
  core re-rates over the hurdle. This is the only path that needs the register to
  permit change (not a HARD-BLOCK) and a Watchlist+ **core**. It *adds* names the
  as-is screen misses; a cheap pre-screen ("would this clear if its core were a
  top compounder?") makes sure they get researched.

Supporting signals (shown for every researched name):
- **Separated re-rating** — each name is also valued on its **core** moat (a longer
  competitive-advantage period): `ER@12%(separated)`, with `SeparationUplift` the
  difference from as-is.
- **Trapped-jewel flag** — set when the core is Watchlist+ (≥6.5), the **moat gap**
  (core − company) is ≥1.0, and the register is separable. It marks freeable upside;
  for an as-is winner that upside is a bonus, for a freed-core pick it is the thesis.
- **ReturnBasis / effective return** — `as-is` when the name clears on its own
  (sized on the as-is return, the value you capture without action); `freed-core`
  only when separation is what carries it over the line.

Limitation: the separated return captures the **moat-longevity** re-rating from
pricing the core's own CAP; it does not add the margin/ROIC uplift from divesting
a loss-making "rest" (no segment financials in the screener), so it is a *floor*
on the true break-up value.

## Expected-return basis (the gate hurdle vs the discount rate)

Each name is valued at two costs of capital: `ER@re` (default 7%) and `ER@re2`
(12%). A lower discount rate lifts intrinsic value, so `ER@re` is the higher read.
Gate 1 tests the **realistic** read (`ER@re`) against the **12% hurdle** — for a
market with a low risk-free rate (Japan ~2.65%) a 7% cost of equity is the
appropriate discount, and gating on the 12%-discount read would double-count
conservatism and screen out genuine compounders. The hurdle (12%) is unchanged
either way. Pass `--conservative-er` to gate on the `ER@re2` read instead.

## Implied-moat regime — cheapness vs a discount-rate artifact

A high expected return can mean two very different things, and the reverse DCF
tells them apart. `aip.implied_moat` solves the competitive-advantage period the
**price** implies (at 12%); `aip.warranted_life` is the CAP the **moat** deserves.
Comparing them sets `ValuationCushion`:

- **buffer** — the price implies far fewer excess-return years than the moat
  warrants (the market is sceptical): the return is *cheapness*, a real margin of
  safety. This is the typical case for the book.
- **none** — the price already implies the full (or a perpetual, `>150y`) moat:
  there is *no* valuation cushion. When such a name still shows a high `ER@re` that
  collapses at the 12% discount, `ReturnFromDiscount=YES` flags it — the return is
  **borrowed from the low cost of capital, not earned from a depressed price** (the
  Lasertec case: high headline ER, but the price needs the advantage to last
  essentially forever). Not an exclusion — a "priced for perfection" warning.

## Return decomposition — carry vs re-rating (`decompose.py`)

Is a name's modelled return the **internal drive of ROIIC**, or a **revaluation**?
`decompose.py` splits each company's expected return into two additive parts
(`ER = carry + re-rate`), straight from the engine's `er1_carry` / `er1_rerate`:

- **carry** — the IRR if you **exit at the price you paid** (entry EV held
  constant): interim distributions + ROIIC-funded growth. The internal return,
  earned even if the multiple never moves. A fairly-valued compounder is almost
  all carry.
- **re-rate** — `ER − carry` — the bonus from the price **closing to intrinsic
  value** over the horizon. The cheapness/revaluation component. A deep-value
  name is almost all re-rate; a name priced above value has a *negative* re-rate.

It also builds a **carry-ranked alternative Core index** (the "compounder" book)
under the same 6/30 segmented-tag rule as the production index — just swapping the
ranking key from total return to carry — so you can compare a compounder book to
the default value-ranked one.

```bash
# Decompose a book's Core Index (or --scope clears / all):
AIP_VALUE_ENGINE=/path/to/roiic_dcf.py python3 decompose.py \
    --book Universe_final.xlsx --ltm LTM_current.xlsx --out decomposition.xlsx

# Carry-ranked compounder index — same pool re-ranked (gate=value),
# or quality-only with the total-ER gauntlet dropped (gate=loose):
python3 decompose.py --book Universe_final.xlsx --ltm LTM_current.xlsx \
    --mode compounder --gate loose --carry-min 0.12 --out compounder.xlsx
```

Programmatic: `from decompose import decompose, build_carry_book` — `decompose(fin)`
returns `{er, carry, rerate, carry_share, driver, ...}`. Financials come from the
LTM screener (the columns `aip` already uses); the engine and the slot-cap rule are
reused, so the decomposition stays in lockstep with the valuation and the book.

## Empirical history cross-check (`--history`)

The qualitative moat scores are a narrative read; pass a long-run time-series
panel and `history.py` tests them against each company's actual multi-year
economics — the v3.2 Domain-B rule that *real ROIC takes precedence over the
corporate narrative*. Per name it reads:

- **marginal ROIC** (`ROICm 7y`) — return on the capital actually deployed (the
  forward-compounding signal; the ROIC *level* is ignored as a disqualifier
  because legacy goodwill on acquisitive names depresses it);
- **margin trend** — latest EBITA margin vs its 10-year average;
- **cycle-normalised EV/EBITA** (`current_multiple`, mid-cycle margin × current
  sales) for a clean valuation read.

It emits a verdict per name — **CONFIRMED / SOFT / CONTRADICTED** — and a
**CONTRADICTED name is vetoed from the book**: a Watchlist+ moat was claimed but
the economics don't support it (marginal ROIC below the cost of capital, or
margins materially eroded). This auto-corrects exactly the case where research
over-rates a business — e.g. a "specification-led lighting compounder" that has
actually earned ~1% on incremental capital while its margin halved. Run:

```bash
python3 pipeline.py screener.xlsx -o scored.xlsx --history panel_timeseries.xlsx ...
```

## Guardrails (fully hands-off)
A whole-universe run completes unattended:

- **Two-pass execution.** Pass 1 values every company and applies the gates with
  no network calls; Pass 2 (the only paid step) researches **only the Gate-1
  survivors** — so you never spend tokens scoring moats for names that already
  fail the 12% return hurdle. `--research-all` overrides this.
- **Concurrent research** with `--workers N` (default 6); a single failed or
  rate-limited call is logged and skipped, never sinking the run. `analyst.py`
  retries with backoff and caches every success.
- **ER-artifact screen.** The ROIIC DCF can print triple-digit "expected returns"
  on cyclically depressed / distressed names (e.g. an AI-disrupted publisher at a
  nominal 110% IRR). Any IRR above `MAX_PLAUSIBLE_IRR` (50%) is flagged
  `ER_Artifact` and kept out of the concentrated book, where it would otherwise
  dominate the return-interpolation.
- **High-moat Phase-1 filter.** The Satellite only draws from names with a core
  moat ≥ 6.5 (Watchlist+), capped at the doc's 8–12 by IRR before sizing.
- **Output sanitization.** A researched score vector of the wrong length or with
  non-numeric values is coerced (or dropped as NEEDS_RESEARCH) rather than
  silently corrupting a moat score.

## Notes
- **Gate 2 proxy.** The hard-valuation-floor downside test is approximated by
  the screener's no-growth value / price (≥0.70 ⇒ a 30% fall lands at/below the
  no-growth floor ⇒ low probability of −30%). Swap in a modelled downside
  probability if you have one.
- **Funding tag (#6)** is overwritten with the hard number (Net Debt/EBITDA >
  3.0x) whenever the financials are present, rather than trusting the analyst.
- Without an `ANTHROPIC_API_KEY` the program still runs every deterministic
  module; the researched fields are simply left blank.
