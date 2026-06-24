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
| **Satellite Book** — gated, 20%-tag-capped, return-sized | Risk Framework portfolio construction |

## How it works

- `frameworks.py` — exact, auditable encodings of all three scoring systems
  (NCC v3.2 weights & bands, EO v3.0 Keep-Score/Detachability/action grid, the
  11 risk tags, the two gates, the TAM screen, the 20% tag rule and the
  return-interpolated sizing engine). Pure functions, no I/O.
- `aip.py` — thin wrapper over the proven `aip-value` engine (`roiic_dcf.py`):
  per-company WACC, lever-glide, expected return at two hurdles, and the
  reverse-DCF implied moat length. The engine is reused, not re-derived.
- `analyst.py` — the qualitative research backend. One structured Claude call
  per company (Anthropic Messages API + the `web_search` server tool) returns
  the ten moat chapter scores (company **and** core), the core identification,
  the shareholder register and the eleven risk tags as strict JSON. Results are
  cached to `.cache/` so re-runs are free.
- `pipeline.py` — the orchestrator. Reads the screener, runs every deterministic
  module, calls the analyst **only for fields the sheet does not already
  contain**, then writes the output workbook (`Scored` + `Satellite Book`).

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
  action, ownership verdict, AIP value, ER@7%/ER@12%, implied moat, both gates,
  and the risk tags.
- **Satellite Book** — the names that clear both gates and are not HARD-BLOCK,
  sized by return interpolation (5–20%) with the 20% tag rule enforced by the
  Trim Protocol (a log of every liquidation is appended).

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
