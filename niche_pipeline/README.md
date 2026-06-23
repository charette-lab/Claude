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

## Notes
- **Gate 2 proxy.** The hard-valuation-floor downside test is approximated by
  the screener's no-growth value / price (≥0.70 ⇒ a 30% fall lands at/below the
  no-growth floor ⇒ low probability of −30%). Swap in a modelled downside
  probability if you have one.
- **Funding tag (#6)** is overwritten with the hard number (Net Debt/EBITDA >
  3.0x) whenever the financials are present, rather than trusting the analyst.
- Without an `ANTHROPIC_API_KEY` the program still runs every deterministic
  module; the researched fields are simply left blank.
