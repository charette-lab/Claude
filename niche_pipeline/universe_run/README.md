# Universe run — completing & valuing the full ~3,300-name book

This folder records the workflow used to (1) complete the qualitative research
over the whole universe and (2) produce a fully research-aware, current-price
scored book.

## Inputs used (not committed — provide your own)

| Input | What it is |
|---|---|
| `LTM_*.xlsx` | current screener — prices, Market Cap, ROICm 7, RR 7, Moat Score, tax, debt |
| `30_file_*.xlsb` | 30-year history panel (3 parts), read via `panel30.py` — fundamentals + ROIC/RR/margins |
| a prior scored book | carries the **derived** v3.2 moats for names already researched |

## Step 1 — complete the research (`research_names.py`)

`to_research.json` is the list of universe names that lacked a cached research
record (699 names in the run this folder documents). With `ANTHROPIC_API_KEY`
set:

```bash
python3 research_names.py to_research.json --workers 8
```

Each name costs ~$0.79 (≈134k input tokens of accumulated web-search context +
~3k output + ~5 web searches). Resumable: cached names are skipped. The 699
records produced are archived in `research_cache_699.tar.gz` — untar into
`niche_pipeline/.cache/` to reuse for free.

## Step 2 — full research-aware valuation (`run_unified.py`)

```bash
AIP_VALUE_ENGINE=/path/to/roiic_dcf.py \
python3 run_unified.py \
    --ltm   LTM_current.xlsx \
    --panel 30_file_1.xlsb 30_file_2.xlsb 30_file_3.xlsb \
    --prior Universe_scored_with_research.xlsx \
    --out   Universe_final.xlsx
```

`run_unified.py` uses real cached research wherever it exists; for names whose
**raw** record is absent but whose **derived** moat is known (in `--prior`), it
injects that moat by solving for the chapter level that reproduces it exactly
(`ncc_score` is monotonic). Selection therefore runs on the true researched
moats. Nothing fabricated reaches the output — the book emits the moat scalar,
never the chapter vector. `research=False`, so no API calls / no spend.

## The clean path (if you have the full original `.cache`)

`run_unified.py` and `--prior` exist only because the raw research records for
the previously-researched names were not on hand. If you have the original
`.cache` for every name, skip all of the above and just run the stock pipeline:

```bash
# untar research_cache_699.tar.gz into niche_pipeline/.cache first, then:
python3 ../complete_universe.py \
    --ltm LTM_current.xlsx \
    --hist 30_file_1.xlsb 30_file_2.xlsb 30_file_3.xlsb \
    --out Universe_final.xlsx
```

That is fully free (all research cached) and is the canonical, no-injection run.
