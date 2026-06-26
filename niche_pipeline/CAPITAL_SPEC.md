# Capital & Quality Spec — single source of truth

This is the locked, reconciled definition of operating income, invested capital,
marginal investment and the through-cycle quality anchor (ROIC\*) that the
pipeline builds against. It is derived from the model's own **Financial Metrics
Definitions** doc (`FINANCIAL_METRICS.pdf`, the *updated* table) and validated
against the whole-universe panel (`*-Moat_5_binary.xlsb`). Where we deviate from
the doc, the deviation and its rationale are stated explicitly.

The governing principle is **base consistency**: the numerator (operating income)
capitalizes the marginal R&D/SG&A investment, so the denominator (invested
capital) must include the corresponding intangible *stock*, and the marginal
investment flow must include the same intangible *growth*. Get those on the same
footing and the stock and the flow reconcile over any two periods.

---

## 1. New Operating Income (NOI) — the numerator

```
NOI = ( EBITA Average (7yr rolling) + R&D Growth + SG&A Growth ) × (1 − tax_by_instrument)
```

- `EBITA Average (7yr rolling)` = macro-weighted 7-yr rolling average EBITA margin × current Revenue.
  The margin is already cycle-smoothed (rolling, macro-weighted), **but it is
  applied to *current* revenue** — so revenue-level cyclicality is NOT handled
  here (see §6, the open revenue-normalization work).
- `R&D Growth` = year-over-year change in `R&D Capital Base` (the marginal R&D investment).
- `SG&A Growth` = year-over-year change in `SG&A Capital Base`.
- The `+ R&D Growth + SG&A Growth` term is the intangible-capitalization add-back.
  It is real and material: NOI's implied pre-tax operating margin runs ~10pp
  above the EBITA margin for R&D-heavy names (NVIDIA +10.3pp, ASML +7.5pp),
  ~0pp for non-R&D names (Somero +0.5pp). **NOI does not "go from EBITA."**

Panel column: `New Operating Income` (already doc-compliant — use as-is).

## 2. Invested Capital (IC) — the ROIC base

```
IC = LT Debt + ST Debt + Minority Interest + Shareholders' Equity
     − Funded Status
     − 0.9 × Cash & Short-Term Investments          ← keep 10% of cash (see note)
     + R&D Capital Base + SG&A Capital Base
```

Component columns: `Debt - Long-Term - Total`, `Short-Term Debt & Current
Portion of Long-Term Debt`, `Minority Interest - Total`, `Shareholders' Equity -
Attributable to Parent ShHold - Total`, `Funded Status (ASR)`, `Cash & Short
Term Investments`, `R&D Capital Base`, `SG&A Capital Base`.

**Deviation from the doc (deliberate).** Doc §1.8 nets *full* cash; the panel's
`Invested Capital` / `Ratio` columns implement that. We instead **keep 10% of
cash** (subtract `0.9 × Cash`). Rationale: netting 100% of cash assumes the
business needs zero operating cash, understating the capital it ties up. Keeping
10% credits a deliberately generous operating buffer (real need is usually below
10% of the balance), so IC is conservative and ROIC is, if anything, understated
— the correct direction to err for a quality screen. This matches the doc's own
**Net debt** convention (§2.2, `Gross debt − 0.9 × Cash`).

Because of this, **do NOT read the panel's `Ratio` / `Ratio_rolling_avg_*`
columns directly** (they use full cash). Recompute IC from the components above
with the `0.9 × Cash` factor, then `NOI / IC`.

Other notes:
- No separate capitalized-lease *stock* in IC. Leases enter via the marginal
  side (`Base Physical Capital & Leases = Change in Operating Lease Liabilities`)
  and the Net-debt/valuation bridge.
- `− Funded Status` (signed), not `+ |Funded Status|`.

## 3. Marginal Invested Capital — the reconciling flow

```
Marginal Investments = Operating Working Capital Growth
                     + Base Physical Capital & Leases            (= ΔOperating Lease Liabilities)
                     + M&A and Inorganic Growth (lagged)
                     + Organic Physical Growth (lagged)
                     + Organic Intangible Growth (lagged)        (= R&D Growth + SG&A Growth, lagged)
```

Includes **working capital growth** and **intangible growth** — symmetric with
IC. This symmetry is what makes the two-period reconciliation close.

## 4. Reconciliation invariant (consistency check)

Over any window, the change in IC should equal cumulative marginal investments,
up to the structural maintenance/baseline wedge:

```
ΔIC(t→t+n)  ≈  Σ Marginal Investments(t→t+n)   + maintenance/baseline residual
```

Validated on NVIDIA over 7y: `ΔIC(core) = 104.8` vs `ΣMI + ΔWC = 107.8` →
residual −3.0 on a $100bn+ base (closes once the intangible growth, +38.6, is on
both sides). The residual is *growth above the 3% baseline, net of depreciation*
vs *all retained capital* — expected, not a missing component. Use this bridge
as a name-by-name consistency test when running the universe.

## 5. Through-cycle quality anchor — ROIC\*

```
ROIC*  =  7-year (macro-weighted) rolling average of ( NOI / IC )
```

with IC per §2 (keep-10% cash). This is the **only** quality / capital-return
measure used for selection, the TAM runway signal, and the history veto.

**Measurement rules (hard):**
- **Never** use a marginal ROIC (`ROICm_total - Nyr`, `b_growth`-based) as a
  spot quality read. It is endpoint-sensitive: it *deflates at troughs* and
  *inflates at peaks*. Somero's marginal ROIC reads −67% / +2% / +27% across
  windows purely from cycle position; its true through-cycle ROIC\* is ~29%.
- The level ROIC (§2) rolling-averaged is the robust anchor because it averages
  *levels*, not endpoint *differences*.
- Always read it through the cycle (7yr rolling min), not at a single year.

**Validated anchor values (keep-10% cash, 7yr):**

| Name | ROIC\* |
|---|---|
| NVIDIA | 34.1% |
| ASML | 28.3% |
| MSFT | 21.3% |
| Somero | 29.3% |
| KO | 8.6% |

Somero at ~29% (a high-quality compounder caught in a cyclical trough) is the
canonical case the anchor must get right — and does.

## 6. Open / pending (NOT locked)

- **Revenue normalization (the cyclicality goal).** NOI applies the averaged
  margin to *current* revenue, so a name over-earning on a peak *revenue* base
  is not corrected. Build `Revenue*` separating structural TAM-expansion (keep)
  from a cyclical revenue spike (fade over the industry supply lag). The
  discriminator is the four-input TAM-proximity test (industry, prior growth,
  gross-margin trend, reinvestment × ROIC\*), with the base-rate book as the
  prior (fade to base unless runway is affirmatively shown).
- **Acquirer intangible overlap.** For serial acquirers, the organic SG&A
  capital base may partially overlap acquired goodwill already in equity. Minor
  for low-intangible names; check before relying on ROIC\* for heavy acquirers.

## 7. Code home & wire-in status

The definitions in §2/§5 live in **`capital.py`** (`invested_capital`,
`roic_star`) — the single code source of truth, imported by both `history.py`
and `overearning.py`. Validated anchor: NVIDIA 34.1%, ASML 28.3%, MSFT 21.3%,
Somero 29.3%, KO 8.6%.

Wire-in (both **done**):

1. **`history.py` veto** — `verdict()` anchors on ROIC\* (through-cycle level),
   not the raw 7yr marginal, so quality cyclicals at their trough are not falsely
   CONTRADICTED (Somero: 7y marginal swings 2%/55% but ROIC\* ~29% → SOFT not
   vetoed; C-RAD ROIC\* 1% → CONTRADICTED, which the marginal would have missed).
2. **Over-earning engine** (`overearning.py`) — wired into `pipeline.run()`
   whenever the `--history` panel is supplied (disable with `--no-overearning`).
   The books rank on the over-earning-adjusted return; the Scored sheet carries
   `ROIC*(7y)`, `OE_RevExcess`, `OE_Barrier`, `OE_Faded%`, `OE_H(y)`, `OE_ER_adj`.
