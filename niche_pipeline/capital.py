"""
capital.py — the locked capital & quality definitions, as code.

Single source of truth for invested capital and the through-cycle quality anchor
ROIC*, exactly as specified in CAPITAL_SPEC.md and validated against the panel.
Both history.py (the empirical veto) and overearning.py (the over-earning engine)
import from here so the capital base is defined once.

Governing principle — base consistency: New Operating Income capitalizes the
marginal R&D/SG&A investment, so invested capital must include the intangible
capital base, and we keep only 10% of cash (a conservative operating buffer;
netting 100% would pretend the business needs zero operating cash).
"""
from __future__ import annotations

# Panel column names for the invested-capital components.
IC_COLS = {
    "ltd": "Debt - Long-Term - Total",
    "std": "Short-Term Debt & Current Portion of Long-Term Debt",
    "minority": "Minority Interest - Total",
    "equity": "Shareholders' Equity - Attributable to Parent ShHold - Total",
    "funded": "Funded Status (ASR)",
    "cash": "Cash & Short Term Investments",
    "rd_base": "R&D Capital Base",
    "sga_base": "SG&A Capital Base",
    "noi": "New Operating Income",
}
CASH_KEEP = 0.10          # keep 10% of cash in the ROIC base (subtract 0.9 * cash)


def _g(row, idx, name):
    i = idx.get(name)
    if i is None or i >= len(row):
        return 0.0
    v = row[i]
    return v if isinstance(v, (int, float)) else 0.0


def invested_capital(row, idx, cash_keep=CASH_KEEP):
    """Base-consistent invested capital for one panel row (CAPITAL_SPEC §2):

        LT Debt + ST Debt + Minority + Equity - Funded Status
        - (1 - cash_keep) * Cash + R&D Capital Base + SG&A Capital Base

    Intangible capital base IN; 10% of cash retained; signed - Funded Status; no
    separate lease stock. NOT the panel's own `Invested Capital` column, which
    nets full cash."""
    return (_g(row, idx, IC_COLS["ltd"]) + _g(row, idx, IC_COLS["std"])
            + _g(row, idx, IC_COLS["minority"]) + _g(row, idx, IC_COLS["equity"])
            - _g(row, idx, IC_COLS["funded"])
            - (1.0 - cash_keep) * _g(row, idx, IC_COLS["cash"])
            + _g(row, idx, IC_COLS["rd_base"]) + _g(row, idx, IC_COLS["sga_base"]))


# Backwards-compatible alias (overearning.py used this name).
def ic_star(row, idx, cash_keep=CASH_KEEP):
    return invested_capital(row, idx, cash_keep)


def roic_series(rows, idx, cash_keep=CASH_KEEP):
    """Per-year NOI / Invested Capital for a company's sorted rows."""
    out = []
    for r in rows:
        ic = invested_capital(r, idx, cash_keep)
        noi = _g(r, idx, IC_COLS["noi"])
        out.append((noi / ic) if ic else None)
    return out


def roic_star(rows, idx, window=7, cash_keep=CASH_KEEP):
    """The through-cycle quality anchor: the `window`-year rolling average of the
    base-consistent level ROIC (CAPITAL_SPEC §5). Cycle-stable and base-
    consistent — the measure used for the quality veto and the durability barrier.
    Returns None if no valid years."""
    s = [x for x in roic_series(rows, idx, cash_keep)[-window:] if x is not None]
    return sum(s) / len(s) if s else None
