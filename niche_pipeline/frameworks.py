"""
frameworks.py — exact, auditable encodings of the three scoring systems.

  1. Unified Niche Compounder Methodology v3.2   (company moat + core moat)
  2. Engaged Ownership / Battle-Ground Moat v3.0 (Keep-Score / Detachability)
  3. Risk Framework / Constrained Quality Compounder (11 binary tags, gates,
     20% tag rule, return-interpolated sizing)

Everything here is deterministic. The qualitative inputs (the 1-10 chapter
scores, the binary tags, the core-business identification, the shareholder
register) are produced by analyst.py; this module only turns those inputs into
the framework's numbers, bands and verdicts.
"""
from __future__ import annotations

# ----------------------------------------------------------------------------
# 1. UNIFIED NICHE COMPOUNDER v3.2
# ----------------------------------------------------------------------------
# Ten chapters, scored 1-10. Order is fixed (Ch1..Ch10).
NCC_CHAPTERS = [
    "Criticality",          # Ch1  Domain A
    "StandardPremium",      # Ch2  Domain A
    "MarketHegemony",       # Ch3  Domain B
    "EcosystemTrap",        # Ch4  Domain A
    "LifecycleAdvantage",   # Ch5  Domain A
    "SubstitutionThreat",   # Ch6  Domain C (1=active displacement .. 10=only way)
    "DemandHorizon",        # Ch7  Domain C
    "CapitalAllocation",    # Ch8  Domain B
    "ReinvestmentScorecard",# Ch9  Domain B
    "ReverseLindy",         # Ch10 Domain A
]

# "The Calculus of Survival (v3.1)" — weights as written in the field guide.
# They deliberately sum to 1.10 (Ch1 and Ch3 carry 0.15); kept verbatim so the
# program reproduces the document's numbers exactly rather than re-normalising.
NCC_WEIGHTS = [0.15, 0.10, 0.15, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10]

# Domain-Based Information Weighting (internal / external blend) per chapter.
# Used to blend an internal-narrative score with an external-evidence score
# before the chapter enters the weighted sum.
NCC_DOMAIN = {
    "Criticality": "A", "StandardPremium": "A", "MarketHegemony": "B",
    "EcosystemTrap": "A", "LifecycleAdvantage": "A", "SubstitutionThreat": "C",
    "DemandHorizon": "C", "CapitalAllocation": "B", "ReinvestmentScorecard": "B",
    "ReverseLindy": "A",
}
DOMAIN_BLEND = {"A": (0.80, 0.20), "B": (0.20, 0.80), "C": (0.10, 0.90)}


def blend_domain(chapter: str, internal: float, external: float) -> float:
    """Blend internal and external evidence for one chapter per its Domain."""
    wi, we = DOMAIN_BLEND[NCC_DOMAIN[chapter]]
    return round(wi * internal + we * external, 3)


def ncc_score(chapter_scores) -> float:
    """Final NCC v3.2 score from ten chapter scores (Ch1..Ch10)."""
    if len(chapter_scores) != 10:
        raise ValueError("need exactly 10 chapter scores")
    return round(sum(s * w for s, w in zip(chapter_scores, NCC_WEIGHTS)), 3)


def ncc_band(score: float) -> str:
    """v3.2 thresholds."""
    if score > 7.8:
        return "Compounder Target"
    if score >= 6.5:
        return "Watchlist"
    return "Pass"


# ----------------------------------------------------------------------------
# 2. ENGAGED OWNERSHIP / BATTLE-GROUND MOAT v3.0
# ----------------------------------------------------------------------------
# Keep Score (quality only) uses Ch1-7 & Ch10, re-weighted to reward permanence.
# Weights as printed in v3.0 (sum to 1.00).
KEEP_WEIGHTS = {
    "Criticality": 0.17, "MarketHegemony": 0.17, "DemandHorizon": 0.13,
    "SubstitutionThreat": 0.12, "ReverseLindy": 0.12, "StandardPremium": 0.10,
    "EcosystemTrap": 0.10, "LifecycleAdvantage": 0.09,
}


def keep_score(chapter_scores) -> float:
    """KS = sum(chapter * quality weight) over Ch1-7 & Ch10."""
    d = dict(zip(NCC_CHAPTERS, chapter_scores))
    return round(sum(d[c] * w for c, w in KEEP_WEIGHTS.items()), 3)


def detachability_score(entanglement: float, capital_independence: float,
                        core_independence: float) -> float:
    """DS = average(Ch9 entanglement, Ch8 capital-independence, Ch11 core-independence).

    All three expressed so that HIGH = cleaner / safer to exit.
      - entanglement (Ch9): high = turnkey, owns its value chain
      - capital_independence (Ch8 CAPITAL DRAIN): high = cash engine / not a sink
      - core_independence (Ch11 CORE-COUPLING): high = own customers, safe to sell
    """
    return round((entanglement + capital_independence + core_independence) / 3.0, 3)


def eo_action(ks: float, ds: float, core_coupling: float, capital_drain: float) -> str:
    """Action grid (KS x DS) + the Ch11/Ch8 overrides from v3.0."""
    if ks >= 6.5:
        return "KEEP - Structural Core" if ds < 6.5 else "KEEP - Crown Jewel"
    # Low keep score -> "the rest"
    if core_coupling <= 3:
        return "HOLD - coupled to core"          # the lone reason to keep a weak unit
    if ds >= 6.5:
        return "SELL NOW - clean disposal"
    method = "CARVE-OUT / WIND-DOWN"
    if capital_drain <= 3:
        method += " (capital sink - do first)"
    return method


# ----------------------------------------------------------------------------
# 3. RISK FRAMEWORK — Constrained Quality Compounder
# ----------------------------------------------------------------------------
RISK_TAGS = [
    "Customer & End-Market Demand",
    "Literal Customer Concentration",
    "Macro & Factor Sensitivity",
    "Currency & FX Exposure",
    "Geographic & Supply Chain Reliance",
    "Capital Structure & Funding Risk",
    "Technological & Platform Dependency",
    "Product Lifecycle & IP Risk",
    "Regulatory & Compliance Risk",
    "Input Cost Structure",
    "Key Person Risk",
]
RISK_TAGS_SHORT = ["Demand", "CustConc", "Macro", "FX", "Geo/SC", "Funding",
                   "TechDep", "Lifecycle/IP", "Regulatory", "InputCost", "KeyPerson"]

# Binary 1/0 definitions, in plain English, for the analyst and for audit.
RISK_TAG_DEFS = {
    "Customer & End-Market Demand": "Highly reliant on a specific cyclical spender (US Consumer Discretionary, Enterprise IT CapEx, SMBs, Government/Defense, Healthcare Providers).",
    "Literal Customer Concentration": "A single customer is >10% of revenue, or the top 5 are >25%.",
    "Macro & Factor Sensitivity": "Extreme vulnerability to an external force (highly rate-sensitive/long-duration, or commodity-dependent).",
    "Currency & FX Exposure": "International revenue >50%, earnings sensitive to a strong domestic currency.",
    "Geographic & Supply Chain Reliance": ">20% of sales or physical production bottlenecked in a vulnerable region (China, Taiwan).",
    "Capital Structure & Funding Risk": "Net Debt/EBITDA > 3.0x, or requires constant access to capital markets to survive.",
    "Technological & Platform Dependency": "Business model relies on a third-party ecosystem (Apple iOS, AWS, Azure).",
    "Product Lifecycle & IP Risk": ">30% of profit from a single product/patent expiring within 3-5 years.",
    "Regulatory & Compliance Risk": "Prime target for antitrust action, sudden FDA/policy changes, or heavy tariffs.",
    "Input Cost Structure": "Gross margins structurally depressed by volatile, uncontrollable inputs (fuel, raw commodities, unionised labour).",
    "Key Person Risk": "Valuation premium and moat tied to a single, irreplaceable visionary founder/CEO.",
}

# Funding tag (#6) is partly computable from the screener.
def funding_tag_from_financials(net_debt, ebita) -> int | None:
    """Net Debt/EBITDA > 3.0x -> 1. Returns None if inputs missing/negative EBITA."""
    if net_debt is None or ebita in (None, 0):
        return None
    if net_debt <= 0:                      # net cash -> not a funding risk
        return 0
    if ebita <= 0:
        return None                        # ratio undefined; let analyst decide
    return 1 if (net_debt / ebita) > 3.0 else 0


# ---- The Entry Gauntlet (two binary gates) --------------------------------
GATE1_IRR = 0.12          # Expected IRR must be >= 12%
GATE2_DOWNSIDE = 0.20     # P(-30% drawdown) must be <= 20%
# Hard-floor proxy: no-growth value / price. If the no-growth fundamental floor
# is within 30% of price (ratio >= 0.70) a 30% fall lands at/below that floor,
# so further downside needs real earnings deterioration -> low probability.
GATE2_FLOOR_RATIO = 0.70


def gate1_pass(expected_irr) -> bool:
    return expected_irr is not None and expected_irr >= GATE1_IRR


def gate2_pass(no_growth_floor_over_price) -> bool:
    f = no_growth_floor_over_price
    return f is not None and f >= GATE2_FLOOR_RATIO


# ---- Guardrail: implausible-return / normalization-artifact screen -----------
# The ROIIC DCF extrapolates "New Operating Income" forward. On cyclically
# depressed, distressed or loss-normalising names that figure can be far above
# the run-rate the price reflects, throwing off a triple-digit "expected return"
# that is an artifact, not an opportunity (e.g. an AI-disrupted publisher at a
# nominal 110% IRR). Any IRR above this ceiling is flagged and kept out of the
# concentrated book — it would otherwise dominate the return-interpolation.
MAX_PLAUSIBLE_IRR = 0.50


def er_is_artifact(expected_irr) -> bool:
    return expected_irr is not None and expected_irr > MAX_PLAUSIBLE_IRR


# ---- TAM Exhaustion screen (Phase 1) --------------------------------------
def tam_exhausted(payout_to_fcf=None, reinv_now=None, reinv_3y_avg=None,
                  rev_g_1y=None, rev_cagr_3y=None) -> bool:
    """Any one rule trips TAM-exhaustion (permanent exclusion)."""
    if payout_to_fcf is not None and payout_to_fcf > 0.70:
        return True
    if (reinv_now is not None and reinv_3y_avg not in (None, 0)
            and (reinv_now - reinv_3y_avg) / abs(reinv_3y_avg) < -0.25):
        return True
    if (rev_g_1y is not None and rev_cagr_3y is not None
            and rev_g_1y < rev_cagr_3y and rev_g_1y < 0.05):
        return True
    return False


# ---- Portfolio construction (Satellite book) ------------------------------
W_MIN, W_MAX = 0.05, 0.20
TAG_CAP = 0.20            # no tag bucket may exceed 20% of total portfolio weight


def interpolate_weights(irr_by_name: dict) -> dict:
    """Return-based linear interpolation between the 5% floor and 20% ceiling."""
    if not irr_by_name:
        return {}
    rmin, rmax = min(irr_by_name.values()), max(irr_by_name.values())
    out = {}
    for t, r in irr_by_name.items():
        if rmax == rmin:
            out[t] = W_MIN
        else:
            out[t] = W_MIN + (r - rmin) / (rmax - rmin) * (W_MAX - W_MIN)
    return out


def trim_to_tag_cap(book, irr, tags_by_name):
    """Apply the 20% rule via the Trim Protocol: while any tag bucket exceeds
    20% of total portfolio weight, liquidate the lowest-IRR name in the worst
    bucket and re-interpolate. Returns (final_book, weights, log)."""
    book = list(book)
    log = []
    while True:
        w = interpolate_weights({t: irr[t] for t in book})
        buckets = {}
        for t in book:
            for tag in tags_by_name[t]:
                buckets[tag] = buckets.get(tag, 0.0) + w[t]
        over = {tag: bw for tag, bw in buckets.items() if bw > TAG_CAP + 1e-9}
        if not over or len(book) < 2:
            return book, w, log
        worst = max(over, key=over.get)
        in_bucket = [t for t in book if worst in tags_by_name[t]]
        drop = min(in_bucket, key=lambda t: irr[t])
        log.append((worst, round(over[worst], 4), drop, round(irr[drop], 4)))
        book.remove(drop)


# ---- Ownership block test (jurisdiction-aware) ----------------------------
# The blocking stake depends on the special-resolution / scheme threshold:
#   Japan  : special resolutions need 2/3 of votes  -> a >1/3 holder vetoes.
#   UK     : a scheme of arrangement needs 75%       -> a >25% holder vetoes.
#   US/EU  : commonly 2/3 by default.
BLOCK_THRESHOLDS = {            # (HARD-BLOCK >=, SOFT-BLOCK >=)
    "Japan": (33.4, 20.0),
    "United Kingdom": (25.0, 10.0),
    "United States": (33.4, 20.0),
    "_default": (33.4, 20.0),
}


def block_thresholds(country):
    return BLOCK_THRESHOLDS.get((country or "").strip(), BLOCK_THRESHOLDS["_default"])


def ownership_verdict(largest_bloc_pct, country=None) -> str:
    if largest_bloc_pct is None:
        return "UNKNOWN"
    hard, soft = block_thresholds(country)
    if largest_bloc_pct >= hard:
        return "HARD-BLOCK"
    if largest_bloc_pct >= soft:
        return "SOFT-BLOCK"
    return "CONTESTABLE"
