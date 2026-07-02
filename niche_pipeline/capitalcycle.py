"""
capitalcycle.py — the reproduction-cost / capital-cycle layer.

A breakthrough or a structural shortage triggers a capital cycle (Perez): demand
outstrips supply, early returns on marginal capital are exceptional, capital
floods in to reproduce the constrained asset, and supply eventually overshoots —
collapsing the scarcity rent. The duration is the gestation period of the
physical infrastructure (component supply-shocks ~3-5y; installation ~5-10y;
broad infrastructure ~10-15y).

The fadeable over-earning is the scarcity rent ABOVE the reproduction-cost
equilibrium (the price/margin at which RMIC = WACC for a greenfield reproduction).
A high *physical* reproduction cost during a shortage is NOT a durable barrier —
it is the signature of a temporary rent that capital will compete away. Only an
*intangible* moat (ecosystem, brand, IP) is irreproducible and therefore durable.

This module classifies each name's industry into a capital-cycle phase with its
gestation horizon, detects which industries are currently in a build-out
(data-driven), and carries explicit tags for live supply-chain cycles the balance
sheet cannot see (the AI infrastructure cycle, whose upstream bottleneck —
CoWoS/HBM/fab/power — makes even a fabless designer like NVIDIA a component-phase
scarcity-rent holder).
"""
from __future__ import annotations

# Gestation horizon (years) over which a scarcity rent fades, by phase.
GESTATION = {"component": 4, "installation": 7, "infrastructure": 12, "intangible": None}

# GICS industry-group substring -> capital-cycle phase. Physical/reproducible
# industries carry a scarcity-rent cycle; intangible-moat industries do not.
INDUSTRY_PHASE = [
    ("Semiconductor", "component"),
    ("Technology Hardware", "component"),
    ("Materials", "infrastructure"),
    ("Energy", "infrastructure"),
    ("Utilities", "installation"),
    ("Transportation", "infrastructure"),
    ("Capital Goods", "installation"),
    ("Automobiles", "installation"),
    ("Real Estate", "installation"),
    ("Telecommunication", "installation"),
    # intangible / non-reproducible moats — no physical reproduction cycle
    ("Software", "intangible"),
    ("Media", "intangible"),
    ("Pharmaceutical", "intangible"),
    ("Biotech", "intangible"),
    ("Life Sciences", "intangible"),
    ("Consumer", "intangible"),
    ("Food", "intangible"),
    ("Household", "intangible"),
    ("Health Care", "intangible"),
    ("Commercial", "intangible"),
    ("Financ", "intangible"),
    ("Bank", "intangible"),
    ("Insurance", "intangible"),
    ("Retail", "intangible"),
]


def industry_phase(industry):
    ind = industry or ""
    for k, v in INDUSTRY_PHASE:
        if k.lower() in ind.lower():
            return v
    return "installation"          # default to a physical phase if unmatched


# Explicit live-cycle overlays — supply-chain bottlenecks the firm balance sheet
# cannot reveal. GICS is too coarse to isolate the AI supply chain from the rest
# of semis/tech-hardware, so the AI cycle is applied at the NAME level: the actual
# picks-and-shovels of the 2024-28 AI infrastructure build-out (advanced logic,
# HBM/memory, the semicap/CoWoS/packaging tools, AI networking, datacenter power).
# This is a hand-maintained judgment list — extend it as the supply chain evolves.
EXPLICIT_CYCLES = {
    "AI infrastructure 2024-2028": {
        # Only names whose CONSOLIDATED economics ARE the AI-chip scarcity rent.
        # Excludes diversified names whose margin is structural, not a chip-price
        # rent: Broadcom (software-heavy via VMware), Eaton (broad electrical),
        # Marvell — those are left to the detector+reversion, which spares them.
        "tickers": {
            # compute & accelerators (pure)
            "NVDA.OQ", "AMD.OQ", "TSM.N", "2330.TW",
            # HBM / memory
            "000660.KS", "MU.OQ", "005930.KS",
            # pure semicap / metrology / packaging tools
            "ASML.AS", "6920.T", "6857.T", "8035.T", "AMAT.OQ", "LRCX.OQ",
            "KLAC.OQ", "BESI.AS", "8036.T", "ACLS.OQ", "ONTO.N", "CAMT.OQ",
            # AI networking & optical (pure-play)
            "ANET.N", "COHR.N", "LITE.OQ",
            # datacenter power & thermal (pure-play)
            "VRT.N",
        },
        "phase": "component", "gestation": 4,
    },
}


def explicit_cycle(industry, ticker=None):
    """Return (phase, gestation, cycle_name) if an explicit live cycle covers this
    name, else None. Matched by ticker (name-level supply-chain membership)."""
    for name, c in EXPLICIT_CYCLES.items():
        if ticker is not None and ticker in c.get("tickers", ()):
            return c["phase"], c["gestation"], name
    return None


def detect_buildout(by, idx, cap_thr=0.06, marg_thr=0.08, min_names=4):
    """Data-driven capital-cycle detector. Per GICS industry group, flag a
    build-out when the median name is BOTH expanding physical capacity (gross
    PP&E growth) AND earning an elevated margin (current vs long-run) — the
    Frenzy/Overbuild signature. Returns {industry: intensity in [0,1]}."""
    import statistics
    ig = idx.get("GICS Industry Group Name")
    ppe_i = idx.get("Property Plant & Equipment - Gross - Total")
    em_i = idx.get("EBITA_Margin")
    agg = {}
    for t, rows in by.items():
        if not rows:
            continue
        ind = rows[-1][ig] if (ig is not None and ig < len(rows[-1])) else None
        if ind is None:
            continue
        # capacity growth (3y gross-PPE CAGR)
        ppe = [r[ppe_i] for r in rows if (ppe_i is not None and ppe_i < len(r)
               and isinstance(r[ppe_i], (int, float)) and r[ppe_i] > 0)]
        cg = ((ppe[-1] / ppe[-4]) ** (1 / 3) - 1) if len(ppe) >= 4 else None
        # margin elevation (recent 3y vs long-run)
        em = [r[em_i] for r in rows if (em_i is not None and em_i < len(r)
              and isinstance(r[em_i], (int, float)))]
        me = None
        if len(em) >= 8:
            recent = statistics.mean(em[-3:]); longrun = statistics.mean(em)
            me = (recent - longrun) / longrun if longrun and longrun > 0 else None
        d = agg.setdefault(ind, {"cg": [], "me": []})
        if cg is not None:
            d["cg"].append(cg)
        if me is not None:
            d["me"].append(me)
    out = {}
    for ind, d in agg.items():
        if len(d["cg"]) < min_names or len(d["me"]) < min_names:
            continue
        mcg = statistics.median(d["cg"]); mme = statistics.median(d["me"])
        if mcg > cap_thr and mme > marg_thr:
            # intensity scales with how far past both thresholds the industry sits
            out[ind] = max(0.0, min(1.0, 0.5 * (mcg / cap_thr - 1) + 0.5 * (mme / marg_thr - 1) + 0.5))
    return out


def classify(industry, buildout_map=None, ticker=None):
    """Resolve a name's capital-cycle phase, gestation horizon, and whether it is
    in an active build-out. Explicit name-level live cycles win; then the
    data-driven industry detector; phase falls back to the industry map.

    Returns (phase, gestation_years, in_buildout: float[0..1], source)."""
    ex = explicit_cycle(industry, ticker)
    if ex:
        phase, gest, name = ex
        return phase, gest, 1.0, name
    phase = industry_phase(industry)
    gest = GESTATION.get(phase)
    intensity = 0.0
    if buildout_map and industry in buildout_map and phase != "intangible":
        intensity = buildout_map[industry]
    return phase, gest, intensity, "detector" if intensity > 0 else "none"
