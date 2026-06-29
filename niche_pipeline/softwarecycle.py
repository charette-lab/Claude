"""
softwarecycle.py — the software reproduction-cost / moat-compression layer.

The demand-side mirror of capitalcycle.py. For a physical name a scarcity rent is
a TRANSIENT over-earning that capital reproduces away (fade DOWN to equilibrium).
For a software franchise the analogue runs the other way: a DURABLE moat whose
reproduction cost is being STRUCTURALLY COMPRESSED by a technology shock
(generative AI), collapsing the equilibrium itself rather than fading an excess
above it.

The reproduction cost of a software franchise splits into two intangible buckets
(Medley / asset-reproduction framework):

  K_tech    ≈ R&D capital base   — the cost to CLONE the product. Reproducible;
                                    genAI compresses it (a clone no longer needs
                                    30 years of hand-built features — a foundation
                                    model generates the output from a prompt).
  K_network ≈ SG&A capital base  — the cost of liquidity: critical-mass users,
                                    switching costs, file-format gravity (.psd/.ai).
                                    Durable; genAI threatens it only if the workflow
                                    shift dissolves the format lock-in.

  durable_moat_frac = K_network / (K_tech + K_network)

measures how much of the barrier is the durable leg. A genAI shock compresses the
competitive-advantage period (CAP) in proportion to (1 - durable_moat_frac): a
pure-K_tech product (Temenos-like) loses most of its CAP; a pure-K_network
franchise (Gartner-like) is near-immune.

WHICH names are actually under assault is a name-level judgment the balance sheet
cannot see — Temenos is K_tech-heavy but a banking core, not genAI-disruptable — so
exposure is carried as a HAND-MAINTAINED tag, exactly like the AI supply-chain
cycle in capitalcycle.py. durable_moat_frac then sets HOW MUCH each tagged name is
compressed.

The valuation consequence is read through the carry/re-rating decomposition (aip):
a name not yet re-rated for the shock is priced for an intact moat (re-rating-
dependent — at risk); a name already compressed to equilibrium (Adobe at ~11.6x) is
priced for the moat to be GONE — its return shows up as CARRY, and if K_network
survives that is a discount to structural earning power (the hidden value). This
module only MEASURES and supplies the compressed-moat override; it does not feed
book selection.
"""
from __future__ import annotations

# CAP-compression severity at FULL vulnerability (durable_moat_frac = 0), in moat-
# score points knocked off the competitive-advantage period, scaled down by each
# name's durable (K_network) share. ~3 points takes a wide-moat 8.0 toward a narrow
# 5.0 — a multi-decade CAP toward a ~5y one (the Adobe 40x -> ~12x re-rating).
GENAI_SEVERITY = 3.0

# Hand-maintained genAI disruption tag — name -> (threat_channel, rationale).
# Membership AND channel are JUDGMENT calls; curate them. The channel selects which
# leg of the reproduction barrier is under assault, because durable_moat_frac alone
# cannot tell a product-cloning threat from a content-commoditization one:
#   "tech"    — the PRODUCT is cloneable (foundation models reproduce the output);
#               the K_tech leg is at risk, so vulnerability = (1 - durable_moat_frac).
#   "network" — the CONTENT / format / liquidity itself is genAI-substitutable
#               (e.g. textbooks, stock imagery); the K_network leg is at risk, so
#               vulnerability = durable_moat_frac. Protects nobody by being "sticky".
#   "both"    — product AND network both dissolve (Adobe: Firefly clones the output
#               AND the workflow shift erodes .psd/.ai gravity); vulnerability = 1.
# Names not in the current panel are kept so the tag is ready when data is added.
EXPLICIT_DISRUPTION = {
    "ADBE.OQ": ("both",    "Adobe — Firefly/Sora/Midjourney clone K_tech; workflow shift "
                           "erodes .psd/.ai format gravity (K_network)"),
    "FIG.N":   ("tech",    "Figma — genAI generates UI/designs from prompts (product clone); "
                           "the multiplayer collaboration network is the surviving defense"),
    "WIX.OQ":  ("tech",    "Wix — genAI builds a whole site from a prompt; thin switching costs, "
                           "so the cloneable product leg is the exposure"),
    "SSTK.N":  ("network", "Shutterstock — generated imagery commoditizes the stock-content "
                           "library/marketplace (the K_network moat itself)"),
    "GETY.N":  ("network", "Getty Images — generic imagery commoditized; editorial/licensed "
                           "content more defensible, but the library is the leg at risk"),
    "CHGG.N":  ("network", "Chegg — homework/answer content directly substituted by ChatGPT "
                           "(already past-tense: operating income negative, no going-concern value)"),
    "PSON.L":  ("network", "Pearson — education content is directly generatable; the threat "
                           "is content/distribution commoditization, not a product clone"),
}
# CANDIDATES left to explicit judgment (NOT active — add with a channel to enable):
#   ADSK.OQ  Autodesk      — CAD/BIM: strong format+engineering-data moat (network-leaning)
#   EA.OQ    Electronic Arts — franchise IP/network; content-gen exposure partial
#   PEGA.OQ  Pegasystems   — low-code/BPM vs AI agents (tech)


def durable_moat_frac(sig):
    """K_network / (K_tech + K_network) from the capital-base shares (tech_share,
    network_share in panel_signals). None if the name carries no measurable
    intangible barrier (so compression cannot be assessed)."""
    tech = sig.get("tech_share") or 0.0
    net = sig.get("network_share") or 0.0
    s = tech + net
    return (net / s) if s > 1e-6 else None


def disruption_tag(ticker):
    """(threat_channel, rationale) if `ticker` is on the hand-maintained genAI
    disruption list, else None. Matched by ticker (name-level membership)."""
    return EXPLICIT_DISRUPTION.get(ticker) if ticker is not None else None


def moat_compression(ticker, sig, moat, severity=GENAI_SEVERITY):
    """Effective (compressed) moat score for a genAI-exposed software name.

    Returns (eff_moat, compression_pts, durable_moat_frac, channel). For an untagged
    name compression is 0 and eff_moat == moat. For a tagged name the CAP is cut by
    severity*vulnerability, where vulnerability is the share of the barrier the threat
    channel attacks: tech -> (1-dmf), network -> dmf, both -> 1. Floored at a no-moat
    1.0. eff_moat is fed to aip.value_and_return as moat_override to model the faster
    ROIIC fade; the rest of the inputs are unchanged."""
    tag = disruption_tag(ticker)
    dmf = durable_moat_frac(sig)
    if tag is None or moat is None:
        return moat, 0.0, dmf, None
    channel = tag[0]
    d = dmf if dmf is not None else 1.0
    vuln = d if channel == "network" else (1.0 if channel == "both" else 1.0 - d)
    comp = severity * max(0.0, min(1.0, vuln))
    return max(1.0, moat - comp), comp, dmf, channel


def analyze(fin, sig, aip, re=0.07, re2=0.12):
    """Analysis overlay (NOT book-feeding): value a software name at its stated moat
    and again at the genAI-compressed moat, and return the decomposition of each.

    Requires the aip module passed in (avoids a circular import). Returns a dict:
      durable_moat_frac, tag, compression, eff_moat,
      er_base / carry_base / rerate_base   (stated moat),
      er_comp / carry_comp / rerate_comp   (compressed moat; == base if untagged),
      er_drop  (er_base - er_comp)          — the disruption's modelled return hit.
    For a tagged, K_network-protected name already re-rated by the market, er_comp
    stays high and carry-rich — the Adobe hidden-value signature."""
    F = aip.FIELDS
    ticker = fin.get(F["ticker"]); moat = fin.get(F["moat"])
    eff_moat, comp, dmf, channel = moat_compression(ticker, sig, moat)
    tg = disruption_tag(ticker)
    base = aip.value_and_return(fin, re=re, re2=re2)
    if not base:
        return None
    out = {"durable_moat_frac": dmf, "tag": (channel if tg else None),
           "channel": channel, "rationale": (tg[1] if tg else None),
           "compression": comp, "eff_moat": eff_moat,
           "er_base": base.get("er1"), "carry_base": base.get("er1_carry"),
           "rerate_base": base.get("er1_rerate")}
    if tg is None or comp <= 1e-9:
        out.update(er_comp=base.get("er1"), carry_comp=base.get("er1_carry"),
                   rerate_comp=base.get("er1_rerate"), er_drop=0.0)
        return out
    shk = aip.value_and_return(fin, re=re, re2=re2, moat_override=eff_moat)
    out.update(er_comp=shk.get("er1") if shk else base.get("er1"),
               carry_comp=shk.get("er1_carry") if shk else base.get("er1_carry"),
               rerate_comp=shk.get("er1_rerate") if shk else base.get("er1_rerate"),
               er_drop=((base.get("er1") or 0) - (shk.get("er1") or 0)) if shk else 0.0)
    return out
