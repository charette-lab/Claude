"""
overearning.py — revenue/over-earning normalization from the long-run panel.

Implements the computable spine of the over-earning architecture (see
CAPITAL_SPEC.md and the cyclicality research thread): New Operating Income
applies a 7-yr-averaged margin to CURRENT revenue, so margin cyclicality is
already handled but REVENUE-level cyclicality is not. When current revenue rests
on a transient supply/demand imbalance (Betancourt scarcity trap), the excess
should fade; when it rests on a durable barrier (Bilbiie sunk-entry-cost moat /
Acemoglu bargaining position), it is kept.

Per name it produces a two-stage expected return:
  Stage 2 — the existing ROIIC DCF on the SUSTAINABLE revenue base
  Stage 1 — the transient revenue excess, faded to zero over horizon H
             (H = industry supply lag scaled by the durability barrier)

Knobs map to the literature:
  * scarcity / revenue spike            -> Betancourt et al. (2024)
  * durability barrier (kept fraction)  -> Bilbiie/Ghironi/Melitz (2007), Acemoglu & Tahbaz-Salehi (2024)
  * fade horizon H = entry delay        -> Savagar & Dixon (2020)
  * through-cycle quality anchor ROIC*  -> CAPITAL_SPEC (base-consistent, keep-10% cash)

What is COMPUTABLE here (panel only): the revenue spike, the through-cycle
quality anchor ROIC*, H, and the reproduction barrier for BOTH asset-heavy
(physical reproduction premium) and asset-light moats (the capitalized R&D/SG&A
capital base — the same intangible stock that sits in invested capital). What is
NOT (needs research, left as default-neutral hooks): qualitative relationship-
specificity beyond the capital base, and demand-durability (take-or-pay /
channel inventory, Check C).
"""
from __future__ import annotations
import math
import aip
import history
import capital
import capitalcycle

# Industry supply lag (years to create competing capacity) — the entry-delay
# clock from Bilbiie. Matched on a GICS-Industry-Group substring.
SUPPLY_LAG = {
    "Materials": 10, "Energy": 8, "Utilities": 15, "Transportation": 12,
    "Capital Goods": 5, "Semiconductors": 5, "Technology Hardware": 5,
    "Pharmaceuticals": 8, "Biotechnology": 8, "Health Care Equipment": 6,
    "Software": 2, "Media": 2, "Food": 3, "Consumer": 4, "Automobiles": 5,
    "Real Estate": 12, "Telecommunication": 10,
}
DEFAULT_LAG = 6


def _num(v):
    return v if isinstance(v, (int, float)) else None


def supply_lag(industry):
    ind = industry or ""
    for k, v in SUPPLY_LAG.items():
        if k.lower() in ind.lower():
            return v
    return DEFAULT_LAG


def cycle_window(industry):
    """Years of history the revenue trend is fit over — a business-appropriate
    cycle length, NOT the full 30-year record. Fitting a single log-trend across
    a very long exponential ramp flags any sustained structural acceleration as a
    'spike'; a cycle-length window tracks recent structure, so a steady compounder
    shows ~no spike while a true cyclical surge still deviates. ~1.6x the supply
    lag, clamped to [8, 21]."""
    return max(8, min(21, round(supply_lag(industry) * 1.6)))


def _col(rows, idx, name):
    i = idx.get(name)
    return [_num(r[i]) if (i is not None and i < len(r)) else None for r in rows]


def _last(xs):
    return next((v for v in reversed(xs) if v is not None), None)


def _median(xs):
    v = sorted(x for x in xs if x is not None)
    n = len(v)
    if n == 0:
        return None
    return v[n // 2] if n % 2 else 0.5 * (v[n // 2 - 1] + v[n // 2])


def _loglinear_current(series, weights=None):
    """Weighted fit of log(y) ~ a + b·t over the valid history; return the fitted
    CURRENT value (the structural trend level today). A steady exponential grower
    has current≈fitted (no spike); a recent surge above its own trend shows
    current>fitted. Returns None if <4 positive points.

    `weights` (macro-cyclical year weights, Finck shock-identification) down-
    weight macro-shock years so a shock-inflated boom does not define the trend
    baseline — the shock revenue then shows as above-trend excess. Defaults to
    equal weights."""
    pts = []
    for i, y in enumerate(series):
        if y is not None and y > 0:
            w = 1.0
            if weights is not None and i < len(weights) and weights[i]:
                w = weights[i]
            pts.append((i, math.log(y), w))
    if len(pts) < 4:
        return None
    W = sum(p[2] for p in pts)
    mx = sum(p[0] * p[2] for p in pts) / W
    my = sum(p[1] * p[2] for p in pts) / W
    den = sum(p[2] * (p[0] - mx) ** 2 for p in pts)
    if den == 0:
        return None
    b = sum(p[2] * (p[0] - mx) * (p[1] - my) for p in pts) / den
    a = my - b * mx
    t_now = pts[-1][0]
    return math.exp(a + b * t_now)


# Base-consistent invested capital (keep-10% cash) — defined in capital.py.
ic_star = capital.ic_star


def panel_signals(rows, idx):
    """History-derived over-earning signals for one company."""
    sales = _col(rows, idx, "Sales")
    noi = _col(rows, idx, "New Operating Income")
    cur_sales = _last(sales)
    industry = None
    ii = idx.get("GICS Industry Group Name")
    if ii is not None and rows and ii < len(rows[-1]):
        industry = rows[-1][ii]

    # --- (Guard 2) revenue spike vs a CYCLE-WINDOW macro-weighted log-trend ---
    # Fit the trend over a business-appropriate window, not the full record, so a
    # steady structural ramp is not mistaken for a cyclical spike.
    macro_w = _col(rows, idx, "Macro Weight")
    W = cycle_window(industry)
    s_win = sales[-W:] if len(sales) > W else sales
    w_win = macro_w[-W:] if len(macro_w) > W else macro_w
    trend = _loglinear_current(s_win, w_win)
    rev_excess_frac = 0.0
    if trend and cur_sales and cur_sales > 0:
        rev_excess_frac = max(0.0, (cur_sales - trend) / cur_sales)
    past_peak = bool(cur_sales is not None and sales
                     and cur_sales < max(v for v in sales if v is not None))
    if past_peak:
        rev_excess_frac = 0.0  # already correcting; no live spike

    # --- (Guard 3) mean-reversion evidence from the firm's OWN long history ---
    # Only fade where the business has actually reverted before. A monotone
    # structural grower (shallow recession dips only) has no reversion precedent
    # and should be spared; a cyclical with deep, repeated drawdowns should fade.
    rev_hist = [v for v in sales if v is not None and v > 0]
    mean_reversion = 0.0
    if len(rev_hist) >= 6:
        peak = rev_hist[0]; maxdd = 0.0
        for v in rev_hist:
            peak = max(peak, v)
            if peak > 0:
                maxdd = max(maxdd, (peak - v) / peak)
        # 15% drawdown = recession noise floor (no reversion credit); full at 45%.
        mean_reversion = max(0.0, min(1.0, (maxdd - 0.15) / 0.30))

    # --- through-cycle quality anchor ROIC* (base-consistent, keep-10% cash) ---
    roic_star = capital.roic_star(rows, idx)

    # --- scarcity corroboration: asset-sweat above trend (Betancourt) ---
    ppe = _col(rows, idx, "Property Plant & Equipment - Gross - Total")
    turns = [(s / p) if (s and p) else None for s, p in zip(sales, ppe)]
    t_now = _last(turns)
    tv = [x for x in turns if x is not None]
    turns_trend = (sum(tv[:-1]) / len(tv[:-1])) if len(tv) > 1 else None
    asset_sweat = ((t_now - turns_trend) / turns_trend) if (t_now and turns_trend) else 0.0

    # --- physical reproduction barrier (Check B) ---
    r = rows[-1]
    def g(name):
        v = _num(r[idx[name]]) if idx.get(name) is not None else None
        return v or 0.0
    grc = g('Gross Reproduction Cost'); ppe_g = g('Property Plant & Equipment - Gross - Total')
    repro_prem = ((grc / ppe_g) - 1) if ppe_g else 0.0
    icv = ic_star(r, idx)
    phys_share = (ppe_g / icv) if icv else 0.0
    # intangible reproduction barrier: the capitalized R&D + SG&A capital base is
    # an asset a competitor must rebuild over years (CUDA-type ecosystem, brand,
    # IP). It is the SAME stock the rest of the model carries in invested
    # capital, so crediting it here keeps the barrier consistent with IC.
    intang_base = g('R&D Capital Base') + g('SG&A Capital Base')
    intang_share = (intang_base / icv) if icv else 0.0
    # K_tech vs K_network split (Medley/reproduction framework): the R&D capital
    # base is the cost to CLONE the tech (K_tech — reproducible, competes away to
    # equilibrium), while the SG&A capital base is the accumulated brand / customer-
    # acquisition spend ≈ K_network = U_min·CAC (the cost to reproduce the network /
    # liquidity / switching costs — durable). Only K_network is a lasting moat.
    tech_share = (g('R&D Capital Base') / icv) if icv else 0.0       # K_tech
    network_share = (g('SG&A Capital Base') / icv) if icv else 0.0   # K_network

    # --- capacity response (CapEx-to-capacity, K_t = K_{t-1}+f(C_{t-L})-δ) ---
    # Gross PP&E stock growth is the reliable capacity-addition signal (the cash-
    # flow CapEx line is sparse/unreliable in this panel). A scarcity rent where
    # capacity is being aggressively added is transient (relief arrives at lead
    # time L); flat capacity means supply is not responding. Only meaningful for
    # asset-heavy capacity owners, so it is weighted by the physical share — a
    # fabless firm's own PP&E is not the constraining capacity (that is its
    # foundry's), so its signal is correctly muted.
    def _cagr3(series):
        v = [(i, x) for i, x in enumerate(series) if x and x > 0]
        if len(v) >= 4 and v[-4][1] > 0:
            return (v[-1][1] / v[-4][1]) ** (1.0 / 3) - 1
        return None
    cap_growth = _cagr3(_col(rows, idx, "Property Plant & Equipment - Gross - Total"))
    # intangible capacity: R&D + SG&A capital base growth. For an intangible-
    # intensive firm the capacity being built is the product/pipeline/ecosystem,
    # not PP&E — and the model capitalizes it, so the capacity signal must see it.
    # (Stock = durability, credited in the barrier; GROWTH = capacity/competitive
    # investment flooding in, credited here as rent-eroding.)
    rd_series = _col(rows, idx, "R&D Capital Base")
    sga_series = _col(rows, idx, "SG&A Capital Base")
    intang_ser = [((rd_series[i] or 0) + (sga_series[i] or 0)) if (i < len(rd_series))
                  else None for i in range(len(rows))]
    intang_cap_growth = _cagr3(intang_ser)

    romic = _last(_col(rows, idx, "ROICm_total - 7 years"))

    # --- margin corroboration (Betancourt scarcity-rent signature) ---
    # A true over-earning rent shows the operating margin DETACHED above its own
    # run-rate. A structural revenue step (M&A, share gain, ordinary volume
    # growth) lifts revenue at NORMAL margins. Gating the revenue fade on margin
    # elevation isolates scarcity rents and exempts structural steps (e.g. a
    # merger-driven revenue jump at flat margin is not faded).
    em = _col(rows, idx, "EBITA_Margin")
    em_cur = _last(em)
    # macro-shock-weighted margin baseline over the full history: down-weighting
    # boom years lowers the baseline so a SUSTAINED multi-year scarcity rent
    # (which contaminates a plain trailing average) still shows as elevated.
    pairs = [(m, (macro_w[i] if i < len(macro_w) and macro_w[i] else 1.0))
             for i, m in enumerate(em) if m is not None]
    em_avg = None
    if pairs:
        Wm = sum(w for _, w in pairs)
        em_avg = sum(m * w for m, w in pairs) / Wm if Wm else None
    margin_elev = ((em_cur - em_avg) / em_avg) if (em_cur is not None and em_avg and em_avg > 0) else 0.0
    # Reproduction-cost margin rent: the recent (NOI-relevant 7yr) margin above the
    # long-run through-cycle equilibrium (em_avg, where RMIC -> WACC). For a
    # physical-bottleneck name in a build-out this is the scarcity premium that
    # capital competes away as it reproduces the constrained asset.
    em7 = [x for x in em[-7:] if x is not None]
    m_recent7 = sum(em7) / len(em7) if em7 else None
    margin_rent_frac = (max(0.0, (m_recent7 - em_avg) / m_recent7)
                        if (m_recent7 and em_avg and m_recent7 > 0) else 0.0)
    # Price-level normalization: the scarcity premium overstates REVENUE, not just
    # the margin. With unit cost ~fixed, an inflated price shows up in the GROSS
    # margin, so when the price reverts revenue falls to (1-gm_recent)/(1-gm_long)
    # of current. Capturing only the margin (and holding revenue at the scarcity
    # price) leaves normalized NOI ~2x too high for a price-spiked name (NVIDIA).
    gm = _col(rows, idx, "Gross Profit Margin")
    gm3 = [g for g in gm[-3:] if g is not None and g < 1]
    gm_recent = sum(gm3) / len(gm3) if gm3 else None
    # MEDIAN equilibria (robust to loss-year outliers): the long-run MEAN margin
    # collapses toward zero for a cyclical/loss-making name (a few deep-loss years),
    # which made cc_rent blow up to hundreds of %. The median is the mid-cycle level.
    gm_eq = _median([g for g in gm if g is not None and g < 1])
    em_eq = _median([x for x in em if x is not None])
    price_norm = 1.0
    if gm_recent is not None and gm_eq is not None and (1 - gm_eq) > 0:
        price_norm = max(0.2, min(1.0, (1 - gm_recent) / (1 - gm_eq)))     # rev_eq / rev_cur
    # Full capital-cycle rent: the NOI fraction that fades to reach the mid-cycle
    # (intangible-inclusive) equilibrium, normalizing BOTH margin AND price. Bounded.
    cc_rent_frac = 0.0
    if m_recent7 and em_eq is not None and m_recent7 > 0 and em_eq > 0:
        cc_rent_frac = max(0.0, min(0.70, 1.0 - (em_eq / m_recent7) * price_norm))
    return {
        "rev_excess_frac": rev_excess_frac, "past_peak": past_peak,
        "roic_star": roic_star, "asset_sweat": asset_sweat,
        "repro_prem": repro_prem, "phys_share": phys_share,
        "intang_share": intang_share, "tech_share": tech_share,
        "network_share": network_share, "romic": romic,
        "margin_elev": margin_elev, "cap_growth": cap_growth,
        "intang_cap_growth": intang_cap_growth,
        "mean_reversion": mean_reversion, "cycle_window": W,
        "margin_rent_frac": margin_rent_frac, "industry": industry,
        "price_norm": price_norm, "cc_rent_frac": cc_rent_frac,
    }


def durability_barrier(sig, moat, wacc):
    """Fraction of the revenue spike that is structurally defensible (kept in the
    base). Blends through-cycle quality, moat band, and the physical
    reproduction barrier (physical premium + intangible capital base). In [0,1].
    The reproduction leg now credits asset-light moats via the R&D/SG&A capital
    base, so CUDA/EUV/brand-type moats are no longer under-credited. A remaining
    research hook (qualitative relationship-specificity / demand-durability)
    would refine it further but is not required for the quantitative barrier."""
    rs = sig.get("roic_star")
    # quality: excess return over cost of capital, saturating at ~+20pp
    q = 0.0
    if rs is not None and wacc:
        q = max(0.0, min(1.0, (rs - wacc) / 0.20))
    # moat band: >7.8 compounder -> strong; 6.5 watchlist -> weak
    m = 0.0
    if moat is not None:
        m = max(0.0, min(1.0, (moat - 6.0) / (8.5 - 6.0)))
    # Durability comes from the IRREPRODUCIBLE moat only. A high *physical*
    # reproduction cost in a shortage is a temporary rent (faded by the capital-
    # cycle margin rent). And an *intangible* moat is durable ONLY when it is
    # K_network (brand / liquidity / switching costs ≈ the SG&A capital base) —
    # NOT K_tech (the R&D capital base = cost to clone the code, which competes
    # away to equilibrium). So credit the network/brand leg, not the R&D leg.
    network = max(0.0, min(1.0, sig.get("network_share", 0.0)))
    durability = 0.45 * q + 0.40 * m + 0.15 * network
    return max(0.0, min(1.0, durability))


# Sector "mature, dominant" EV/EBIT bands — above these the structural business is
# priced beyond what a mature franchise warrants (the dialogue's 20-25x for
# hardware, higher for durable software/network moats).
MATURE_EV_EBIT = {"component": 22, "installation": 18, "infrastructure": 14, "intangible": 30}


def normalized_ev_ebit(ev, nopat, faded_frac, H, wacc, tax, phase):
    """The two-tranche reality check (expectations-investing). Give the company
    FULL credit for banking the scarcity rent as a temporary after-tax annuity,
    then test the multiple on the STRUCTURAL business only — the rent is never
    capitalized into a perpetual multiple:

        Norm EV/EBIT = [ EV − PV(after-tax scarcity rent over H) ] / equilibrium EBIT

    Returns (multiple, overpriced_bool). A multiple above the sector's mature band
    means the price requires the scarcity premium to last forever."""
    if ev is None or nopat in (None, 0) or not wacc:
        return None, False
    t = tax if (tax is not None and 0 <= tax < 1) else 0.20
    eq_nopat = nopat * (1.0 - faded_frac)                 # structural after-tax earnings
    if eq_nopat <= 0:
        return None, False
    eq_ebit = eq_nopat / (1.0 - t)                        # pre-tax structural EBIT
    excess_nopat = nopat * faded_frac                     # scarcity rent (after-tax)
    Hh = int(round(H)) if H else 0
    annuity = sum(1.0 / (1 + wacc) ** k for k in range(1, Hh + 1)) if Hh > 0 else 0.0
    mult = (ev - excess_nopat * annuity) / eq_ebit
    # The OVERPRICED flag is a scarcity-detachment signal, not a generic
    # high-multiple call: it fires only when a name has a real scarcity rent AND
    # the STRUCTURAL business still trades above its mature band — i.e. the market
    # is pricing the temporary rent as permanent. A high multiple with no rent is a
    # growth/quality judgement, not a capital-cycle detachment, so it is not flagged.
    overpriced = mult > MATURE_EV_EBIT.get(phase, 22) and faded_frac > 0.10
    return mult, overpriced


def two_stage_return(fin, sig, re=0.07, re2=0.12, cycle_map=None):
    """ER on the two-stage path: Stage-2 DCF on the sustainable revenue base plus
    the PV of the transient revenue excess fading over H. Returns a dict with the
    unadjusted and adjusted expected returns and the components used."""
    base = aip.value_and_return(fin, re=re, re2=re2)
    if not base:
        return None
    er_current = base["er1"]
    cur_noi = fin.get(aip.FIELDS["nopat"])
    if not cur_noi:
        return {"er_current": er_current, "er_adj": er_current, "faded_frac": 0.0,
                "H": 0.0, "barrier": None, "rev_excess": sig.get("rev_excess_frac", 0.0)}

    barrier = durability_barrier(sig, fin.get(aip.FIELDS["moat"]), base["wacc"])
    rev_excess = sig.get("rev_excess_frac", 0.0)
    # scarcity-rent gate: only fade revenue excess corroborated by an elevated
    # margin (price detaching from cost). 30%+ margin elevation = full rent.
    SCARCITY_FULL = 0.30
    scarcity = max(0.0, min(1.0, sig.get("margin_elev", 0.0) / SCARCITY_FULL))
    # capacity response erodes the barrier: a rent being flooded with new
    # capacity is less durable (relief arrives at lead time L). 12%/yr gross-PP&E
    # growth = full response, weighted by physical share so a fabless firm's own
    # PP&E growth (not the constraining capacity) does not erode its moat.
    # total capacity response = physical + intangible capacity additions, each
    # weighted by where the firm's capacity actually lives (physical vs intangible
    # share of invested capital). Symmetric with the reproduction barrier.
    CAP_FULL = 0.12
    def _leg(growth, share):
        return (max(0.0, min(1.0, growth / CAP_FULL)) * min(1.0, share)) if growth is not None else 0.0
    cap_resp = min(1.0, _leg(sig.get("cap_growth"), sig.get("phys_share", 0.0))
                   + _leg(sig.get("intang_cap_growth"), sig.get("intang_share", 0.0)))
    eff_barrier = barrier * (1.0 - 0.6 * cap_resp)
    # (Guard 3) gate by mean-reversion evidence: a business that has never
    # reverted is treated as structural growth, not a fadeable cyclical rent.
    reversion = sig.get("mean_reversion", 1.0)
    volume_faded = rev_excess * (1.0 - eff_barrier) * scarcity * reversion

    # --- Capital-cycle reproduction-equilibrium margin rent ---
    # A physical-bottleneck name in an active build-out fades the margin ABOVE its
    # long-run reproduction equilibrium, over the gestation period, REGARDLESS of
    # intangible durability (the rent is a physical shortage, not a moat). This
    # catches memory/semicap and — via the explicit AI tag — fabless rent-holders
    # (NVIDIA) whose upstream bottleneck is invisible on their own balance sheet.
    ticker = fin.get(aip.FIELDS["ticker"])
    industry = fin.get(aip.FIELDS["industry"]) or sig.get("industry")
    phase, gestation, in_buildout, cyc_src = capitalcycle.classify(industry, cycle_map, ticker)
    margin_rent = 0.0
    if phase != "intangible" and in_buildout > 0:
        # Full rent to the long-run (mid-cycle, intangible-inclusive) equilibrium:
        # normalizes BOTH the margin AND the price level (revenue) toward the
        # median, so a price/margin-spiked name is not left at the scarcity level.
        # Applied only to a TAGGED live cycle or a DETECTED industry build-out (not
        # to every physical cyclical above its median — see CALIBRATION note below).
        mr = min(0.60, sig.get("cc_rent_frac", 0.0))
        gate = 1.0 if cyc_src.startswith("AI") else reversion
        margin_rent = mr * in_buildout * gate
    # CALIBRATION (open): extending the margin rent to ANY physical cyclical at a
    # margin peak (gated by reversion) catches capital-goods peaks the detector
    # misses (Mikron, IES, Powell) but, in a late-cycle year, fades ~1/4 of the
    # physical universe — a philosophy choice (targeted scarcity-rents vs broad
    # cyclical-margin normalization) left for explicit decision.

    # combine the volume rent and the margin rent into one normalization
    faded_frac = 1.0 - (1.0 - volume_faded) * (1.0 - margin_rent)
    if faded_frac <= 1e-4:
        nm0, op0 = normalized_ev_ebit(base.get("ev"), cur_noi, 0.0, 0.0, base["wacc"],
                                      fin.get(aip.FIELDS["tax"]), phase)
        return {"er_current": er_current, "er_adj": er_current, "faded_frac": 0.0,
                "H": 0.0, "barrier": barrier, "rev_excess": rev_excess,
                "scarcity": scarcity, "cap_resp": cap_resp, "margin_rent": 0.0,
                "phase": phase, "cycle": cyc_src, "norm_ev_ebit": nm0, "overpriced": op0}

    ind = fin.get(aip.FIELDS["industry"])
    if margin_rent > 1e-4 and gestation:
        # The supply-side rent persists until BOTH the physical capacity AND the
        # intangible know-how (process IP / yield engineering / ecosystem) are
        # reproduced — and the intangible build is the slower, binding leg for a
        # know-how-intensive component. So a pure-physical gestation understates
        # the duration; extend it by the supply chain's intangible intensity.
        H = min(12.0, float(gestation) * (1.0 + 2.5 * sig.get("intang_share", 0.0)))
    else:
        H = max(2.0, min(15.0, supply_lag(ind) * (0.5 + barrier)))   # entry delay, barrier-scaled

    base_noi = cur_noi * (1.0 - faded_frac)
    f2 = dict(fin); f2[aip.FIELDS["nopat"]] = base_noi
    v2 = aip.value_and_return(f2, re=re, re2=None)
    if not v2 or not v2.get("op_value"):
        return {"er_current": er_current, "er_adj": er_current, "faded_frac": faded_frac,
                "H": H, "barrier": barrier, "rev_excess": rev_excess}
    V2 = v2["op_value"]; wacc = v2["wacc"]
    excess0 = cur_noi * faded_frac
    V1 = sum(excess0 * max(0.0, 1.0 - tt / H) / (1 + wacc) ** tt
             for tt in range(1, math.ceil(H) + 1))
    comb_noi = base_noi * ((V2 + V1) / V2)
    f3 = dict(fin); f3[aip.FIELDS["nopat"]] = comb_noi
    adj = aip.value_and_return(f3, re=re, re2=None)
    er_adj = adj["er1"] if adj else er_current
    nm_mult, overpriced = normalized_ev_ebit(base.get("ev"), cur_noi, faded_frac, H,
                                             base["wacc"], fin.get(aip.FIELDS["tax"]), phase)
    return {"er_current": er_current, "er_adj": er_adj, "faded_frac": faded_frac,
            "H": H, "barrier": barrier, "rev_excess": rev_excess, "scarcity": scarcity,
            "cap_resp": cap_resp, "margin_rent": margin_rent, "phase": phase,
            "cycle": cyc_src, "norm_ev_ebit": nm_mult, "overpriced": overpriced,
            "er_s2": v2["er1"]}   # full-revert (no Stage-1 credit)


def run(tickers, panel_path, fins_by_ticker, moats_by_ticker=None,
        sheet=None, re=0.07, re2=0.12):
    """Run the over-earning two-stage on a list of tickers.

    fins_by_ticker: {ticker: fin dict with aip screener field names}
    moats_by_ticker: optional {ticker: researched moat} overriding fin's Moat Score
    Returns a list of result dicts (one per ticker, in input order)."""
    panel, idx = history.load_panel(panel_path, sheet)
    out = []
    for t in tickers:
        fin = dict(fins_by_ticker.get(t, {}))
        if moats_by_ticker and moats_by_ticker.get(t) is not None:
            fin[aip.FIELDS["moat"]] = moats_by_ticker[t]
        rows = panel.get(t)
        rec = {"ticker": t, "company": fin.get(aip.FIELDS["name"])}
        if not rows or not fin:
            rec["error"] = "no panel/fin"; out.append(rec); continue
        sig = panel_signals(rows, idx)
        ts = two_stage_return(fin, sig, re=re, re2=re2)
        if ts:
            rec.update(ts); rec["roic_star"] = sig.get("roic_star")
            rec["industry"] = fin.get(aip.FIELDS["industry"])
        out.append(rec)
    return out
