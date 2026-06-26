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


def _col(rows, idx, name):
    i = idx.get(name)
    return [_num(r[i]) if (i is not None and i < len(r)) else None for r in rows]


def _last(xs):
    return next((v for v in reversed(xs) if v is not None), None)


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

    # --- revenue spike: current vs its own macro-shock-weighted log-trend ---
    macro_w = _col(rows, idx, "Macro Weight")
    trend = _loglinear_current(sales, macro_w)
    rev_excess_frac = 0.0
    if trend and cur_sales and cur_sales > 0:
        rev_excess_frac = max(0.0, (cur_sales - trend) / cur_sales)
    past_peak = bool(cur_sales is not None and sales
                     and cur_sales < max(v for v in sales if v is not None))
    if past_peak:
        rev_excess_frac = 0.0  # already correcting; no live spike

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
    return {
        "rev_excess_frac": rev_excess_frac, "past_peak": past_peak,
        "roic_star": roic_star, "asset_sweat": asset_sweat,
        "repro_prem": repro_prem, "phys_share": phys_share,
        "intang_share": intang_share, "romic": romic,
        "margin_elev": margin_elev, "cap_growth": cap_growth,
        "intang_cap_growth": intang_cap_growth,
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
    # reproduction barrier = hard-to-reproduce capital / total capital. Two legs,
    # symmetric: the PHYSICAL premium over book (reproduction-cost dislocation,
    # credited to the physical share) PLUS the intangible capital base (CUDA/EUV/
    # brand/IP — the asset-light moat, the same stock that sits in invested
    # capital). This makes the barrier consistent across asset-heavy and
    # asset-light moats instead of only crediting physical capacity.
    repro = (max(0.0, sig.get("repro_prem", 0.0)) * sig.get("phys_share", 0.0)
             + sig.get("intang_share", 0.0))
    repro = max(0.0, min(1.0, repro))
    barrier = 0.45 * q + 0.40 * m + 0.15 * repro
    return max(0.0, min(1.0, barrier))


def two_stage_return(fin, sig, re=0.07, re2=0.12):
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
    faded_frac = rev_excess * (1.0 - eff_barrier) * scarcity   # transient supply-erodable rent
    if faded_frac <= 1e-4:
        return {"er_current": er_current, "er_adj": er_current, "faded_frac": 0.0,
                "H": 0.0, "barrier": barrier, "rev_excess": rev_excess,
                "scarcity": scarcity, "cap_resp": cap_resp}

    ind = fin.get(aip.FIELDS["industry"])
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
    return {"er_current": er_current, "er_adj": er_adj, "faded_frac": faded_frac,
            "H": H, "barrier": barrier, "rev_excess": rev_excess, "scarcity": scarcity,
            "cap_resp": cap_resp, "er_s2": v2["er1"]}   # full-revert (no Stage-1 credit)


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
