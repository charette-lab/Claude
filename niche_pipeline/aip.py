"""
aip.py — thin wrapper over the proven aip-value engine (roiic_dcf.py).

Reuses the firm's ROIIC persistence-fade DCF rather than re-deriving it:
  NOPAT_0 <- "New Operating Income" ; ROIIC_0 <- "ROICm 7" ; RR_0 <- "RR 7"
  per-company WACC from an equity hurdle (--re), optional lever-glide,
  Moat Score -> CAP / persistence phi, sector sales-growth floor,
  terminal at competitive equilibrium.

Exposes two functions used by the pipeline:
  value_and_return(fin, re, re2, ...)  -> dict(wacc, rating, op_value, er1, er2, ...)
  implied_moat(fin, r, ...)            -> str like "27y"
where `fin` is a dict of the company's screener fields.
"""
from __future__ import annotations
import importlib.util
import os
import sys

_ENGINE = None


def engine():
    """Locate and import roiic_dcf.py once."""
    global _ENGINE
    if _ENGINE is not None:
        return _ENGINE
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.environ.get("AIP_VALUE_ENGINE", ""),
        os.path.join(here, "roiic_dcf.py"),
        os.path.expanduser("~/.claude/skills/aip-value/roiic_dcf.py"),
        os.path.join(here, "..", "skills", "aip-value", "roiic_dcf.py"),
    ]
    for path in candidates:
        if path and os.path.exists(path):
            spec = importlib.util.spec_from_file_location("aipval", path)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            _ENGINE = mod
            return mod
    sys.exit("aip.py: could not find roiic_dcf.py (the aip-value engine). "
             "Set AIP_VALUE_ENGINE=/path/to/roiic_dcf.py")


# Screener field names the engine expects.
FIELDS = {
    "name": "Company Name", "nopat": "New Operating Income", "roiic": "ROICm 7",
    "rr": "RR 7", "moat": "Moat Score", "industry": "GICS Industry Group Name",
    "country": "Country of Headquarters", "gross": "Gross debt",
    "tax": "Income Tax Rate - Instrument", "mktcap": "Market Cap",
    "netdebt": "Net debt", "ticker": "Instrument", "sales": "Sales",
}


def _discount_rate(m, re, nopat0, mktcap, netdebt, gross, tax, country, ind,
                   lever_glide, gterm, country_base=None, country_crp=None):
    if re is None:
        return None, None, None, None
    cbase, _ = m.currency_base(country, m.parse_kv_rates(country_base))
    t = tax if (tax is not None and tax < 1) else 0.25
    rd, rating, _ = m.synthetic_rd(nopat0 / (1 - t), gross, mktcap, cbase)
    crp = m.country_risk_premium(country, m.parse_kv_rates(country_crp))
    if lever_glide:
        L = m.target_leverage(ind, None)
        rd_m = m.mature_cost_of_debt(L, cbase, mktcap)[0]
        w0 = min(max(m.firm_wacc_taxed(re, rd, mktcap, netdebt, t) + crp, 0.04), re + crp)
        wm = min(max(m.mature_wacc(re, rd_m, L, t)[0] + crp, 0.04), 0.25)
        return w0, rd, rating, (w0, wm)
    return min(max(m.firm_wacc(re, rd, mktcap, netdebt) + crp, 0.04), 0.25), rd, rating, None


def _glide_path(m, glide, n1, n2):
    return (m.wacc_glide(glide[0], glide[1], n1, n2), glide[1]) if glide else (None, None)


def value_and_return(fin, re=0.07, re2=0.12, lever_glide=True, gterm=0.025,
                     horizon=5, sales_floor=True, country_base=None, country_crp=None,
                     moat_override=None):
    """Run the DCF for one company. `fin` keys are the screener field names.

    moat_override (0-10) replaces the file's Moat Score for the competitive-
    advantage period only — used to value the durable CORE on its own moat
    (a longer CAP) rather than the blended group, i.e. the separated re-rating.
    """
    m = engine()
    nopat0 = fin.get(FIELDS["nopat"]); roiic0 = fin.get(FIELDS["roiic"])
    rr0 = fin.get(FIELDS["rr"]); mktcap = fin.get(FIELDS["mktcap"])
    netdebt = fin.get(FIELDS["netdebt"]) or 0.0
    if None in (nopat0, roiic0, rr0) or not mktcap:
        return None
    gross = fin.get(FIELDS["gross"]); tax = fin.get(FIELDS["tax"])
    moat = moat_override if moat_override is not None else fin.get(FIELDS["moat"])
    country = fin.get(FIELDS["country"])
    ind = " ".join(str(fin.get(FIELDS["industry"]) or "").split())
    mp = m.moat_to_cap_persistence(moat); phi = mp[1] if mp else 0.75
    life = m.moat_to_life(moat) or 15
    base = m.base_rate_for(ind, None)[0]
    sales0 = fin.get(FIELDS["sales"]) if sales_floor else None
    floor = sales_floor

    def exp_ret(r, glide):
        n1 = min(3, max(1, life - 1)); n2 = max(1, life - n1)
        wp, wt = _glide_path(m, glide, n1, n2)
        res = m.value_company(nopat0, roiic0, rr0, r, gterm, n1, n2, phi, base,
                              sales0=sales0, gics_industry=ind, sales_floor=floor,
                              wacc_path=wp, wacc_terminal=wt)
        L = m.target_leverage(ind, None) if glide else None
        t = tax if (tax is not None and tax < 1) else 0.25
        val, method = m.equity_return(res, mktcap, netdebt, horizon, tax=t, lever_L=L)[:2]
        return val, method, res.get("op_value", res.get("total"))

    r1, rd, rating, glide1 = _discount_rate(m, re, nopat0, mktcap, netdebt, gross,
                                            tax, country, ind, lever_glide, gterm,
                                            country_base, country_crp)
    r1 = r1 if re is not None else 0.12
    er1, mth1, opval = exp_ret(r1, glide1)
    out = {"wacc": r1, "rd": rd, "rating": rating, "op_value": opval, "ev": mktcap + netdebt,
           "er1": er1, "er1_method": mth1, "moat": moat, "mktcap": mktcap, "netdebt": netdebt,
           "is_financial": ind in m.FINANCIAL_SECTORS}
    if re2 is not None:
        r2, _, _, glide2 = _discount_rate(m, re2, nopat0, mktcap, netdebt, gross,
                                          tax, country, ind, lever_glide, gterm,
                                          country_base, country_crp)
        out["er2"], out["er2_method"], _ = exp_ret(r2, glide2)
    return out


def warranted_life(moat):
    """The competitive-advantage period the engine assigns to a moat score —
    the benchmark the implied (price-derived) moat length is compared against."""
    return engine().moat_to_life(moat) or 15


def implied_moat(fin, r=0.12, lever_glide=True, gterm=0.025, max_n2=150,
                 country_base=None, country_crp=None, moat_override=None):
    """Reverse DCF: solve total moat life so model op-value == market EV at r."""
    m = engine()
    nopat0 = fin.get(FIELDS["nopat"]); roiic0 = fin.get(FIELDS["roiic"])
    rr0 = fin.get(FIELDS["rr"]); mktcap = fin.get(FIELDS["mktcap"])
    netdebt = fin.get(FIELDS["netdebt"]) or 0.0
    if None in (nopat0, roiic0, rr0) or not mktcap:
        return None
    gross = fin.get(FIELDS["gross"]); tax = fin.get(FIELDS["tax"])
    moat = moat_override if moat_override is not None else fin.get(FIELDS["moat"])
    country = fin.get(FIELDS["country"])
    ind = " ".join(str(fin.get(FIELDS["industry"]) or "").split())
    mp = m.moat_to_cap_persistence(moat); phi = mp[1] if mp else 0.75
    base = m.base_rate_for(ind, None)[0]
    sales0 = fin.get(FIELDS["sales"])
    _, _, _, glide = _discount_rate(m, r, nopat0, mktcap, netdebt, gross, tax,
                                    country, ind, lever_glide, gterm,
                                    country_base, country_crp)
    target = mktcap + netdebt

    def tot(n2):
        wp, wt = _glide_path(m, glide, 3, n2)
        return m.value_company(nopat0, roiic0, rr0, r, gterm, 3, n2, phi, base,
                               sales0=sales0, gics_industry=ind, sales_floor=True,
                               wacc_path=wp, wacc_terminal=wt)["total"]
    t0 = tot(0)
    if t0 >= target:
        return "<=3y"
    prev = t0
    for n2 in range(1, max_n2 + 1):
        tt = tot(n2)
        if tt >= target:
            return f"{((n2 - 1) + (target - prev) / (tt - prev) + 3):.0f}y"
        prev = tt
    return ">150y"
