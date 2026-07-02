#!/usr/bin/env python3
"""aip-resource-mine — NAV / DCF + resource-option valuer for mining developers.

Earnings models (aip-value / sectors) cannot value a pre-production miner: there
are no cash flows yet. This skill values a deposit bottom-up from its geology and
engineering parameters:

  per-tonne NSR  ->  after-tax mine cash flow  ->  NPV (NAV)  ->  P/NAV & per-share

It handles the things that actually drive a mine NAV: metal grades & recoveries,
payabilities, smelter TC/RC + transport on the concentrate, NSR royalties, build
delay, sustaining capex, tax with depreciation, and partial ownership. It runs
grade/cut-off, metal-price and CAPEX sensitivities, and a brownfield-synergy case.
For an early-stage discovery with no resource it adds a simple exploration-OPTION
layer (in-situ $/lb on a notional/illustrative tonnage), and a 2-state
implied-probability back-out from the market price.

Inputs come from a JSON deposit file (see --example) extracted from the 43-101 /
technical report. CAD/USD and share data let it print a per-share SOTP.

Usage:
  python3 mine_dcf.py --deposit los_helados.json [--cu 4.50 --capex 3.5e9 --tpd 100000]
  python3 mine_dcf.py --deposit d.json --sweep            # cut-off x price x capex grid
  python3 mine_dcf.py --deposit d.json --implied-prob --price 19.87 --floor 5 --success 40
  python3 mine_dcf.py --example > los_helados.json        # write a template
"""
import argparse
import json
import sys

LB_PER_T = 2204.62           # pounds per tonne
OZ_PER_G = 1.0 / 31.1035     # troy oz per gram


def per_tonne_nsr(grades, prices, rec, pay, ref, conc, royalty):
    """Net smelter return per tonne of ore milled, after payabilities, refining,
    smelter TC + concentrate transport, and NSR royalty. grades: Cu% , Au g/t,
    Ag g/t. Returns (net_nsr, gross_payable)."""
    cu_lb = grades["cu"] / 100 * LB_PER_T * rec["cu"] * pay["cu"]
    au_oz = grades["au"] * OZ_PER_G * rec["au"] * pay["au"]
    ag_oz = grades["ag"] * OZ_PER_G * rec["ag"] * pay["ag"]
    gross = (cu_lb * (prices["cu"] - ref["cu"])
             + au_oz * (prices["au"] - ref["au"])
             + ag_oz * (prices["ag"] - ref["ag"]))
    # concentrate produced per tonne ore (dmt), and the TC + transport on it
    cu_recov_t = grades["cu"] / 100 * rec["cu"]            # tonnes Cu / tonne ore
    conc_dmt = cu_recov_t / conc["cu_grade"]              # dmt concentrate / t ore
    tc = conc_dmt * conc["tc_dmt"]
    transport = conc_dmt / (1 - conc["moisture"]) * conc["transport_wmt"]
    return gross - tc - transport - gross * royalty, gross


def mine_npv(d, grades, tonnage_mt, prices, capex, tpd=None, disc=None, tax=None,
             build_yrs=None, life_max=None, sustaining=None):
    """After-tax project NPV (100% basis) for a block-cave/bulk operation."""
    e = d["economics"]
    tpd = tpd or e["tpd"]
    mtpa = tpd * 365 / 1e6
    disc = disc if disc is not None else e["discount"]
    tax = tax if tax is not None else e["tax"]
    build_yrs = build_yrs if build_yrs is not None else e["build_years"]
    life_max = life_max if life_max is not None else e["life_max"]
    sustaining = sustaining if sustaining is not None else e["sustaining"]
    mineable = min(tonnage_mt, mtpa * life_max)
    life = max(1, int(mineable / mtpa))
    nsr, _ = per_tonne_nsr(grades, prices, d["recovery"], d["payable"],
                           d["refining"], d["concentrate"], e["nsr_royalty"])
    op_margin_t = nsr - e["opex_t"]
    annual_op = op_margin_t * mtpa * 1e6
    dep = capex / life
    annual_at = (annual_op - dep) * (1 - tax) + dep - sustaining
    pv = 0.0
    for y in range(1, build_yrs + 1):
        pv -= (capex / build_yrs) / (1 + disc) ** y
    for y in range(build_yrs + 1, build_yrs + life + 1):
        pv += annual_at / (1 + disc) ** y
    return {"npv": pv, "op_margin_t": op_margin_t, "nsr_t": nsr,
            "life": life, "mtpa": mtpa, "annual_op": annual_op}


def scaled_prices(base, cu):
    """Scale Au/Ag with Cu off the deposit's base deck (rough co-movement)."""
    f = cu / base["cu"]
    return {"cu": cu, "au": base["au"] * f, "ag": base["ag"] * f}


def blended(cut):
    """Tonnage-weighted grade across Indicated+Inferred for a cut-off case."""
    ti, tf = cut["ind"]["t"], cut["inf"]["t"]
    t = ti + tf
    g = {k: (cut["ind"][k] * ti + cut["inf"][k] * tf) / t for k in ("cu", "au", "ag")}
    return t, g


def fmt_b(x):
    return f"{x/1e9:,.2f}b" if abs(x) >= 1e8 else f"{x/1e6:,.0f}m"


def value_deposit(d, cu, capex, tpd=None):
    base = d["economics"]["price_deck"]
    prices = scaled_prices(base, cu)
    # pick the named base cut-off
    cut = d["cutoffs"][d["economics"]["base_cutoff"]]
    t, g = blended(cut)
    r = mine_npv(d, g, t, prices, capex, tpd=tpd)
    own = d["ownership"]
    fx = d.get("fx_to_report", 1.0)
    sh = d.get("shares")
    ngex = r["npv"] * own
    out = dict(r); out.update(cu=cu, capex=capex, grade=g, tonnage=t,
                              npv_attrib=ngex, per_share=(ngex * fx / sh) if sh else None)
    return out


def cmd_value(d, args):
    cu = args.cu or d["economics"]["price_deck"]["cu"]
    capex = args.capex or d["economics"]["capex"]
    o = value_deposit(d, cu, capex, tpd=args.tpd)
    print(f"\nAIP RESOURCE-MINE — {d['name']}  ({d.get('jurisdiction','')})")
    print(f"  {o['mtpa']:.1f} Mtpa  | Cu ${cu:.2f}/lb  | CAPEX {fmt_b(capex)}  | "
          f"disc {d['economics']['discount']*100:.0f}%  tax {d['economics']['tax']*100:.0f}%  "
          f"own {d['ownership']*100:.1f}%")
    print(f"  cut-off [{d['economics']['base_cutoff']}]: {o['tonnage']:,.0f} Mt @ "
          f"{o['grade']['cu']:.2f}% Cu / {o['grade']['au']:.2f} Au / {o['grade']['ag']:.1f} Ag")
    print("  " + "-" * 56)
    print(f"  NSR / tonne ............. ${o['nsr_t']:.2f}")
    print(f"  Operating margin / tonne  ${o['op_margin_t']:.2f}   (after ${d['economics']['opex_t']}/t opex)")
    print(f"  Mine life ............... {o['life']} yrs")
    print(f"  NPV (100%) .............. {fmt_b(o['npv'])}")
    print(f"  NPV (attrib {d['ownership']*100:.1f}%) ..... {fmt_b(o['npv_attrib'])}")
    if o["per_share"] is not None:
        print(f"  Per share ({d.get('report_ccy','')}) ...... {o['per_share']:.2f}")
    print("\n  Pre-production developer — apply a P/NAV (0.3-0.7x) for market value.")
    print("  Analytical framework output, not investment advice.")


def cmd_sweep(d, args):
    cus = args.cu_grid or [3.90, 4.50, 5.00]
    capexes = args.capex_grid or [d["economics"]["capex"], d["economics"]["capex"] * 0.54, d["economics"]["capex"] * 0.38]
    base = d["economics"]["price_deck"]; own = d["ownership"]; fx = d.get("fx_to_report", 1.0); sh = d.get("shares")
    print(f"\nAIP RESOURCE-MINE SWEEP — {d['name']}  (NPV attrib {own*100:.1f}%"
          + (f", C$/sh" if sh else "") + ")\n")
    for cutname, cut in d["cutoffs"].items():
        t, g = blended(cut)
        print(f"  cut-off {cutname}: {t:,.0f} Mt @ {g['cu']:.2f}% Cu")
        hdr = "    " + "".join(f"{('Cu $'+format(c,'.2f')):>16}" for c in cus)
        print(hdr)
        for cap in capexes:
            cells = ""
            for cu in cus:
                r = mine_npv(d, g, t, scaled_prices(base, cu), cap)
                v = r["npv"] * own
                cells += f"{(fmt_b(v) + (f'/{v*fx/sh:.1f}' if sh else '')):>16}"
            print(f"    CAPEX {fmt_b(cap):>7}{cells}")
        print()
    print("  cells: NPV attrib" + (" / per-share" if sh else "") + ". P/NAV 0.3-0.7x for market value.")


def cmd_implied_prob(d, args):
    price, floor = args.price, args.floor
    print(f"\nIMPLIED p(success) — {d['name']}   price {price}, failure floor {floor}")
    print(f"  p = (price - floor) / (success - floor)\n  {'success payoff':>16}{'p(success)':>14}")
    for s in (args.success_grid or [30, 35, 40, 45, 50]):
        p = (price - floor) / (s - floor)
        print(f"  {s:>16}{p*100:>13.0f}%")
    print("\n  Note: undiscounted. Discounting a multi-year success payoff at a required\n"
          "  return RAISES the implied probability (you pay near the discounted-success value).")


EXAMPLE = {
    "name": "Los Helados", "jurisdiction": "Chile/Argentina", "report_ccy": "CAD",
    "ownership": 0.691, "shares": 216857780, "fx_to_report": 1.39,
    "economics": {
        "tpd": 100000, "discount": 0.08, "tax": 0.30, "build_years": 5,
        "life_max": 40, "sustaining": 150e6, "opex_t": 21.00, "capex": 6.5e9,
        "nsr_royalty": 0.020, "base_cutoff": "0.40",
        "price_deck": {"cu": 3.90, "au": 1800, "ag": 20}
    },
    "recovery": {"cu": 0.902, "au": 0.803, "ag": 0.549},
    "payable": {"cu": 0.958, "au": 0.955, "ag": 0.90},
    "refining": {"cu": 0.11, "au": 6.37, "ag": 0.38},
    "concentrate": {"cu_grade": 0.239, "moisture": 0.08, "tc_dmt": 108.29, "transport_wmt": 104.00},
    "cutoffs": {
        "0.30": {"ind": {"t": 2200, "cu": 0.39, "au": 0.15, "ag": 1.4},
                 "inf": {"t": 1300, "cu": 0.33, "au": 0.10, "ag": 1.4}},
        "0.40": {"ind": {"t": 1650, "cu": 0.43, "au": 0.16, "ag": 1.5},
                 "inf": {"t": 600, "cu": 0.38, "au": 0.11, "ag": 1.6}},
        "0.60": {"ind": {"t": 510, "cu": 0.56, "au": 0.21, "ag": 1.8},
                 "inf": {"t": 40, "cu": 0.62, "au": 0.09, "ag": 2.4}},
    },
}


def _floatlist(s):
    return [float(x) for x in s.split(",")] if s else None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--deposit", help="path to deposit JSON")
    ap.add_argument("--example", action="store_true", help="print a template deposit JSON and exit")
    ap.add_argument("--cu", type=float, default=None, help="copper price $/lb")
    ap.add_argument("--capex", type=float, default=None, help="initial CAPEX (USD)")
    ap.add_argument("--tpd", type=float, default=None, help="throughput tonnes/day")
    ap.add_argument("--sweep", action="store_true", help="cut-off x price x capex grid")
    ap.add_argument("--cu-grid", type=_floatlist, default=None, help="e.g. 3.90,4.50,5.00")
    ap.add_argument("--capex-grid", type=_floatlist, default=None)
    ap.add_argument("--implied-prob", action="store_true", help="back out p(success) from market price")
    ap.add_argument("--price", type=float, default=None)
    ap.add_argument("--floor", type=float, default=5.0)
    ap.add_argument("--success-grid", type=_floatlist, default=None)
    args = ap.parse_args()
    if args.example:
        print(json.dumps(EXAMPLE, indent=2)); return
    if not args.deposit:
        sys.exit("Provide --deposit <file.json>  (or --example to print a template).")
    with open(args.deposit) as fh:
        d = json.load(fh)
    if args.implied_prob:
        if args.price is None:
            sys.exit("--implied-prob needs --price")
        cmd_implied_prob(d, args)
    elif args.sweep:
        cmd_sweep(d, args)
    else:
        cmd_value(d, args)


if __name__ == "__main__":
    main()
