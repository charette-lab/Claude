#!/usr/bin/env python3
"""AIP Value — ROIIC persistence-fade DCF, run against an attached Excel file.

Economic premise: growth requires capital (g = ROIIC x RR), and a company's
*excess* return (ROIIC - WACC) does not persist — it decays geometrically toward
the cost of capital at a speed set by the durability of the moat. This directly
handles a cyclically-inflated starting ROIIC (e.g. a semiconductor firm posting
60% ROIIC at a demand peak): the surge is mean-reverted away over the empirical
Competitive Advantage Period (CAP) rather than held flat.

Fade engine (per year t, over the CAP of N years):
    ROIIC_t   = WACC + (ROIIC_0 - WACC) * persistence**t      (excess decays)
    g_t       = ROIIC_t * RR                                  (RR held at RR_0)
    NOPAT_t   = NOPAT_(t-1) * (1 + g_t)
    FCF_t     = NOPAT_t * (1 - RR)
    PV        = sum_t FCF_t / (1+WACC)**t
Terminal (after the CAP, ROIIC has reached WACC so growth is value-neutral):
    TV        = NOPAT_N * (1 + g_term) / WACC ;  PV_TV = TV / (1+WACC)**N

Empirical CAP durations & persistence, mapped from the Moat Score:
    Moat > 7.5   Superior / Wide moat  -> CAP 10-20y, persistence 0.85-0.95
    6.0 - 7.5    Narrow / Standard     -> CAP  5-10y, persistence 0.70-0.80
    Moat < 6.0   No moat / cyclical    -> CAP  1-5y,  persistence 0.50-0.60
(interpolated within each band by score; override with --cap / --persistence).

NOTE on notation: the persistence factor is written as 'r' in the source
framework, but here 'r' is the discount rate / WACC. The persistence factor is
called 'persistence' (phi) to avoid the clash.

DEFAULT EXCEL COLUMN MAPPING (override with --col-* flags):
  NOPAT_0  <- "New Operating Income"   ROIIC_0 <- "ROICm 7"   RR <- "RR 7"
  moat     <- "Moat Score"
  (comparison) "EV", "Market Cap", "Net debt",
               "Shares used to calculate Diluted EPS - Total", "Close Price"

DEFAULT PARAMETERS (override with flags):
  r (WACC) = 0.12     g_term = 0.025 (< r)     horizon = 5 (for the IRR)

Usage:
  python3 roiic_dcf.py <file.xlsx> "<company>" [options]
  python3 roiic_dcf.py <file.xlsx> --list
Options:
  --r FLOAT --gterm FLOAT --cap INT --persistence FLOAT --horizon INT
  --payout-total FLOAT --sheet NAME --col-* STR
"""

import argparse
import sys

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required:  pip install openpyxl")


def moat_to_cap_persistence(score):
    """Map a Moat Score to (CAP years, persistence factor, tier label).

    Empirical CAP durations:
      score > 7.5      Superior/Wide  -> CAP 10-20y, persistence 0.85-0.95
      6.0 <= s <= 7.5  Narrow/Std     -> CAP  5-10y, persistence 0.70-0.80
      score < 6.0      No moat/cyclic -> CAP  1-5y,  persistence 0.50-0.60
    Values are interpolated linearly within each band. Returns None if missing.
    """
    if score is None:
        return None
    if score > 7.5:                                  # Superior / Wide
        f = min(max((score - 7.5) / (9.0 - 7.5), 0.0), 1.0)
        cap = 10 + f * (20 - 10)
        phi = 0.85 + f * (0.95 - 0.85)
        tier = "Wide"
    elif score >= 6.0:                               # Narrow / Standard
        f = (score - 6.0) / (7.5 - 6.0)
        cap = 5 + f * (10 - 5)
        phi = 0.70 + f * (0.80 - 0.70)
        tier = "Narrow"
    else:                                            # No moat / cyclical
        f = max(0.0, score) / 6.0
        cap = 1 + f * (5 - 1)
        phi = 0.50 + f * (0.60 - 0.50)
        tier = "Cyclical"
    return max(1, int(round(cap))), phi, tier


def find_columns(ws, wanted):
    """Map each wanted header string to its 1-based column index (exact, then
    case-insensitive, then prefix match on the first row)."""
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None:
            headers[str(v).strip()] = c
    lower = {k.lower(): idx for k, idx in headers.items()}
    out = {}
    for key, label in wanted.items():
        if label in headers:
            out[key] = headers[label]
        elif label.lower() in lower:
            out[key] = lower[label.lower()]
        else:
            match = next((idx for h, idx in lower.items()
                          if h.startswith(label.lower())), None)
            out[key] = match
    return out


def find_company_row(ws, name_col, query):
    q = query.strip().lower()
    matches = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=name_col).value
        if v and q in str(v).lower():
            matches.append((r, str(v)))
    return matches


def num(ws, row, col):
    if col is None:
        return None
    v = ws.cell(row=row, column=col).value
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def money(x):
    if x is None:
        return "n/a"
    a = abs(x)
    if a >= 1e9:
        return f"{x/1e9:,.2f} bn"
    if a >= 1e6:
        return f"{x/1e6:,.1f} mn"
    return f"{x:,.0f}"


def pct(x):
    return "n/a" if x is None else f"{x*100:.2f}%"


def value_company(nopat0, roiic0, rr0, r, g_term, cap, phi):
    """Run the persistence-fade DCF. Returns a dict with the schedule, PVs,
    total operating value, and a per-year free-cash-flow function."""
    excess0 = roiic0 - r
    nopat_path = [nopat0]          # index t -> NOPAT at end of year t (0 = today)
    fcf_path = [None]
    sched = []                     # (t, roiic_t, g_t)
    pv_explicit = 0.0
    for t in range(1, cap + 1):
        roiic_t = r + excess0 * (phi ** t)
        g_t = roiic_t * rr0
        nopat_t = nopat_path[t - 1] * (1 + g_t)
        fcf_t = nopat_t * (1 - rr0)
        nopat_path.append(nopat_t)
        fcf_path.append(fcf_t)
        sched.append((t, roiic_t, g_t))
        pv_explicit += fcf_t / (1 + r) ** t
    nopat_n = nopat_path[cap]
    tv = nopat_n * (1 + g_term) / r          # ROIIC_term = r -> value-neutral
    pv_tv = tv / (1 + r) ** cap
    total = pv_explicit + pv_tv

    def cf_for_year(t):
        if 1 <= t <= cap:
            return fcf_path[t]
        nopat_t = nopat_n * (1 + g_term) ** (t - cap)
        return nopat_t * (1 - g_term / r)

    return {"sched": sched, "pv_explicit": pv_explicit, "tv": tv, "pv_tv": pv_tv,
            "total": total, "nopat_n": nopat_n, "cf_for_year": cf_for_year}


def main():
    ap = argparse.ArgumentParser(add_help=True, description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("company", nargs="?", default=None)
    ap.add_argument("--list", action="store_true", help="list companies and exit")
    ap.add_argument("--r", type=float, default=0.12, help="WACC / discount rate")
    ap.add_argument("--gterm", type=float, default=0.025)
    ap.add_argument("--cap", type=int, default=None,
                    help="CAP length (years); if omitted, derived from the Moat Score")
    ap.add_argument("--persistence", type=float, default=None,
                    help="persistence factor phi (0-1); if omitted, from the Moat Score")
    ap.add_argument("--horizon", type=int, default=5,
                    help="holding period in years for the expected-return / IRR (default 5)")
    ap.add_argument("--payout-total", type=float, default=0.0,
                    help="cumulative dividends + net buybacks over the horizon, "
                         "subtracted from the cash sweep (default 0 = all FCF de-levers)")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--col-nopat", default="New Operating Income")
    ap.add_argument("--col-roiic", default="ROICm 7")
    ap.add_argument("--col-rr", default="RR 7")
    ap.add_argument("--col-name", default="Company Name")
    ap.add_argument("--col-moat", default="Moat Score")
    args = ap.parse_args()

    if args.gterm >= args.r:
        sys.exit(f"g_term ({args.gterm}) must be < r ({args.r}).")

    wb = openpyxl.load_workbook(args.file, data_only=True)
    ws = wb[args.sheet] if args.sheet else wb.worksheets[0]

    cols = find_columns(ws, {
        "name": args.col_name, "nopat": args.col_nopat,
        "roiic": args.col_roiic, "rr": args.col_rr, "moat": args.col_moat,
        "ev": "EV", "mktcap": "Market Cap", "netdebt": "Net debt",
        "shares": "Shares used to calculate Diluted EPS - Total",
        "price": "Close Price", "ticker": "Instrument",
    })
    if cols["name"] is None:
        sys.exit(f'Could not find a "{args.col_name}" column in row 1.')

    if args.list or not args.company:
        print(f"Companies in '{ws.title}':")
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=cols["name"]).value
            if v:
                print(f"  {v}")
        if not args.company and not args.list:
            print("\nPass a company name substring to value one.")
        return

    matches = find_company_row(ws, cols["name"], args.company)
    if not matches:
        sys.exit(f'No company matching "{args.company}". Use --list to see options.')
    if len(matches) > 1:
        names = ", ".join(n for _, n in matches)
        sys.exit(f'Ambiguous "{args.company}" matches: {names}. Be more specific.')
    row, company = matches[0]

    nopat0 = num(ws, row, cols["nopat"])
    roiic0 = num(ws, row, cols["roiic"])
    rr0 = num(ws, row, cols["rr"])
    if None in (nopat0, roiic0, rr0):
        sys.exit(f"Missing driver inputs for {company}: "
                 f"NOPAT={nopat0}, ROIIC={roiic0}, RR={rr0}")

    r, g_term = args.r, args.gterm

    # CAP + persistence from the Moat Score (overridable).
    moat = num(ws, row, cols["moat"])
    mp = moat_to_cap_persistence(moat)
    if mp is not None:
        d_cap, d_phi, tier = mp
        src = f"Moat {moat:.2f} [{tier}]"
    else:
        d_cap, d_phi, tier = 8, 0.75, "Narrow(default)"
        src = "no Moat Score; default Narrow"
    cap = args.cap if args.cap is not None else d_cap
    phi = args.persistence if args.persistence is not None else d_phi

    res = value_company(nopat0, roiic0, rr0, r, g_term, cap, phi)
    total = res["total"]

    ticker = ws.cell(row=row, column=cols["ticker"]).value if cols["ticker"] else ""
    print(f"\nAIP VALUE — ROIIC persistence-fade DCF — {company} {f'({ticker})' if ticker else ''}")
    print(f"WACC r = {pct(r)}   CAP = {cap}y   persistence phi = {phi:.2f}   "
          f"g_term = {pct(g_term)}   ({src})")
    print("=" * 70)
    print(f"  NOPAT_0 (New Operating Income) .... {money(nopat0)}")
    print(f"  ROIIC_0 (ROICm 7, starting) ....... {pct(roiic0)}   excess over WACC {pct(roiic0 - r)}")
    print(f"  RR (RR 7, held constant) .......... {pct(rr0)}")

    print(f"\n  FADE SCHEDULE (ROIIC -> WACC at phi={phi:.2f} per year)")
    sched = res["sched"]
    marks = sorted(set([1, max(1, cap // 2), cap]))
    for t, roiic_t, g_t in sched:
        if t in marks:
            print(f"    year {t:>2}:  ROIIC {pct(roiic_t):>8}   g {pct(g_t):>8}")

    print(f"\n  PV(explicit FCF, yrs 1-{cap}) ...... {money(res['pv_explicit'])}")
    print(f"  Terminal value (at yr {cap}) ........ {money(res['tv'])}")
    print(f"  PV(terminal) ...................... {money(res['pv_tv'])}")
    print("  " + "-" * 50)
    print(f"  TOTAL OPERATING VALUE ............. {money(total)}")
    if total:
        print(f"    explicit {res['pv_explicit']/total*100:5.1f}%   "
              f"terminal {res['pv_tv']/total*100:5.1f}%")

    # --- Market comparison ---
    ev = num(ws, row, cols["ev"])
    netdebt = num(ws, row, cols["netdebt"])
    shares = num(ws, row, cols["shares"])
    price = num(ws, row, cols["price"])
    mktcap = num(ws, row, cols["mktcap"])
    print("\n  MARKET COMPARISON")
    if ev is not None:
        upside = total / ev - 1 if ev else None
        print(f"    Enterprise Value (market) ..... {money(ev)}")
        if ev:
            print(f"    Model / EV .................... {total/ev:.2f}x  "
                  f"({'+' if upside>=0 else ''}{upside*100:.1f}% vs EV)")
    if netdebt is not None:
        equity = total - netdebt
        print(f"    Net debt (neg = net cash) ..... {money(netdebt)}")
        print(f"    Implied equity value .......... {money(equity)}")
        if shares:
            iv_ps = equity / shares
            line = f"    Implied value / share ......... {iv_ps:,.2f}"
            if price:
                line += f"   (market {price:,.2f}, {'+' if iv_ps>=price else ''}{(iv_ps/price-1)*100:.1f}%)"
            print(line)
    if mktcap is not None:
        print(f"    Market cap .................... {money(mktcap)}")

    # --- Expected return (IRR) over the holding horizon ---
    cf_for_year = res["cf_for_year"]
    n = args.horizon
    print(f"\n  EXPECTED RETURN  (horizon n = {n}y)")
    cf_sum = sum(cf_for_year(t) for t in range(1, n + 1))
    print(f"    EV_target (operating value) ... {money(total)}")
    print(f"    Sum FCF yrs 1..{n} ............. {money(cf_sum)}")
    if args.payout_total:
        print(f"    - Dividends/buybacks (horizon)  {money(args.payout_total)}")
    if netdebt is not None:
        nd5 = netdebt - (cf_sum - args.payout_total)
        eqv_target = total - nd5
        print(f"    ND_0 (current net debt) ....... {money(netdebt)}")
        print(f"    ND_{n} (after cash sweep) ....... {money(nd5)}")
        print(f"    EqV_target = EV_target - ND_{n} . {money(eqv_target)}")
        if mktcap and mktcap > 0 and eqv_target > 0:
            eq_irr = (eqv_target / mktcap) ** (1 / n) - 1
            print(f"    EqV_0 (current market cap) .... {money(mktcap)}")
            print(f"    >> Expected equity return (IRR) {pct(eq_irr)} / yr")
        elif eqv_target <= 0:
            print("    >> Projected equity value <= 0; equity IRR undefined.")
    if ev is not None and ev > 0 and total > 0:
        unlev = (total / ev) ** (1 / n) - 1
        print(f"    >> Unlevered return (EV_target/EV_0) {pct(unlev)} / yr")
        print("    (If equity IRR >> unlevered, the thesis leans on leverage.)")

    print("\n  Analytical framework output, not investment advice.")


if __name__ == "__main__":
    main()
