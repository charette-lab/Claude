"""UK INVESTMENT-TRUST pitch — dark theme to match the 'Engaged Ownership as an
Asset Class' house style (navy background, bright-blue accents, geometric font,
big stats, cards), 16:9. Blends the stronger asset-class framings from that deck
with the trust-specific content. Launch specifics flagged indicative/proposed.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import openpyxl

# ---- dark palette / fonts -------------------------------------------------
BG      = RGBColor(0x0E, 0x1A, 0x2B)   # navy background
BLUE    = RGBColor(0x3E, 0x9C, 0xFF)   # bright accent
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
BODY    = RGBColor(0xC2, 0xCC, 0xD8)   # light-grey body
MUTED   = RGBColor(0x8A, 0x97, 0xA6)   # subtitle / secondary
CARD    = RGBColor(0x16, 0x24, 0x39)   # card fill
CARDLN  = RGBColor(0x2A, 0x3B, 0x54)   # card / rule border
PANEL   = RGBColor(0x14, 0x22, 0x36)
HEAD = BODY_F = "Poppins"              # geometric; renders if installed
BGIMG = "assets/bg_dark.png"
LOGO_WHITE = "assets/logo_white.png"
_LW_AR = (lambda s: s[0] / s[1])(Image.open(LOGO_WHITE).size)

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
_state = {"n": 0}


def rrect(s, x, y, w, h, fill=None, line=None, lw=1.0, rad=0.06, rounded=True):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    if rounded:
        try:
            shp.adjustments[0] = rad
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(lw)
    return shp


def tbox(s, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(x, y, w, h); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


def para(tf, text, size, color, bold=False, first=False, align=PP_ALIGN.LEFT,
         after=8, lead=1.12, font=BODY_F, italic=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align; p.space_after = Pt(after); p.line_spacing = lead
    r = p.add_run(); r.text = text; f = r.font
    f.size = Pt(size); f.bold = bold; f.italic = italic
    f.color.rgb = color; f.name = font
    return p


def runs(p, parts, size, after=8, lead=1.15):
    p.space_after = Pt(after); p.line_spacing = lead
    for txt, col, bold in parts:
        r = p.add_run(); r.text = txt
        r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = col; r.font.name = BODY_F


def newslide():
    s = prs.slides.add_slide(BLANK)
    s.shapes.add_picture(BGIMG, 0, 0, SW, SH)
    return s


def footer(s):
    _state["n"] += 1
    rrect(s, Inches(0.6), Inches(7.04), Inches(12.13), Pt(0.8), fill=CARDLN, rounded=False)
    para(tbox(s, Inches(0.6), Inches(7.1), Inches(9), Inches(0.3)),
         "Athanase  ·  strictly confidential  ·  professional & intermediary use only",
         8, MUTED, first=True, after=0)
    para(tbox(s, Inches(11.8), Inches(7.1), Inches(0.95), Inches(0.3)),
         str(_state["n"]), 8, MUTED, first=True, after=0, align=PP_ALIGN.RIGHT)


def content(num, section, title, subtitle=None):
    s = newslide()
    t = tbox(s, Inches(0.6), Inches(0.45), Inches(12.1), Inches(0.95))
    p = t.paragraphs[0]; p.line_spacing = 1.0
    r0 = p.add_run(); r0.text = f"{num}. "
    r0.font.size = Pt(30); r0.font.bold = True; r0.font.color.rgb = BLUE; r0.font.name = HEAD
    r1 = p.add_run(); r1.text = title
    r1.font.size = Pt(30); r1.font.bold = True; r1.font.color.rgb = WHITE; r1.font.name = HEAD
    para(tbox(s, Inches(0.62), Inches(0.18), Inches(11), Inches(0.3)),
         section.upper(), 10.5, BLUE, first=True, after=0)
    body_top = Inches(1.55)
    if subtitle:
        para(tbox(s, Inches(0.62), Inches(1.42), Inches(11.9), Inches(0.7)),
             subtitle, 13.5, MUTED, first=True, after=0, lead=1.22)
        body_top = Inches(2.35)
    footer(s)
    return s, body_top


def bullets(s, items, top, x=Inches(0.62), w=Inches(5.9), size=13, gap=12, arrow=True):
    tf = tbox(s, x, top, w, Inches(4.6))
    for i, it in enumerate(items):
        lead, b = it if isinstance(it, tuple) else ("", it)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.line_spacing = 1.22
        if arrow:
            r = p.add_run(); r.text = "→  "
            r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = BLUE; r.font.name = BODY_F
        if lead:
            r0 = p.add_run(); r0.text = lead + " "
            r0.font.size = Pt(size); r0.font.bold = True; r0.font.color.rgb = WHITE; r0.font.name = BODY_F
        r1 = p.add_run(); r1.text = b
        r1.font.size = Pt(size); r1.font.color.rgb = BODY; r1.font.name = BODY_F
    return tf


def card(s, x, y, w, h, title, body, tlead=None):
    rrect(s, x, y, w, h, fill=CARD, line=CARDLN, lw=1.0, rad=0.05)
    inx = Emu(int(x) + int(Inches(0.28)))
    para(tbox(s, inx, Emu(int(y) + int(Inches(0.26))), Emu(int(w) - int(Inches(0.56))),
              Inches(0.5)), title, 15, BLUE, first=True, bold=True, after=0)
    bt = tbox(s, inx, Emu(int(y) + int(Inches(0.95))), Emu(int(w) - int(Inches(0.56))),
              Emu(int(h) - int(Inches(1.1))))
    if tlead:
        para(bt, tlead, 11, WHITE, first=True, bold=True, after=10, lead=1.2)
        para(bt, body, 11.5, BODY, after=0, lead=1.3)
    else:
        para(bt, body, 12, BODY, first=True, after=0, lead=1.32)


def bigstats(s, stats, y, h=Inches(1.55)):
    rrect(s, Inches(0.62), y, Inches(12.1), h, fill=PANEL, line=CARDLN, lw=1.0, rad=0.04)
    n = len(stats); cw = Emu(int(Inches(12.1)) // n)
    for i, (big, lab) in enumerate(stats):
        cx = Emu(int(Inches(0.62)) + i * int(cw))
        para(tbox(s, cx, Emu(int(y) + int(Inches(0.22))), cw, Inches(0.7),
                  anchor=MSO_ANCHOR.MIDDLE), big, 38, BLUE, first=True, bold=True,
             after=0, align=PP_ALIGN.CENTER)
        para(tbox(s, Emu(int(cx) + int(Inches(0.2))), Emu(int(y) + int(Inches(0.92))),
                  Emu(int(cw) - int(Inches(0.4))), Inches(0.5)), lab, 11, BODY,
             first=True, after=0, align=PP_ALIGN.CENTER, lead=1.15)


def callout(s, text, y=Inches(6.05), h=Inches(0.85)):
    rrect(s, Inches(0.62), y, Inches(12.1), h, fill=CARD, line=CARDLN, lw=1.0, rad=0.06)
    rrect(s, Inches(0.62), y, Inches(0.07), h, fill=BLUE, rounded=False)
    para(tbox(s, Inches(0.95), y, Inches(11.5), h, anchor=MSO_ANCHOR.MIDDLE),
         text, 13, WHITE, first=True, after=0, lead=1.22)


def dtable(s, headers, rows, top, colx, hl=True, sizes=(13, 12.5), rh=Inches(0.62)):
    # headers in blue, blue rule, rows with thin separators; first col blue
    for j, htext in enumerate(headers):
        para(tbox(s, colx[j], top, Inches(4.5), Inches(0.4)), htext, sizes[0], BLUE,
             first=True, bold=True, after=0)
    ry = Emu(int(top) + int(Inches(0.5)))
    rrect(s, colx[0], ry, Emu(int(Inches(12.73)) - int(colx[0])), Pt(1.4), fill=BLUE, rounded=False)
    ry = Emu(int(ry) + int(Inches(0.12)))
    for row in rows:
        for j, cell in enumerate(row):
            bold = isinstance(cell, tuple)
            txt = cell[0] if bold else cell
            col = BLUE if j == 0 else (WHITE if bold else BODY)
            para(tbox(s, colx[j], ry, Emu((int(colx[j + 1]) if j + 1 < len(colx)
                      else int(Inches(12.8))) - int(colx[j]) - int(Inches(0.2))), rh,
                      anchor=MSO_ANCHOR.MIDDLE), txt, sizes[1], col, first=True,
                 bold=(j == 0 or bold), after=0, lead=1.1)
        ry = Emu(int(ry) + int(rh))
        rrect(s, colx[0], ry, Emu(int(Inches(12.73)) - int(colx[0])), Pt(0.8),
              fill=CARDLN, rounded=False)
    return ry


# ---- performance data + dark cumulative-growth chart ----------------------
import numpy as _np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as _plt


def _load_cmp():
    ws = openpyxl.load_workbook("data/Comparison_returns.xlsx", data_only=True)["Sheet1"]

    def grab(rows):
        out = []
        for c in range(2, ws.max_column + 1):
            for r in rows:
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)):
                    out.append(v)
        return out
    return grab(range(4, 16)), grab(range(22, 34))


_MSCI_X, _ATH_X = _load_cmp()


def _growth_chart(path):
    ga, gm = [1.0], [1.0]
    for r in _ATH_X:
        ga.append(ga[-1] * (1 + r))
    for r in _MSCI_X:
        gm.append(gm[-1] * (1 + r))
    x = _np.linspace(2006, 2026, len(ga))
    fig, ax = _plt.subplots(figsize=(7.4, 4.1))
    fig.patch.set_alpha(0); ax.set_facecolor("none")
    ax.plot(x, ga, color="#3E9CFF", lw=3.0, label="Athanase (net)", zorder=3)
    ax.plot(x, gm, color="#7C8BA0", lw=2.0, label="MSCI World IMI", zorder=2)
    ax.set_yscale("log"); ax.set_yticks([1, 2, 5, 10, 20])
    ax.set_yticklabels(["1×", "2×", "5×", "10×", "20×"])
    ax.set_ylim(0.8, 27); ax.set_xlim(2006, 2027.6)
    ax.set_xticks([2006, 2010, 2014, 2018, 2022, 2026])
    ax.annotate(f"{ga[-1]:.0f}×", (x[-1], ga[-1]), xytext=(7, 0), textcoords="offset points",
                color="#3E9CFF", fontsize=17, fontweight="bold", va="center")
    ax.annotate(f"{gm[-1]:.1f}×", (x[-1], gm[-1]), xytext=(7, 0), textcoords="offset points",
                color="#9AA7B6", fontsize=12, fontweight="bold", va="center")
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color("#3A4A60")
    ax.tick_params(colors="#9AA7B6", labelsize=11)
    ax.grid(axis="y", color="#243349", lw=0.9); ax.set_axisbelow(True)
    leg = ax.legend(loc="upper left", frameon=False, fontsize=11)
    for txt in leg.get_texts():
        txt.set_color("#C2CCD8")
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight", transparent=True)
    _plt.close(fig)


_growth_chart("/tmp/trust_growth_dark.png")


def _pct(v):
    return "n.m." if not isinstance(v, (int, float)) or abs(v) > 3 else f"{v*100:+.0f}%"


def _irr(v):
    return "n.m." if not isinstance(v, (int, float)) or v > 3 else f"{v*100:.0f}%"


def _moic(v):
    return f"{v:.1f}×" if isinstance(v, (int, float)) else str(v)


def load_fund2():
    ws = openpyxl.load_workbook("data/AIP_Trackrecord.xlsx", data_only=True)["Transactions"]
    out = []
    for r in range(4, 21):
        c = ws.cell(r, 3).value
        if not c:
            continue
        out.append(dict(period=str(ws.cell(r, 2).value or "").strip(),
                        company=str(c).strip(), irr=ws.cell(r, 6).value,
                        moic=ws.cell(r, 11).value, outp=ws.cell(r, 8).value))
    return out


IND = "(indicative)"

# ===========================================================================
# 1 · COVER
# ===========================================================================
s = newslide()
s.shapes.add_picture(LOGO_WHITE, Inches(0.62), Inches(0.55), height=Inches(0.5),
                     width=Emu(int(int(Inches(0.5)) * _LW_AR)))
para(tbox(s, Inches(0.62), Inches(2.65), Inches(11), Inches(0.4)),
     "A PROPOSED UK INVESTMENT TRUST", 13, MUTED, first=True, after=0)
rrect(s, Inches(0.64), Inches(3.2), Inches(0.6), Pt(3), fill=BLUE, rounded=False)
para(tbox(s, Inches(0.6), Inches(3.45), Inches(12), Inches(1.7)),
     "Engaged ownership, in a listed trust", 46, WHITE, first=True, bold=True,
     after=0, font=HEAD, lead=1.04)
para(tbox(s, Inches(0.62), Inches(5.0), Inches(11), Inches(0.5)),
     "An institutional, 20-year strategy — daily-dealing access for wealth "
     "portfolios", 18, BLUE, first=True, after=0)
para(tbox(s, Inches(0.62), Inches(6.7), Inches(12), Inches(0.4)),
     "For discussion · June 2026 · professional and intermediary use only", 12,
     MUTED, first=True, after=0)

# ===========================================================================
# 2 · AT A GLANCE
# ===========================================================================
s, top = content("1", "Section 1 · Fund overview", "Athanase at a glance",
                 "A proven engaged-ownership strategy — proposed in a listed, "
                 "daily-dealing trust for UK wealth.")
bigstats(s, [("38", "investments as\nengaged owners"),
             ("19%", "avg annualised NET,\nany entry month"),
             ("7%", "worst-ever entry,\nstill positive net"),
             ("18×", "growth of capital\nsince 2006")], top)
by = Emu(int(top) + int(Inches(1.85)))
bullets(s, [
    ("~16% net p.a. since 2006", "versus 6.2% for the MSCI World IMI — +10 pts a "
     "year, net of all fees."),
    ("Entry timing barely matters", "the average entry month returned ~19% net; "
     "even the worst still compounded at +7%."),
    ("No investor loss", "over any full holding period in 20 years."),
], by, x=Inches(0.62), w=Inches(6.0), size=12.5, gap=11)
card(s, Inches(6.9), by, Inches(5.82), Inches(2.55), "An integrated team",
     "A core team together since 2006 (some since 1996) — 38 companies and 30+ "
     "public board seats. Operations, valuation and risk sit apart from "
     "investment decisions, with custody, administration and audit by SEB, MUFG "
     "and KPMG.")

# ===========================================================================
# 3 · TRUST AT A GLANCE (terms)
# ===========================================================================
s, top = content("2", "Section 1 · Fund overview", "The proposed trust at a glance",
                 "Indicative parameters for discussion — final terms to be set "
                 "with the sponsor and prospectus.")
dtable(s, ["Term", "Proposal"], [
    ("Structure", "UK closed-end investment trust, London-listed"),
    ("Listing venue", ("LSE Main Market or Specialist Fund Segment " + IND,)),
    ("Objective", "Long-term capital growth via engaged ownership"),
    ("Portfolio", "8–12 high-conviction listed Nordic / European mid-caps"),
    ("Liquidity", "Daily dealing in the shares; no capital calls"),
    ("Discount control", ("Buyback when shares trade materially below NAV " + IND,)),
    ("Alignment", "Significant team co-investment alongside shareholders"),
    ("Providers", "Custody SEB · Administration MUFG · Audit KPMG"),
], top, colx=[Inches(0.62), Inches(4.1)], rh=Inches(0.46))
para(tbox(s, Inches(0.62), Inches(6.95), Inches(12), Inches(0.3)),
     "Indicative and subject to contract, regulatory approval and final fund "
     "documentation. Not an offer of securities.", 8, MUTED, first=True, after=0)

# ===========================================================================
# 4 · WHY NOW — BETA
# ===========================================================================
s, top = content("3", "Section 2 · The opportunity",
                 "Why now — the easy decade for beta is over",
                 "Equities are expensive on almost every measure, and the "
                 "zero-rate tailwind that lifted all assets has reversed.")
para(tbox(s, Inches(0.62), top, Inches(5.9), Inches(0.4)),
     "Expensive on most measures", 15, WHITE, first=True, bold=True, after=10)
bullets(s, [
    ("", "Shiller CAPE near ~36× — close to the dot-com extreme, roughly double "
     "the long-run ~17×."),
    ("", "Forward P/E ~22× vs a ~16× average; the equity risk premium is at "
     "multi-decade lows."),
    ("", "Most of the 2010s’ return came from re-rating, not earnings growth — "
     "that lever is largely spent."),
], Emu(int(top) + int(Inches(0.5))), x=Inches(0.62), w=Inches(5.85), size=12.5, gap=12)
para(tbox(s, Inches(6.9), top, Inches(5.85), Inches(0.4)),
     "The macro tailwind has reversed", 15, WHITE, first=True, bold=True, after=10)
bullets(s, [
    ("", "The zero-rate, disinflationary backdrop that lifted every asset for a "
     "decade has reversed."),
    ("", "Deglobalisation, deficits and demographics point to structurally "
     "higher inflation and rates."),
    ("", "A higher cost of capital compresses valuations and punishes "
     "long-duration, unprofitable growth."),
], Emu(int(top) + int(Inches(0.5))), x=Inches(6.9), w=Inches(5.85), size=12.5, gap=12)
callout(s, "Whether or not AI lifts productivity, beta is priced for perfection "
           "and the easy tailwind has shifted — so idiosyncratic, "
           "valuation-disciplined returns matter more, not less.")

# ===========================================================================
# 5 · PASSIVE CONCENTRATED
# ===========================================================================
s, top = content("4", "Section 2 · The opportunity",
                 "Passive has become a concentrated bet",
                 "A cap-weighted index puts the most money into the stocks that "
                 "have already risen most — concentration dressed as diversification.")
bigstats(s, [("~38%", "of the S&P 500 in its\n10 largest stocks"),
             ("~33%", "in the “Magnificent\nSeven” alone"),
             ("~50 yrs", "since US equity was\nthis concentrated")], top, h=Inches(1.5))
cy = Emu(int(top) + int(Inches(1.75))); cw = Inches(3.86); gx = Inches(0.26)
for i, (t, b) in enumerate([
    ("Cap-weighting buys high", "The more a stock rises, the more you must own — "
     "the opposite of buy-low discipline."),
    ("A concentrated factor bet", "The “diversified” index is now a leveraged "
     "bet on large-cap growth, momentum and a single theme (AI)."),
    ("Breadth is fragile", "A stumble in a handful of mega-caps now moves the "
     "whole index; the rest has been left behind.")]):
    cx = Emu(int(Inches(0.62)) + i * (int(cw) + int(gx)))
    card(s, cx, cy, cw, Inches(2.15), t, b)

# ===========================================================================
# 6 · PASSIVE RISK -> ACTIVE OPPORTUNITY (+AI)
# ===========================================================================
s, top = content("5", "Section 2 · The opportunity",
                 "From passive risk to active opportunity",
                 "An expensive, concentrated, passive-dominated market is what "
                 "active is built for — and as analysis commoditises, the durable "
                 "edge is ownership, not stock-picking.")
bullets(s, [
    ("The setup —", "expensive equities, record concentration and "
     "price-insensitive passive flows that inflate the most crowded names."),
    ("Why active now —", "wider mispricing and dispersion reward valuation "
     "discipline again; the “active underperforms” critique is cyclical, and "
     "weakest in exactly this market."),
    ("AI commoditises analysis —", "quant and specialist edges compete on the "
     "same public data, which is what commoditises fastest (Kim, Muhn & "
     "Nikolaev, 2024; McLean & Pontiff, 2016)."),
], top, x=Inches(0.62), w=Inches(11.9), size=13, gap=13)
callout(s, "The most durable edge is ownership: buying mispriced businesses and "
           "changing the outcome — an edge AI cannot copy.", y=Inches(5.7))

# ===========================================================================
# 7 · VALUATIONS CRACK (+ live PE gating)
# ===========================================================================
s, top = content("6", "Section 2 · The opportunity",
                 "When valuations crack, where will clients be holding?",
                 "At record valuations the cycle will turn — and the hedge can "
                 "only be bought before the crack, not after.")
cw = Inches(3.86); gx = Inches(0.26)
for i, (t, b) in enumerate([
    ("Global equities", "Fully exposed — clients own the most-expensive, "
     "most-concentrated names, with nothing to do when they fall."),
    ("Private equity", "The trap — marks lag then catch down, while illiquidity "
     "blocks the exit exactly when it is needed."),
    ("This trust", "The hedge — board control floors the risk, the shares stay "
     "liquid, and the dislocation becomes the entry.")]):
    cx = Emu(int(Inches(0.62)) + i * (int(cw) + int(gx)))
    card(s, cx, top, cw, Inches(2.05), t, b)
rrect(s, Inches(0.62), Inches(4.75), Inches(12.1), Inches(1.0), fill=CARD, line=CARDLN, rad=0.06)
rrect(s, Inches(0.62), Inches(4.75), Inches(0.07), Inches(1.0), fill=BLUE, rounded=False)
para(tbox(s, Inches(0.95), Inches(4.86), Inches(11.5), Inches(0.3)),
     "LIVE — JUNE 2026 · BLOOMBERG / CNBC", 9.5, BLUE, first=True, bold=True, after=4)
para(tbox(s, Inches(0.95), Inches(5.16), Inches(11.6), Inches(0.5)),
     "Partners Group capped redemptions in its $8.6bn evergreen PE fund at 5% of "
     "NAV after requests hit 9.8% — KKR, Blackstone and Ares fell. The PE exit is "
     "closing, exactly as feared.", 12, WHITE, first=True, after=0, lead=1.2)
para(tbox(s, Inches(0.62), Inches(6.05), Inches(12.1), Inches(0.5)),
     "While evergreen PE gates redemptions, a listed trust trades every day the "
     "market is open — the liquidity clients keep when it matters.", 13, BLUE,
     first=True, bold=True, after=0, align=PP_ALIGN.CENTER)

# ===========================================================================
# 8 · TWO WAYS TO OWN ACTIVELY
# ===========================================================================
s, top = content("7", "Section 3 · The strategy",
                 "Two ways to own actively: public or private",
                 "The same operational value creation is available in PE and "
                 "public markets — but the public route keeps what private gives up.")
dtable(s, ["Feature", "Private — PE / buyout", "Public — Engaged ownership"], [
    ("How you act", "Buy and control the whole company", ("Board seat and real influence",)),
    ("Entry price", "A ~40% control premium, up front", ("Market price — often at a discount",)),
    ("Liquidity", "10-year lock-up; blind-pool calls", ("Daily liquidity; exit any time",)),
    ("Pricing", "Smoothed quarterly appraisal marks", ("Honest daily market prices",)),
    ("Fees", "2 & 20, plus fees on idle dry powder", ("Lower, on capital deployed",)),
    ("Value engine", "Board-led operational improvement", ("The same — board-led improvement",)),
], top, colx=[Inches(0.62), Inches(3.5), Inches(8.0)], rh=Inches(0.52))

# ===========================================================================
# 9 · THREE HOMES
# ===========================================================================
s, top = content("8", "Section 3 · The strategy",
                 "Where it fits in a portfolio: three homes",
                 "Engaged ownership needs no bucket of its own — it can be funded "
                 "from public equity, alternatives or private equity.")
cw = Inches(3.86); gx = Inches(0.26)
homes = [
    ("Public equity", "A concentrated, active equity sleeve — long-only listed "
     "companies with daily pricing and full transparency.",
     "Adds idiosyncratic, governance-driven alpha to a beta-dominated book."),
    ("Alternatives", "A catalyst strategy — returns from board change, capital "
     "returns and corporate events.",
     "The return driver is event-specific and largely independent of rates and "
     "growth; it diversifies."),
    ("Private equity", "PE-style value creation — board influence, operational "
     "improvement, multi-year holds — in public markets.",
     "The same engaged-ownership engine as PE, but liquid, with no J-curve or "
     "lock-up."),
]
for i, (t, lens, why) in enumerate(homes):
    cx = Emu(int(Inches(0.62)) + i * (int(cw) + int(gx)))
    rrect(s, cx, top, cw, Inches(3.5), fill=CARD, line=CARDLN, rad=0.05)
    inx = Emu(int(cx) + int(Inches(0.28)))
    para(tbox(s, inx, Emu(int(top) + int(Inches(0.24))), Emu(int(cw) - int(Inches(0.5))),
              Inches(0.4)), t, 15, BLUE, first=True, bold=True, after=0)
    bt = tbox(s, inx, Emu(int(top) + int(Inches(0.85))), Emu(int(cw) - int(Inches(0.56))),
              Inches(2.5))
    para(bt, "THE LENS", 9, MUTED, first=True, bold=True, after=4)
    para(bt, lens, 11.5, BODY, after=12, lead=1.3)
    para(bt, "WHY FUND IT HERE", 9, MUTED, bold=True, after=4)
    para(bt, why, 11.5, BODY, after=0, lead=1.3)

# ===========================================================================
# 10 · WHAT WE DO
# ===========================================================================
s, top = content("9", "Section 3 · The strategy",
                 "What we do — we don’t pick stocks, we build companies",
                 "Quality-core constructivism: own good businesses and make them "
                 "better from the boardroom.")
bullets(s, [
    ("Quality at a discount —", "profitable, high-return businesses below "
     "intrinsic value; never broken turnarounds we cannot underwrite."),
    ("Value we create, not discover —", "board control, capital-allocation "
     "discipline and operational change, not leverage or financial engineering."),
    ("Returns from earnings, not re-rating —", "the gain comes from the business "
     "improving under our influence — bankable cash, not a hoped-for multiple."),
    ("Constructive, by agreement —", "board seats won collaboratively, not proxy "
     "wars fought — a Nordic governance model, 20 years honed."),
], top, x=Inches(0.62), w=Inches(11.9), size=13.5, gap=15)
callout(s, "Private-equity-style value creation — in transparent, listed, "
           "daily-liquid form.", y=Inches(6.05))

# ===========================================================================
# 11 · THREE TOLLGATES
# ===========================================================================
s, top = content("10", "Section 4 · Process", "How we invest: three tollgates",
                 "Every idea passes the same three gates — each with an explicit "
                 "bar and a named decision-maker.")
tg = [
    ("Tollgate 1 — Valuation", "“Is it cheap enough even if we change nothing?”",
     "The core business alone justifies the price — a 30–40% margin of safety vs a "
     "PE bidder.  (Investment Team)"),
    ("Tollgate 2 — The plan", "“Is the improvement plan real — and can we deliver it?”",
     "Upside quantified across value levers and validated, with a credible path to "
     "a board seat; includes psychometric evaluation of CEO / Chair.  (Deal Lead + "
     "Head of Research)"),
    ("Tollgate 3 — Go / No-Go", "“Right risk, right size, right fit?”",
     "Within risk and portfolio limits, conviction-sized, with no better use of "
     "the capital.  (Investment Committee)"),
]
yy = top
for i, (h, q, p) in enumerate(tg):
    cy = Emu(int(yy) + i * int(Inches(1.42)))
    shp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.62), cy, Inches(0.5), Inches(0.5))
    shp.shadow.inherit = False; shp.fill.solid(); shp.fill.fore_color.rgb = BLUE
    shp.line.fill.background()
    para(tbox(s, Inches(0.62), cy, Inches(0.5), Inches(0.5), anchor=MSO_ANCHOR.MIDDLE),
         str(i + 1), 16, BG, first=True, bold=True, after=0, align=PP_ALIGN.CENTER)
    tf = tbox(s, Inches(1.35), cy, Inches(11.3), Inches(1.3))
    para(tf, h, 15, BLUE, first=True, bold=True, after=3)
    rp = tf.add_paragraph()
    runs(rp, [("The question:  ", WHITE, True), (q, BODY, False)], 11.5, after=2)
    pp = tf.add_paragraph()
    runs(pp, [("Pass only if:  ", WHITE, True), (p, BODY, False)], 11.5, after=0)

# ===========================================================================
# 12 · OWNERSHIP PLAYBOOK
# ===========================================================================
s, top = content("11", "Section 4 · Process", "The ownership playbook",
                 "A repeatable sequence, honed over 38 investments, that turns a "
                 "board seat into operational value.")
pb = [
    ("1 · Secure board seat", "Often a precondition to invest — the nomination "
     "committee proposes an AIP director."),
    ("2 · Align the thesis", "Collaborate with management and stakeholders to "
     "agree where the value is."),
    ("3 · Refocus earnings", "Cut overhead, optimise footprint, exit loss-making "
     "“growth-trap” ventures."),
    ("4 · Reallocate to winners", "Redeploy freed-up capital into the profitable "
     "core and its strongest adjacencies."),
    ("5 · Grow by acquisition", "Add complementary products or scale through "
     "value-accretive bolt-ons."),
    ("6 · Exit", "Realise via M&A or the equity market once the core is "
     "refocused and compounding."),
]
cw = Inches(3.86); gx = Inches(0.26); chh = Inches(1.4)
for i, (t, b) in enumerate(pb):
    cx = Emu(int(Inches(0.62)) + (i % 3) * (int(cw) + int(gx)))
    cy = Emu(int(top) + (i // 3) * (int(chh) + int(Inches(0.2))))
    rrect(s, cx, cy, cw, chh, fill=CARD, line=CARDLN, rad=0.05)
    inx = Emu(int(cx) + int(Inches(0.26)))
    para(tbox(s, inx, Emu(int(cy) + int(Inches(0.18))), Emu(int(cw) - int(Inches(0.5))),
              Inches(0.35)), t, 13, BLUE, first=True, bold=True, after=5)
    para(tbox(s, inx, Emu(int(cy) + int(Inches(0.6))), Emu(int(cw) - int(Inches(0.52))),
              Inches(0.7)), b, 11, BODY, first=True, after=0, lead=1.22)
para(tbox(s, Inches(0.62), Inches(6.35), Inches(12.1), Inches(0.4)),
     "We don’t buy broken companies — we buy hidden cores.", 15, BLUE, first=True,
     bold=True, after=0, align=PP_ALIGN.CENTER)

# ===========================================================================
# 13 · RISK, HONESTLY
# ===========================================================================
s, top = content("12", "Section 4 · Process", "Risk, honestly",
                 "A risk system built to engineer out permanent loss — controlled "
                 "before the investment is made.")
para(tbox(s, Inches(0.62), top, Inches(5.9), Inches(0.4)),
     "Three pillars of protection", 15, WHITE, first=True, bold=True, after=10)
bullets(s, [
    ("Predictability —", "only durable, market-leading cores in slow-moving "
     "industries where future demand is underwritable."),
    ("Price (a valuation floor) —", "value the rectifiable core, zero for "
     "“growth-trap” divisions — a structural 30–40% margin of safety vs PE."),
    ("Influence —", "a board seat as the “kill switch” — the ability to stop a "
     "bad decision lowers risk below passive or active managers."),
], Emu(int(top) + int(Inches(0.5))), x=Inches(0.62), w=Inches(5.9), size=12, gap=13)
rrect(s, Inches(6.9), top, Inches(5.82), Inches(3.7), fill=CARD, line=CARDLN, rad=0.05)
para(tbox(s, Inches(7.18), Emu(int(top) + int(Inches(0.24))), Inches(5.3), Inches(0.4)),
     "Automatic guardrails", 14, BLUE, first=True, bold=True, after=0)
bullets(s, [
    ("Entry gauntlet —", "ideas must clear ≥12% expected IRR and ≤20% probability "
     "of a 30% drawdown."),
    ("Diversification limits —", "capped exposures to sector (≤30%), cycle, "
     "maturity and geography."),
    ("Mechanised discipline —", "automatic −10 / −20 / −30% triggers force a "
     "re-underwrite, and an exit if the thesis is broken."),
    ("No leverage —", "isolated to single investments, never at the fund level."),
], Emu(int(top) + int(Inches(0.85))), x=Inches(7.18), w=Inches(5.3), size=11, gap=10)

# ===========================================================================
# 14 · VERSUS PRIVATE EQUITY
# ===========================================================================
s, top = content("13", "Section 5 · Track record", "Versus private equity",
                 "On true risk, PE and Athanase carry similar volatility — but "
                 "Athanase returns more, with daily liquidity.")
dtable(s, ["Metric", "Private equity (realised)", "Athanase (engaged owner)"], [
    ("Real return", "~11–14% (above PME, below us)", ("~16% net, real, fully invested",)),
    ("Liquidity", "10-year lock-up, blind pools", ("Daily liquidity, listed",)),
    ("Transparency", "Smoothed, model-marked quarterly", ("Daily marks, public scoreboard",)),
    ("True volatility", "~28% (de-smoothed, levered)", ("27% (downside only ~11%)",)),
    ("Downside risk", "~16% chance of 30%+ loss / 3 yrs", ("~2% chance of 30%+ loss / 3 yrs",)),
], top, colx=[Inches(0.62), Inches(3.5), Inches(8.0)], rh=Inches(0.5))
para(tbox(s, Inches(0.62), Inches(6.7), Inches(12), Inches(0.4)),
     "Athanase captures 93% of the market’s upside but only 43% of its downside; "
     "at 0.44 correlation it improves the whole portfolio’s efficient frontier.",
     10, MUTED, first=True, after=0, italic=True)

# ===========================================================================
# 15 · TRACK RECORD CHART
# ===========================================================================
s, top = content("14", "Section 5 · Track record",
                 "£1 has become ~£18, net — versus ~£3 for the index",
                 "Twenty years of compounding from operating improvement — net of "
                 "all fees, independently reconciled.")
s.shapes.add_picture("/tmp/trust_growth_dark.png", Inches(0.5), Inches(2.35), width=Inches(7.0))
rx = Inches(7.7); ry = Inches(2.45)
for big, lab in [("~16%", "net p.a. over 20 years"),
                 ("+10 pts", "a year vs the index, net"),
                 ("0%", "investor loss over a holding period"),
                 ("2.40", "Sortino ratio — downside-aware")]:
    para(tbox(s, rx, ry, Inches(5.0), Inches(0.5)), big, 27, BLUE, first=True, bold=True, after=0)
    para(tbox(s, rx, Emu(int(ry) + int(Inches(0.48))), Inches(5.0), Inches(0.32)),
         lab, 11, BODY, first=True, after=0)
    ry = Emu(int(ry) + int(Inches(0.86)))
para(tbox(s, Inches(0.5), Inches(6.78), Inches(12.2), Inches(0.3)),
     "Cumulative growth of capital, net of fees, 2006–2025 (log scale). Newly "
     "proposed trust; manager’s existing strategy. Past performance is not a "
     "guide to future returns.", 8, MUTED, first=True, after=0)

# ===========================================================================
# 16 · FULL FUND II BOOK
# ===========================================================================
s, top = content("15", "Section 5 · Track record",
                 "Every investment in the current fund (AIP Fund II)",
                 "The full book, ranked by money multiple — winners, and the "
                 "discipline on the one that disappointed.")
deals = load_fund2(); deals.sort(key=lambda d: d["moic"] if isinstance(d["moic"], (int, float)) else 0, reverse=True)
heads = ("Company", "Holding", "Gross IRR", "MOIC", "vs index")
colx = [Inches(0.62), Inches(3.6), Inches(6.0), Inches(8.3), Inches(10.6)]
for j, h in enumerate(heads):
    para(tbox(s, colx[j], top, Inches(2.4), Inches(0.3)), h, 11, BLUE, first=True, bold=True,
         after=0, align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT)
ry = Emu(int(top) + int(Inches(0.36)))
rrect(s, colx[0], ry, Inches(12.11), Pt(1.2), fill=BLUE, rounded=False)
ry = Emu(int(ry) + int(Inches(0.06)))
rh = Inches(0.255)
for d in deals:
    loss = isinstance(d["moic"], (int, float)) and d["moic"] < 1.0
    vals = [d["company"], d["period"], _irr(d["irr"]), _moic(d["moic"]), _pct(d["outp"])]
    for j, v in enumerate(vals):
        col = (MUTED if loss else (WHITE if j == 0 else BODY))
        para(tbox(s, colx[j], ry, Inches(2.4) if j else Inches(2.9), rh, anchor=MSO_ANCHOR.MIDDLE),
             v, 10, col, first=True, bold=(j == 0), italic=loss,
             align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT, after=0)
    ry = Emu(int(ry) + int(rh))
para(tbox(s, Inches(0.62), Emu(int(ry) + int(Inches(0.05))), Inches(12.1), Inches(0.3)),
     "Gross deal-level figures, AIP Fund II to 2025; “n.m.” where holding periods "
     "are too short to annualise. Past performance is not a guide to future returns.",
     8, MUTED, first=True, after=0)

# ===========================================================================
# 17 · WHY A TRUST
# ===========================================================================
s, top = content("16", "Section 6 · Structure & governance",
                 "Why a listed trust is the ideal home",
                 "A closed-end structure matches long-horizon engaged ownership "
                 "better than any open-ended wrapper.")
bullets(s, [
    ("Permanent capital = patient capital —", "no forced redemptions, so we are "
     "never a forced seller; the multi-year engagement is protected and "
     "dislocations become opportunities."),
    ("Daily liquidity for clients —", "shares trade on the LSE — daily dealing, "
     "no capital calls — while evergreen PE funds gate redemptions (Partners "
     "Group, June 2026)."),
    ("An independent board —", "a majority-independent board oversees the manager "
     "on shareholders’ behalf."),
    ("Tier-one providers —", "operations, valuation and audit segregated from "
     "investment — SEB · MUFG · KPMG."),
], top, x=Inches(0.62), w=Inches(11.9), size=13.5, gap=14)
callout(s, "An institutional strategy clients can buy and sell like any listed "
           "share — without giving up the patient capital it needs.", y=Inches(6.0))

# ===========================================================================
# 18 · DISTRIBUTION & DISCOUNT
# ===========================================================================
s, top = content("17", "Section 6 · Structure & governance",
                 "Income and discount control",
                 "A total-return strategy that the trust structure can still wrap "
                 "with a predictable income and a tight discount.")
card(s, Inches(0.62), top, Inches(5.95), Inches(3.3), "Distribution policy (proposed)",
     "Capital growth at the core — realised as listed gains. UK trusts can pay a "
     "smoothed annual dividend from realised gains and revenue reserves, so a "
     "steady, quarterly income is available where advisers want it. Target yield "
     "and frequency to be set at launch.")
card(s, Inches(6.77), top, Inches(5.95), Inches(3.3), "Discount control",
     "A buyback policy that triggers when the shares trade materially below NAV — "
     "supported by realised exit gains — to keep any discount tight. Listing on "
     "the LSE Main Market or Specialist Fund Segment, weighed on index "
     "eligibility, investor base and compliance.")
para(tbox(s, Inches(0.62), Inches(6.85), Inches(12), Inches(0.3)),
     "Proposed; distribution policy, yield and discount mechanism to be agreed "
     "with the sponsor and board, subject to reserves and regulation.", 8, MUTED,
     first=True, after=0)

# ===========================================================================
# 19 · TIMETABLE
# ===========================================================================
s, top = content("18", "Section 6 · Structure & governance",
                 "An indicative launch timetable",
                 "A standard 3–4 month roadmap from mandate to admission.")
steps = [("MONTH 1", "Structuring, board and provider appointments; mandate and "
          "policies agreed."),
         ("MONTH 2", "Draft prospectus and due diligence; discount and "
          "distribution policies finalised."),
         ("MONTH 3", "Regulatory review; marketing and cornerstone engagement."),
         ("MONTH 4", "Roadshow, book-build and admission to listing.")]
cw = Inches(2.92); gx = Inches(0.16)
for i, (h, b) in enumerate(steps):
    cx = Emu(int(Inches(0.62)) + i * (int(cw) + int(gx)))
    rrect(s, cx, top, cw, Inches(2.4), fill=CARD, line=CARDLN, rad=0.06)
    para(tbox(s, Emu(int(cx) + int(Inches(0.26))), Emu(int(top) + int(Inches(0.24))),
              Emu(int(cw) - int(Inches(0.5))), Inches(0.4)), h, 14, BLUE, first=True,
         bold=True, after=8)
    para(tbox(s, Emu(int(cx) + int(Inches(0.26))), Emu(int(top) + int(Inches(0.85))),
              Emu(int(cw) - int(Inches(0.5))), Inches(1.4)), b, 11.5, BODY, first=True,
         after=0, lead=1.3)
para(tbox(s, Inches(0.62), Inches(6.55), Inches(12.1), Inches(0.3)),
     "Illustrative; actual timetable, venue and discount policy subject to the "
     "sponsor, board and regulatory approval.", 8, MUTED, first=True, after=0)

# ===========================================================================
# 20 · CLOSE
# ===========================================================================
s = newslide()
s.shapes.add_picture(LOGO_WHITE, Inches(0.62), Inches(0.55), height=Inches(0.5),
                     width=Emu(int(int(Inches(0.5)) * _LW_AR)))
rrect(s, Inches(0.64), Inches(2.5), Inches(0.6), Pt(3), fill=BLUE, rounded=False)
para(tbox(s, Inches(0.6), Inches(2.7), Inches(12), Inches(1.0)),
     "Bringing institutional engaged ownership to UK wealth", 34, WHITE,
     first=True, bold=True, after=0, font=HEAD, lead=1.06)
tf = tbox(s, Inches(0.62), Inches(4.0), Inches(12), Inches(2.2))
for ld, b in [
    ("For your clients:", " a liquid, listed, differentiated equity holding — a "
     "20-year net track record, real diversification and downside protection, "
     "daily dealing."),
    ("Why now:", " expensive, concentrated markets, gating private equity and "
     "decaying analytical alpha — the hedge is bought before the crack."),
    ("The proposal:", " partner to launch a UK investment trust on this strategy.")]:
    p = tf.add_paragraph() if tf.paragraphs[0].runs else tf.paragraphs[0]
    runs(p, [(ld, BLUE, True), (b, BODY, False)], 15, after=14, lead=1.25)
para(tbox(s, Inches(0.62), Inches(6.45), Inches(12), Inches(0.5)),
     "Stefan Charette   ·   Athanase Industrial Partner   ·   charette@athanase.se",
     13, MUTED, first=True, after=0)

prs.save("Athanase_Investment_Trust_Pitch.pptx")
print(f"saved ({len(prs.slides._sldIdLst)} slides)")
