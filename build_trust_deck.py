"""Standalone UK INVESTMENT-TRUST pitch (up to 20 slides), structured to the
investment-trust pitch outline (7 sections) but populated with Athanase's actual
engaged-ownership strategy. Launch-specific items (terms, dividend, timetable,
listing) are clearly marked indicative / proposed — not fabricated. Styled to
the AIP brand and rescaled to 4:3. Content distilled from build_combined_deck.py.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn as _qn
from pptx.enum.shapes import MSO_SHAPE_TYPE as _MST
from PIL import Image
import openpyxl

# ---- brand palette / fonts (AIP Blue group only) --------------------------
BLUE1 = RGBColor(0x0E, 0x16, 0x20); BLUE2 = RGBColor(0x15, 0x21, 0x30)
BLUE3 = RGBColor(0x31, 0x43, 0x59); BLUE4 = RGBColor(0x55, 0x6A, 0x83)
BLUE5 = RGBColor(0xE2, 0xE5, 0xE9); BLUE6 = RGBColor(0xF6, 0xF7, 0xF9)
NAVY, NAVY_TX, SLATE, SLATE_LT = BLUE2, BLUE3, BLUE3, BLUE4
HEADERBG, BODY, SUBTLE, DIVIDER = BLUE6, BLUE1, BLUE4, BLUE5
WHITE = RGBColor(0xFF, 0xFF, 0xFF); FOOT = BLUE4; LOSS = BLUE4
SERIF, SANS = "Times New Roman", "Arial"
LOGO_WHITE, MARK_DARK = "assets/logo_white.png", "assets/mark_dark.png"
_LW_AR = (lambda s: s[0] / s[1])(Image.open(LOGO_WHITE).size)
_MD_AR = (lambda s: s[0] / s[1])(Image.open(MARK_DARK).size)

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
_state = {"n": 0}


def rect(slide, l, t, w, h, fill=None):
    sp = slide.shapes.add_shape(1, l, t, w, h); sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    sp.line.fill.background()
    return sp


def tbox(slide, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(l, t, w, h); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


def para(tf, text, size, color, bold=False, italic=False, first=False,
         align=PP_ALIGN.LEFT, after=8, lead=1.1, font=SANS, track=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align; p.space_after = Pt(after); p.line_spacing = lead
    r = p.add_run(); r.text = text; f = r.font
    f.size = Pt(size); f.bold = bold; f.italic = italic
    f.color.rgb = color; f.name = font
    if track is None:
        track = -5 if size >= 48 else (-3 if size >= 24 else 0)
    if track:
        r._r.get_or_add_rPr().set("spc", str(int(round(size * track / 100.0 * 100))))
    return p


def place_mark(s, l, t, h):
    s.shapes.add_picture(MARK_DARK, l, t, height=h, width=Emu(int(int(h) * _MD_AR)))


def place_logo_white(s, l, t, h):
    s.shapes.add_picture(LOGO_WHITE, l, t, height=h, width=Emu(int(int(h) * _LW_AR)))


def footer(s):
    _state["n"] += 1
    para(tbox(s, Inches(8.6), Inches(7.05), Inches(4.4), Inches(0.3)),
         f"Strictly confidential  ·  professional & intermediary use only      "
         f"{_state['n']}", 8, FOOT, first=True, align=PP_ALIGN.RIGHT, after=0)


def content(section, title, subtitle=None):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=WHITE)
    place_mark(s, Inches(0.55), Inches(0.24), Inches(0.26))
    para(tbox(s, Inches(0.98), Inches(0.27), Inches(9.0), Inches(0.3)),
         section, 11, SLATE_LT, first=True, after=0)
    band_h = Inches(1.45) if subtitle else Inches(1.05)
    rect(s, 0, Inches(0.62), SW, band_h, fill=HEADERBG)
    para(tbox(s, Inches(0.6), Inches(0.74), Inches(12.1), Inches(0.85)),
         title, 29, NAVY_TX, first=True, after=2, font=SERIF)
    body_top = Inches(1.95)
    if subtitle:
        para(tbox(s, Inches(0.62), Inches(1.55), Inches(11.9), Inches(0.7)),
             subtitle, 13, SUBTLE, first=True, italic=True, after=0, lead=1.16)
        body_top = Inches(2.4)
    footer(s)
    return s, body_top


def bullets(s, items, top, x=Inches(0.75), w=Inches(11.9), size=14, gap=12, lead=1.18):
    tf = tbox(s, x, top, w, Inches(4.2))
    for i, it in enumerate(items):
        ld, body = it if isinstance(it, tuple) else ("", it)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.line_spacing = lead
        if ld:
            r0 = p.add_run(); r0.text = ld + "  "
            r0.font.size = Pt(size); r0.font.bold = True
            r0.font.color.rgb = NAVY_TX; r0.font.name = SANS
        r1 = p.add_run(); r1.text = body
        r1.font.size = Pt(size); r1.font.color.rgb = BODY; r1.font.name = SANS
    return tf


def closer(s, text, y=Inches(6.46), h=Inches(0.6), size=12.5):
    rect(s, Inches(0.6), y, Inches(12.16), h, fill=NAVY)
    para(tbox(s, Inches(0.8), y, Inches(11.8), h, anchor=MSO_ANCHOR.MIDDLE),
         text, size, WHITE, first=True, italic=True, after=0, track=0)


def statrow(s, stats, top, h=Inches(1.25)):
    n = len(stats); gap = Inches(0.15)
    cw = Emu(int((int(Inches(12.13)) - (n - 1) * int(gap)) / n))
    for i, (big, lab) in enumerate(stats):
        cx = Emu(int(Inches(0.6)) + i * (int(cw) + int(gap)))
        rect(s, cx, top, cw, h, fill=HEADERBG)
        para(tbox(s, cx, Emu(int(top) + int(Inches(0.16))), cw, Inches(0.6),
                  anchor=MSO_ANCHOR.MIDDLE), big, 27, NAVY_TX, first=True,
             after=0, font=SERIF, align=PP_ALIGN.CENTER)
        para(tbox(s, Emu(int(cx) + int(Inches(0.1))),
                  Emu(int(top) + int(Inches(0.8))), Emu(int(cw) - int(Inches(0.2))),
                  Inches(0.4)), lab, 9.5, SLATE_LT, first=True, after=0,
             align=PP_ALIGN.CENTER, lead=1.05)


def kvtable(s, rows, top, x=Inches(0.6), w=Inches(12.13), rh=Inches(0.42),
            lw=Inches(3.3)):
    y = top
    for i, (k, v, *flag) in enumerate(rows):
        rect(s, x, y, w, rh, fill=HEADERBG if i % 2 == 0 else WHITE)
        rect(s, x, y, lw, rh, fill=SLATE if i % 2 else BLUE5)
        para(tbox(s, Emu(int(x) + int(Inches(0.16))), y, Emu(int(lw) - int(Inches(0.3))),
                  rh, anchor=MSO_ANCHOR.MIDDLE), k, 10.5,
             WHITE if i % 2 else NAVY_TX, bold=True, first=True, after=0, lead=1.0)
        vt = tbox(s, Emu(int(x) + int(lw) + int(Inches(0.18))), y,
                  Emu(int(w) - int(lw) - int(Inches(0.32))), rh,
                  anchor=MSO_ANCHOR.MIDDLE)
        p = vt.paragraphs[0]; p.line_spacing = 1.02
        r = p.add_run(); r.text = v
        r.font.size = Pt(10.5); r.font.color.rgb = BODY; r.font.name = SANS
        if flag and flag[0]:
            r2 = p.add_run(); r2.text = "   " + flag[0]
            r2.font.size = Pt(8.5); r2.font.italic = True
            r2.font.color.rgb = SLATE_LT; r2.font.name = SANS
        y = Emu(int(y) + int(rh))
    return y


def stepflow(s, steps, top, h=Inches(1.55)):
    n = len(steps); gap = Inches(0.16)
    cw = Emu(int((int(Inches(12.1)) - (n - 1) * int(gap)) / n))
    for i, (head, body) in enumerate(steps):
        cx = Emu(int(Inches(0.6)) + i * (int(cw) + int(gap)))
        rect(s, cx, top, cw, Inches(0.5), fill=SLATE)
        para(tbox(s, Emu(int(cx) + int(Inches(0.12))), top,
                  Emu(int(cw) - int(Inches(0.2))), Inches(0.5),
                  anchor=MSO_ANCHOR.MIDDLE), head, 9.5, WHITE, first=True,
             bold=True, after=0, lead=1.0, track=0)
        by = Emu(int(top) + int(Inches(0.5)))
        rect(s, cx, by, cw, h, fill=HEADERBG)
        para(tbox(s, Emu(int(cx) + int(Inches(0.14))), Emu(int(by) + int(Inches(0.12))),
                  Emu(int(cw) - int(Inches(0.26))), Emu(int(h) - int(Inches(0.2)))),
             body, 9.5, BODY, first=True, after=0, lead=1.14)
        if i < n - 1:
            para(tbox(s, Emu(int(cx) + int(cw) - int(Inches(0.06))),
                      Emu(int(top) + int(Inches(0.42))), Inches(0.3), Inches(0.5),
                      anchor=MSO_ANCHOR.MIDDLE), "›", 16, SLATE_LT, first=True,
                 after=0, align=PP_ALIGN.CENTER)


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


def dealtable(s, deals, top, x=Inches(0.6), rh=Inches(0.34), font=9.5, hfont=10):
    cw = [Inches(3.2), Inches(2.4), Inches(2.1), Inches(2.0), Inches(2.43)]
    heads = ("Company", "Holding", "Gross IRR", "MOIC", "vs index")
    cx = x
    for ci, ht in enumerate(heads):
        rect(s, cx, top, cw[ci], rh, fill=SLATE)
        para(tbox(s, Emu(int(cx) + int(Inches(0.1))), top, Emu(int(cw[ci]) - int(Inches(0.16))),
                  rh, anchor=MSO_ANCHOR.MIDDLE), ht, hfont, WHITE, bold=True, first=True,
             align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT, after=0)
        cx = Emu(int(cx) + int(cw[ci]))
    y = Emu(int(top) + int(rh))
    for ri, d in enumerate(deals):
        loss = isinstance(d["moic"], (int, float)) and d["moic"] < 1.0
        fill = HEADERBG if ri % 2 == 0 else WHITE
        cx = x
        vals = [d["company"], d["period"], _irr(d["irr"]), _moic(d["moic"]), _pct(d["outp"])]
        for ci, v in enumerate(vals):
            rect(s, cx, y, cw[ci], rh, fill=fill)
            isl = loss and ci >= 2
            col = LOSS if isl else (NAVY_TX if ci == 0 else BODY)
            para(tbox(s, Emu(int(cx) + int(Inches(0.1))), y, Emu(int(cw[ci]) - int(Inches(0.16))),
                      rh, anchor=MSO_ANCHOR.MIDDLE), v, font, col, bold=(ci == 0),
                 italic=isl, first=True,
                 align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT, after=0, track=0)
            cx = Emu(int(cx) + int(cw[ci]))
        y = Emu(int(y) + int(rh))
    return y


IND = "(indicative — to be agreed)"

# ---- performance data + cumulative-growth chart ---------------------------
import numpy as _np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as _plt
H_NAVY, H_BLUE4, H_GRID = "#152130", "#556A83", "#E2E5E9"
from matplotlib import font_manager as _fmgr
for _ff in ("Arial", "Liberation Sans", "DejaVu Sans"):
    try:
        _fmgr.findfont(_ff, fallback_to_default=False)
        _plt.rcParams["font.family"] = _ff
        break
    except Exception:
        continue


def _load_cmp():
    ws = openpyxl.load_workbook("data/Comparison_returns.xlsx",
                                data_only=True)["Sheet1"]

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
    fig, ax = _plt.subplots(figsize=(7.4, 4.05))
    ax.plot(x, ga, color=H_NAVY, lw=2.7, label="Athanase (net)", zorder=3)
    ax.plot(x, gm, color=H_BLUE4, lw=2.0, label="MSCI World IMI", zorder=2)
    ax.set_yscale("log")
    ax.set_yticks([1, 2, 5, 10, 20])
    ax.set_yticklabels(["1×", "2×", "5×", "10×", "20×"])
    ax.set_ylim(0.8, 27); ax.set_xlim(2006, 2027.6)
    ax.set_xticks([2006, 2010, 2014, 2018, 2022, 2026])
    ax.annotate(f"{ga[-1]:.0f}×", (x[-1], ga[-1]), xytext=(7, 0),
                textcoords="offset points", color=H_NAVY, fontsize=16,
                fontweight="bold", va="center")
    ax.annotate(f"{gm[-1]:.1f}×", (x[-1], gm[-1]), xytext=(7, 0),
                textcoords="offset points", color=H_BLUE4, fontsize=12,
                fontweight="bold", va="center")
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color("#AAB2BD")
    ax.tick_params(colors="#556A83", labelsize=10.5)
    ax.grid(axis="y", color=H_GRID, lw=0.8); ax.set_axisbelow(True)
    ax.legend(loc="upper left", frameon=False, fontsize=11)
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight", transparent=True)
    _plt.close(fig)


_growth_chart("/tmp/trust_growth.png")

# ===========================================================================
# 1 · COVER
# ===========================================================================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, fill=NAVY)
place_logo_white(s, Inches(0.6), Inches(0.55), Inches(0.55))
rect(s, Inches(0.62), Inches(3.0), Inches(0.5), Pt(2.4), fill=SLATE_LT)
para(tbox(s, Inches(0.6), Inches(2.45), Inches(11), Inches(0.4)),
     "FOR DISCUSSION · A PROPOSED UK INVESTMENT TRUST", 13, SLATE_LT, first=True, after=0)
para(tbox(s, Inches(0.6), Inches(3.2), Inches(11.8), Inches(1.3)),
     "Engaged ownership, in a listed trust", 42, WHITE, first=True, after=0, font=SERIF)
para(tbox(s, Inches(0.6), Inches(4.55), Inches(11.2), Inches(1.0)),
     "A 20-year institutional engaged-ownership strategy — proposed as a UK "
     "investment trust, giving wealth portfolios listed, liquid access.", 16,
     DIVIDER, first=True, italic=True, after=0, font=SERIF, lead=1.22)
para(tbox(s, Inches(0.6), Inches(6.62), Inches(12), Inches(0.4)),
     "Athanase Industrial Partner    ·    Strictly confidential    ·    For "
     "professional and intermediary use only", 10.5, SLATE_LT, first=True, after=0)

# ===========================================================================
# 2 · EXECUTIVE SUMMARY
# ===========================================================================
s, top = content("Section 1 · Fund overview", "The proposition, in brief",
                 "A listed wrapper around a proven, 20-year engaged-ownership "
                 "strategy — built for UK wealth portfolios.")
bullets(s, [
    ("The mandate.", "long-term capital growth through engaged ownership of "
     "quality, mispriced listed Nordic and European mid-caps — value we create "
     "from the boardroom, not market direction."),
    ("The proof.", "~16% net p.a. over 20 years (+10 pts a year vs the index), a "
     "92% deal hit rate, and no investor loss over a full holding period."),
    ("The wrapper.", "a closed-end trust gives the strategy permanent (patient) "
     "capital, and gives your clients daily dealing on the London Stock Exchange."),
    ("Why now.", "expensive, concentrated markets and AI-driven alpha decay make "
     "a created-alpha, low-correlation diversifier timely."),
], top, size=14, gap=13)
closer(s, "An institutional strategy, in a structure your clients can buy and "
          "sell like any listed share.")

# ===========================================================================
# 3 · AT A GLANCE (INDICATIVE TERMS)
# ===========================================================================
s, top = content("Section 1 · Fund overview",
                 "The proposed trust at a glance",
                 "Indicative parameters for discussion — final terms to be set "
                 "with the sponsor and in the prospectus.")
kvtable(s, [
    ("Structure", "UK closed-end investment trust, London-listed"),
    ("Listing venue", "LSE Main Market or Specialist Fund Segment", IND),
    ("Investment objective", "Long-term capital growth via engaged ownership"),
    ("Portfolio", "8–12 high-conviction listed positions; Nordic / European mid-caps"),
    ("Reference index", "Global / European equities, total return"),
    ("Target raise & fees", "to be agreed with the sponsor", IND),
    ("Liquidity", "daily dealing in the listed shares; no capital calls"),
    ("Discount control", "buyback when shares trade materially below NAV", IND),
    ("Manager alignment", "significant team co-investment alongside shareholders"),
    ("Service providers", "Custody SEB · Administration MUFG · Audit KPMG"),
], top, rh=Inches(0.39))
para(tbox(s, Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.3)),
     "Indicative only and subject to contract, regulatory approval and final "
     "fund documentation. Not an offer of securities.", 8, FOOT, first=True,
     italic=True, after=0)

# ===========================================================================
# 4 · TEAM & ALIGNMENT
# ===========================================================================
s, top = content("Section 1 · Fund overview",
                 "The team, and its alignment with shareholders",
                 "Boutique conviction, tier-one governance — and real skin in "
                 "the game.")
bullets(s, [
    ("A 20-year team.", "the same people compounding capital through multiple "
     "cycles — deep operating and boardroom experience, not a single star "
     "manager."),
    ("Skin in the game.", "the team co-invests its own capital alongside "
     "shareholders — aligned to net realised returns, not assets gathered."),
    ("A proprietary edge.", "a database of ~20,000 CEOs and a behavioural "
     "diligence model — the discipline behind a 92% deal-level hit rate."),
    ("An institutional fortress.", "operations, risk and valuation independent "
     "of investment decisions; custody, administration and audit by SEB, MUFG "
     "and KPMG."),
], top, size=14, gap=13)
closer(s, "The combination wealth committees ask for: conviction, alignment and "
          "independent, tier-one governance.")

# ===========================================================================
# 5 · MARKET OPPORTUNITY — VALUATION DISCONNECT
# ===========================================================================
s, top = content("Section 2 · The opportunity",
                 "Your clients own the most crowded trade in a generation",
                 "Passive equity has become a concentrated, expensive bet on a "
                 "handful of mega-caps — correlated risk dressed as diversification.")
bullets(s, [
    ("A concentrated factor bet.", "the index is more concentrated, and more "
     "expensive, than at almost any point in 50 years — the same names, owned by "
     "everyone."),
    ("It falls together.", "when the leaders re-rate down, “diversified” passive "
     "falls with them; there is no hiding place in more of the same."),
    ("The de-rated alternative.", "engaged ownership of quality mid-caps offers "
     "an attractive entry — a return earned from corporate change, not from the "
     "same mega-caps re-rating ever higher."),
], top, size=14.5, gap=14)
closer(s, "The opportunity: a liquid, differentiated equity holding that earns "
          "its return from company change, not market direction.")

# ===========================================================================
# 6 · STRUCTURAL INEFFICIENCY + AI DURABILITY
# ===========================================================================
s, top = content("Section 2 · The opportunity",
                 "An inefficiency AI can’t arbitrage away",
                 "The edge is not faster analysis — it is governance and "
                 "ownership, where the inefficiency persists.")
bullets(s, [
    ("Under-researched, under-governed.", "rising regulation and falling broker "
     "coverage leave quality mid-caps mispriced — and many carry a correctable "
     "capital-allocation or governance gap."),
    ("Analytical alpha is commoditizing.", "AI lets every manager run the same "
     "financial analysis instantly; the screenable names get bid up and that "
     "edge decays toward zero."),
    ("Created, not discovered.", "our return comes from board control and "
     "operational change — you cannot download a board seat, so the edge "
     "endures as analysis commoditizes."),
], top, size=14.5, gap=14)
para(tbox(s, Inches(0.75), Inches(5.95), Inches(11.9), Inches(0.3)),
     "Mechanism documented: LLMs already perform financial-statement analysis at "
     "analyst level (Kim, Muhn & Nikolaev, 2024); predictability decays once a "
     "strategy is replicable (McLean & Pontiff, 2016).", 8, FOOT, first=True,
     italic=True, after=0)
closer(s, "As analytical alpha commoditizes, engaged-ownership alpha becomes "
          "scarcer — and more valuable.", y=Inches(6.42), h=Inches(0.56))

# ===========================================================================
# 7 · WHY NOW — VALUATIONS CRACK
# ===========================================================================
s, top = content("Section 2 · The opportunity",
                 "When valuations crack, where will your clients be holding?",
                 "At record valuations the cycle will turn — and the hedge can "
                 "only be bought before the crack, not after.")
bullets(s, [
    ("Global equities — fully exposed.", "your clients own the most-expensive, "
     "most-concentrated names, with nothing to do when they fall."),
    ("Private equity — the trap.", "marks lag then catch down, while "
     "illiquidity blocks the exit exactly when it is needed."),
    ("This trust — the hedge.", "board control floors the risk, the shares stay "
     "liquid, and the dislocation becomes the entry point — not the trap."),
], top, size=14.5, gap=15)
closer(s, "The time to own a liquid, low-correlation hedge is while markets are "
          "still expensive.")

# ===========================================================================
# 8 · PHILOSOPHY
# ===========================================================================
s, top = content("Section 3 · Strategy & philosophy",
                 "We don’t pick stocks — we build companies",
                 "Quality-core constructivism: own good businesses, and make them "
                 "better from the boardroom.")
bullets(s, [
    ("Quality at a discount.", "profitable, high-return businesses bought below "
     "intrinsic value — never broken companies or turnarounds we cannot "
     "underwrite."),
    ("Value we create, not discover.", "board control, capital-allocation "
     "discipline and operational change — not leverage or financial engineering."),
    ("Returns from earnings, not re-rating.", "the gain comes from the business "
     "improving under our influence — bankable cash, not a hoped-for multiple."),
    ("Constructive, by agreement.", "board seats won collaboratively, not proxy "
     "wars fought — a Nordic governance model, 20 years honed."),
], top, size=14, gap=12)
closer(s, "Private-equity-style value creation — in transparent, listed, "
          "daily-liquid form.")

# ===========================================================================
# 9 · TARGET COMPANY CRITERIA
# ===========================================================================
s, top = content("Section 3 · Strategy & philosophy",
                 "What makes a target — the tollgates every idea must clear",
                 "A repeatable underwriting bar, applied before a single pound is "
                 "committed.")
bullets(s, [
    ("A high-quality core.", "durable franchise economics and high returns on "
     "incremental capital (ROIIC) — “attractive even if we change nothing.”"),
    ("A structural margin of safety.", "the core alone justifies the price; a "
     "~30–40% discount to intrinsic value underwrites the downside."),
    ("A correctable gap.", "an identifiable capital-allocation, cost or "
     "governance issue we can fix from the board."),
    ("A credible path to alignment.", "a realistic route to a board seat — no "
     "path, no position (our walk-away discipline)."),
], top, size=14, gap=12)
closer(s, "Roughly one new investment a year clears all three tollgates — "
          "selectivity is the first risk control.")

# ===========================================================================
# 10 · PORTFOLIO ARCHITECTURE
# ===========================================================================
s, top = content("Section 3 · Strategy & philosophy",
                 "Portfolio architecture: concentrated, governed, floored",
                 "Deep ownership where we have influence — sized so no single "
                 "position can dominate.")
bullets(s, [
    ("8–12 high-conviction positions.", "concentration is the source of return "
     "and of boardroom influence — we own enough to matter."),
    ("Built in stages.", "a researched toehold, then board alignment secured, "
     "then sized to conviction — never sized to conviction before alignment."),
    ("Hardened guardrails.", "position and concentration limits, plus automatic "
     "−10 / −20 / −30% re-underwrite triggers — the determined-seller framework."),
    ("Liquidity by design.", "listed holdings the trust can exit at will — no "
     "lock-ups, no capital-call queue, no blind pool."),
], top, size=14, gap=12)
closer(s, "Concentrated enough to influence; governed and floored so no holding "
          "can blow up the trust.")

# ===========================================================================
# 11 · INVESTMENT PROCESS
# ===========================================================================
s, top = content("Section 4 · Process & engagement",
                 "The investment process, end to end",
                 "A disciplined lifecycle from idea to exit — with an independent "
                 "Investment Committee at the gate.")
stepflow(s, [
    ("1 · SOURCE", "Screen ~20,000 scored companies for tangible profits and "
     "correctable gaps; TR-1 / ownership signals and executive networks."),
    ("2 · UNDERWRITE", "Desktop modelling, site visits and management "
     "touchpoints; board and management quality scored as rigorously as the "
     "numbers."),
    ("3 · TOLLGATES", "An independent Investment Committee owns sizing and the "
     "Go / No-Go — three tollgates before capital is committed."),
    ("4 · ENGAGE & EXIT", "Secure the board seat, execute the plan, and sell on "
     "a changed thesis — completely, on the facts."),
], top, h=Inches(1.95))
closer(s, "Repeatable and governed — the same blueprint behind a 20-year record.",
       y=Inches(6.42), h=Inches(0.56))

# ===========================================================================
# 12 · ACTIVE ENGAGEMENT & VALUE CREATION
# ===========================================================================
s, top = content("Section 4 · Process & engagement",
                 "Active engagement: how the value is created",
                 "A high-touch, boardroom strategy — value built one disciplined "
                 "decision at a time.")
bullets(s, [
    ("Capital allocation first.", "redirect cash to its highest return — "
     "disciplined M&A, buybacks where the stock is cheap, and an end to "
     "value-destructive spending."),
    ("Operational improvement.", "cost base, pricing and focus — working with "
     "the management we back, and replacing it where we must."),
    ("Incentives realigned.", "pay-for-performance and ownership, so management "
     "is paid to compound value, not to grow assets."),
    ("Consensus, not confrontation.", "we win the boardroom one conversation at "
     "a time — so decisions pass quietly, not in a public fight."),
], top, size=14, gap=12)
closer(s, "The board seat is the lever — it de-risks the asset and turns "
          "governance into the return engine.")

# ===========================================================================
# 13 · RISK SYSTEM
# ===========================================================================
s, top = content("Section 4 · Process & engagement",
                 "A risk system we use — and keep building",
                 "Every engagement tests the system; our own mistakes are written "
                 "back into it.")
bullets(s, [
    ("Agreement before capital.", "no credible path to board alignment, no "
     "position — and if alignment fails, we walk away (e.g. Robit: exited at "
     "~14% IRR; the shares later roughly halved)."),
    ("The determined-seller framework.", "when the thesis breaks, exit "
     "completely on the facts — born from selling DistIT too gradually."),
    ("Mechanised, not emotional.", "−10 / −20 / −30% triggers force a "
     "re-underwrite; sizing and exits owned by the Investment Committee."),
    ("Avoid the un-recoverable.", "we accept optical volatility, but underwrite "
     "against the permanent loss that ends a track record."),
], top, size=14, gap=12)
closer(s, "Discipline that protected capital — and got sharper each time it was "
          "tested.")

# ===========================================================================
# 14 · REPRESENTATIVE INVESTMENTS (track record table)
# ===========================================================================
s, top = content("Section 5 · Portfolio",
                 "Every investment in the current fund (AIP Fund II)",
                 "The full book, not a selection — winners, and the discipline on "
                 "the one that disappointed.")
deals = load_fund2()
deals.sort(key=lambda d: d["moic"] if isinstance(d["moic"], (int, float)) else 0,
           reverse=True)
ytab = dealtable(s, deals, top, rh=Inches(0.235), font=9, hfont=9.5)
para(tbox(s, Inches(0.6), Emu(int(ytab) + int(Inches(0.08))), Inches(12.1), Inches(0.4)),
     "Gross deal-level figures, AIP Fund II, to 2025, ranked by money multiple; "
     "“n.m.” where holding periods are too short to annualise. Past performance "
     "is not a guide to future returns.", 7.5, FOOT, first=True, italic=True,
     after=0, lead=1.1)

# ===========================================================================
# 15 · PERFORMANCE / TRACK RECORD
# ===========================================================================
s, top = content("Section 6 · Track record",
                 "£1 has become ~£18, net — versus ~£3 for the index",
                 "Twenty years of compounding from operating improvement — "
                 "reconciled and audited by tier-one institutions.")
s.shapes.add_picture("/tmp/trust_growth.png", Inches(0.45), Inches(2.5),
                     width=Inches(6.9))
rx = Inches(7.6); ry = Inches(2.55)
for big, lab in [("~16%", "net p.a. over 20 years"),
                 ("+10 pts", "a year vs the index, net"),
                 ("0%", "investor loss over a holding period"),
                 ("2.40", "Sortino ratio — downside-aware")]:
    para(tbox(s, rx, ry, Inches(5.0), Inches(0.5)), big, 26, NAVY_TX,
         first=True, after=0, font=SERIF)
    para(tbox(s, rx, Emu(int(ry) + int(Inches(0.47))), Inches(5.0), Inches(0.32)),
         lab, 11, SLATE_LT, first=True, after=0)
    ry = Emu(int(ry) + int(Inches(0.86)))
para(tbox(s, rx, ry, Inches(5.0), Inches(0.7)),
     "Two market drawdowns (2010, 2020) fully recovered; even the worst entry "
     "month still compounded at +7% net.", 10.5, BODY, first=True, after=0,
     lead=1.18)
para(tbox(s, Inches(0.45), Inches(6.46), Inches(6.95), Inches(0.4)),
     "Cumulative growth of capital, net of fees, 2006–2025 (log scale); SEB · "
     "MUFG · KPMG. Newly-proposed trust; manager’s existing strategy. Past "
     "performance is not a guide to future returns.", 7.5, FOOT, first=True,
     italic=True, after=0, lead=1.1)

# ===========================================================================
# 16 · DIVERSIFIER / DOWNSIDE
# ===========================================================================
s, top = content("Section 6 · Track record",
                 "A genuine diversifier that defends the book",
                 "Low-correlation equity exposure that earns its place beside "
                 "passive and private equity.")
bullets(s, [
    ("0.44 correlation to global equities.", "real diversification — not more of "
     "what your clients already own through passive."),
    ("A board seat floors the fundamental risk.", "we own the cash flow, capex "
     "and pay — so the asset is de-risked whatever the share price does."),
    ("Uncorrelated by construction.", "returns come from earnings and governance "
     "change, not from the index re-rating — a different engine entirely."),
], top, size=14.5, gap=14)
closer(s, "When the index falls, this is built to defend the portfolio — not "
          "amplify the fall.")

# ===========================================================================
# 17 · DISTRIBUTION POLICY (PROPOSED)
# ===========================================================================
s, top = content("Section 6 · Track record",
                 "Distribution policy — designed around the strategy",
                 "The strategy is total-return; the trust structure can still "
                 "deliver a predictable income to clients who want one.")
bullets(s, [
    ("Primarily capital growth.", "the engine is compounding corporate value, "
     "realised as listed gains — a total-return strategy at its core."),
    ("An optional managed distribution.", "UK trusts can pay a smoothed annual "
     "dividend from realised gains and revenue reserves — a steady, "
     "predictable payout where advisers want income.", ),
    ("Quarterly and smoothed.", "a target yield and frequency to be set at "
     "launch — funded from reserves so it does not force selling."),
], top, size=14, gap=14)
para(tbox(s, Inches(0.75), Inches(5.95), Inches(11.9), Inches(0.3)),
     "Proposed; any distribution policy and yield to be set with the sponsor and "
     "board, subject to available reserves and regulation.", 8, FOOT, first=True,
     italic=True, after=0)
closer(s, "Growth at the core — with the option of a smoothed income the trust "
          "wrapper makes possible.", y=Inches(6.42), h=Inches(0.56))

# ===========================================================================
# 18 · WHY A TRUST / GOVERNANCE
# ===========================================================================
s, top = content("Section 7 · Structure & governance",
                 "Why a listed trust is the ideal home for this strategy",
                 "A closed-end structure matches long-horizon engaged ownership "
                 "better than any open-ended wrapper.")
bullets(s, [
    ("Permanent capital = patient capital.", "no forced redemptions, so we are "
     "never a forced seller — the multi-year engagement is protected and "
     "dislocations become opportunities."),
    ("Daily liquidity for your clients.", "shares trade on the London Stock "
     "Exchange — an institutional strategy with daily dealing and no capital "
     "calls."),
    ("An independent board.", "a majority-independent board oversees the manager "
     "on shareholders’ behalf — governance advisers and clients can rely on."),
    ("Tier-one service providers.", "custody, administration and audit "
     "segregated from investment — SEB · MUFG · KPMG."),
], top, size=14, gap=12)
closer(s, "Institutional strategy, listed-share convenience, independent "
          "oversight — without giving up the patient capital it needs.")

# ===========================================================================
# 19 · DISCOUNT CONTROL, LISTING & TIMETABLE
# ===========================================================================
s, top = content("Section 7 · Structure & governance",
                 "Discount control, listing and an indicative timetable",
                 "Proactive NAV discipline, a considered listing venue, and a "
                 "standard launch roadmap.")
bullets(s, [
    ("Discount control.", "a buyback policy that triggers when the shares trade "
     "materially below NAV — keeping any discount tight, supported by realised "
     "exit gains."),
    ("Listing venue.", "LSE Main Market or the Specialist Fund Segment — weighed "
     "on index eligibility, investor base and compliance."),
], top, size=14, gap=12, w=Inches(12.0))
stepflow(s, [
    ("MONTH 1", "Structuring, board and service-provider appointments; mandate "
     "and policies agreed."),
    ("MONTH 2", "Draft prospectus and due diligence; discount and distribution "
     "policies finalised."),
    ("MONTH 3", "Regulatory review; marketing and cornerstone engagement."),
    ("MONTH 4", "Roadshow, book-build and admission to listing."),
], Emu(int(top) + int(Inches(1.7))), h=Inches(1.25))
para(tbox(s, Inches(0.6), Inches(6.5), Inches(12.1), Inches(0.3)),
     "Illustrative 3–4 month roadmap; actual timetable, venue and discount "
     "policy subject to the sponsor, board and regulatory approval.", 8, FOOT,
     first=True, italic=True, after=0)

# ===========================================================================
# 20 · THE ASK / CLOSE
# ===========================================================================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, fill=NAVY)
place_logo_white(s, Inches(0.6), Inches(0.55), Inches(0.5))
rect(s, Inches(0.62), Inches(2.35), Inches(0.5), Pt(2.4), fill=SLATE_LT)
para(tbox(s, Inches(0.6), Inches(2.55), Inches(12), Inches(1.0)),
     "Bringing institutional engaged ownership to UK wealth", 30, WHITE,
     first=True, after=0, font=SERIF)
tf = tbox(s, Inches(0.6), Inches(3.65), Inches(12), Inches(2.4))
for ld, b in [
    ("For your clients:", " a liquid, listed, differentiated equity holding — a "
     "20-year net track record, real diversification and downside protection, "
     "with daily dealing."),
    ("Why now:", " expensive, concentrated markets and decaying analytical alpha "
     "make a created-alpha diversifier timely — and the hedge is bought before "
     "the crack."),
    ("The proposal:", " partner to launch a UK investment trust built on this "
     "strategy.")]:
    p = tf.add_paragraph() if tf.paragraphs[0].runs else tf.paragraphs[0]
    p.space_after = Pt(13); p.line_spacing = 1.2
    r0 = p.add_run(); r0.text = ld; r0.font.size = Pt(15); r0.font.bold = True
    r0.font.color.rgb = WHITE; r0.font.name = SANS
    r1 = p.add_run(); r1.text = b; r1.font.size = Pt(15)
    r1.font.color.rgb = DIVIDER; r1.font.name = SANS
para(tbox(s, Inches(0.6), Inches(6.35), Inches(12), Inches(0.6)),
     "Stefan Charette   ·   Athanase Industrial Partner   ·   "
     "charette@athanase.se", 13, SLATE_LT, first=True, after=0)

# ===========================================================================
# final 4:3 rescale
# ===========================================================================
TARGET_W, TARGET_H = 9753600, 7315200


def _rescale_shape(sh, sx, sy):
    try:
        L, T, W, H = sh.left, sh.top, sh.width, sh.height
    except Exception:
        L = T = W = H = None
    is_pic = sh.shape_type == _MST.PICTURE
    if None not in (L, T, W, H):
        sh.left = int(L * sx); sh.top = int(T * sy)
        sh.width = int(W * sx)
        sh.height = int(H * sx) if is_pic else int(H * sy)
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size is not None:
                    r.font.size = Pt(round(r.font.size.pt * sx, 1))
            pPr = p._p.find(_qn("a:pPr"))
            if pPr is not None:
                for a in ("marL", "indent"):
                    v = pPr.get(a)
                    if v is not None:
                        pPr.set(a, str(int(round(int(v) * sx))))


_sx = TARGET_W / int(prs.slide_width); _sy = TARGET_H / int(prs.slide_height)
for _sl in prs.slides:
    for _sh in _sl.shapes:
        _rescale_shape(_sh, _sx, _sy)
prs.slide_width = TARGET_W; prs.slide_height = TARGET_H

prs.save("Athanase_Investment_Trust_Pitch.pptx")
print(f"saved Athanase_Investment_Trust_Pitch.pptx ({len(prs.slides._sldIdLst)} slides)")
