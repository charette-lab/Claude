"""Standalone 10-slide pitch: Athanase engaged-ownership strategy proposed as a
UK INVESTMENT TRUST, summarised for a wealth advisor. Separate from the main
deck; styled to the AIP brand (Blue group, Times/Arial, real logo) and rescaled
to the same 4:3 output. Content distilled from build_combined_deck.py."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn as _qn
from pptx.enum.shapes import MSO_SHAPE_TYPE as _MST
from PIL import Image

# ---- brand palette / fonts (AIP Blue group only) --------------------------
BLUE1 = RGBColor(0x0E, 0x16, 0x20); BLUE2 = RGBColor(0x15, 0x21, 0x30)
BLUE3 = RGBColor(0x31, 0x43, 0x59); BLUE4 = RGBColor(0x55, 0x6A, 0x83)
BLUE5 = RGBColor(0xE2, 0xE5, 0xE9); BLUE6 = RGBColor(0xF6, 0xF7, 0xF9)
NAVY, NAVY_TX, SLATE, SLATE_LT = BLUE2, BLUE3, BLUE3, BLUE4
HEADERBG, BODY, SUBTLE, DIVIDER = BLUE6, BLUE1, BLUE4, BLUE5
WHITE = RGBColor(0xFF, 0xFF, 0xFF); FOOT = BLUE4
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


def place_mark(slide, l, t, h):
    slide.shapes.add_picture(MARK_DARK, l, t, height=h, width=Emu(int(int(h) * _MD_AR)))


def place_logo_white(slide, l, t, h):
    slide.shapes.add_picture(LOGO_WHITE, l, t, height=h, width=Emu(int(int(h) * _LW_AR)))


def footer(s, dark=False):
    _state["n"] += 1
    para(tbox(s, Inches(9.0), Inches(7.05), Inches(4.0), Inches(0.3)),
         f"Strictly confidential          {_state['n']}", 8.5,
         SLATE_LT if dark else FOOT, first=True, align=PP_ALIGN.RIGHT, after=0)


def content(section, title, subtitle=None):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=WHITE)
    place_mark(s, Inches(0.55), Inches(0.24), Inches(0.26))
    para(tbox(s, Inches(0.98), Inches(0.27), Inches(8.0), Inches(0.3)),
         section, 11, SLATE_LT, first=True, after=0)
    band_h = Inches(1.45) if subtitle else Inches(1.05)
    rect(s, 0, Inches(0.62), SW, band_h, fill=HEADERBG)
    para(tbox(s, Inches(0.6), Inches(0.74), Inches(12.1), Inches(0.85)),
         title, 30, NAVY_TX, first=True, after=2, font=SERIF)
    body_top = Inches(1.95)
    if subtitle:
        para(tbox(s, Inches(0.62), Inches(1.55), Inches(11.9), Inches(0.7)),
             subtitle, 13, SUBTLE, first=True, italic=True, after=0, lead=1.16)
        body_top = Inches(2.4)
    footer(s)
    return s, body_top


def bullets(s, items, top, x=Inches(0.75), w=Inches(11.9), size=14, gap=12):
    tf = tbox(s, x, top, w, Inches(4.2))
    for i, it in enumerate(items):
        lead, body = it if isinstance(it, tuple) else ("", it)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.line_spacing = 1.18
        if lead:
            r0 = p.add_run(); r0.text = lead + "  "
            r0.font.size = Pt(size); r0.font.bold = True
            r0.font.color.rgb = NAVY_TX; r0.font.name = SANS
        r1 = p.add_run(); r1.text = body
        r1.font.size = Pt(size); r1.font.color.rgb = BODY; r1.font.name = SANS
    return tf


def closer(s, text, y=Inches(6.46), h=Inches(0.6), size=12.5):
    rect(s, Inches(0.6), y, Inches(12.16), h, fill=NAVY)
    para(tbox(s, Inches(0.8), y, Inches(11.8), h, anchor=MSO_ANCHOR.MIDDLE),
         text, size, WHITE, first=True, italic=True, after=0, track=0)


def statrow(s, stats, top):
    n = len(stats); gap = Inches(0.15)
    cw = Emu(int((int(Inches(12.13)) - (n - 1) * int(gap)) / n))
    for i, (big, lab) in enumerate(stats):
        cx = Emu(int(Inches(0.6)) + i * (int(cw) + int(gap)))
        rect(s, cx, top, cw, Inches(1.3), fill=HEADERBG)
        para(tbox(s, cx, Emu(int(top) + int(Inches(0.18))), cw, Inches(0.6),
                  anchor=MSO_ANCHOR.MIDDLE),
             big, 29, NAVY_TX, first=True, after=0, font=SERIF,
             align=PP_ALIGN.CENTER)
        para(tbox(s, Emu(int(cx) + int(Inches(0.12))),
                  Emu(int(top) + int(Inches(0.82))), Emu(int(cw) - int(Inches(0.24))),
                  Inches(0.4)),
             lab, 10, SLATE_LT, first=True, after=0, align=PP_ALIGN.CENTER,
             lead=1.05)


# ===========================================================================
# 1 · COVER
# ===========================================================================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, fill=NAVY)
place_logo_white(s, Inches(0.6), Inches(0.55), Inches(0.55))
rect(s, Inches(0.62), Inches(3.0), Inches(0.5), Pt(2.4), fill=SLATE_LT)
para(tbox(s, Inches(0.6), Inches(2.45), Inches(11), Inches(0.4)),
     "FOR DISCUSSION · A PROPOSED UK INVESTMENT TRUST", 13, SLATE_LT, first=True,
     after=0)
para(tbox(s, Inches(0.6), Inches(3.2), Inches(11.8), Inches(1.3)),
     "Engaged ownership, in a listed trust", 42, WHITE, first=True, after=0,
     font=SERIF)
para(tbox(s, Inches(0.6), Inches(4.55), Inches(11.2), Inches(1.0)),
     "A 20-year institutional engaged-ownership strategy — proposed as a UK "
     "investment trust, giving wealth portfolios listed, liquid access.", 16,
     DIVIDER, first=True, italic=True, after=0, font=SERIF, lead=1.22)
para(tbox(s, Inches(0.6), Inches(6.62), Inches(12), Inches(0.4)),
     "Athanase Industrial Partner    ·    Strictly confidential    ·    For "
     "professional and intermediary use only", 10.5, SLATE_LT, first=True, after=0)

# ===========================================================================
# 2 · THE OPPORTUNITY
# ===========================================================================
s, top = content(
    "The opportunity",
    "Your clients own the most crowded trade in a generation",
    "Passive equity has quietly become a concentrated bet on a handful of "
    "expensive mega-caps — correlated risk dressed as diversification.")
bullets(s, [
    ("A concentrated factor bet.", "the index is more concentrated, and more "
     "expensive, than at almost any point in 50 years — the same names, owned by "
     "everyone."),
    ("It falls together.", "when the leaders re-rate down, “diversified” passive "
     "falls with them; there is no hiding place in more of the same."),
    ("Clients need a real diversifier.", "an equity return that compounds from "
     "company change — not from the same mega-caps re-rating ever higher."),
], top, size=14.5, gap=14)
closer(s, "The opportunity: a liquid, differentiated equity holding that earns "
          "its return from corporate change, not market direction.")

# ===========================================================================
# 3 · WHAT WE DO
# ===========================================================================
s, top = content(
    "What we do",
    "We don’t pick stocks — we build companies",
    "Concentrated, long-only public equities, where we create the value "
    "ourselves through ownership and governance.")
bullets(s, [
    ("8–12 high-conviction positions.", "deep ownership in quality, mispriced "
     "businesses — roughly one new investment a year, held for years."),
    ("Value we create, not discover.", "board control, capital-allocation "
     "discipline and operational change — not leverage or financial engineering."),
    ("Returns from earnings, not re-rating.", "the gain comes from the business "
     "getting better under our influence, whatever the market does."),
    ("A Nordic governance model, 20 years honed.", "constructive engagement by "
     "agreement — board seats won, not proxy wars fought."),
], top, size=14, gap=12)
closer(s, "Private-equity-style operational value creation — in transparent, "
          "listed, daily-liquid form.")

# ===========================================================================
# 4 · THE PROOF
# ===========================================================================
s, top = content(
    "The track record",
    "Two decades of net returns, independently validated",
    "A 20-year record built on operating improvement — reconciled and audited by "
    "tier-one institutions.")
statrow(s, [("~16%", "net p.a. over 20 years"),
            ("+10 pts", "a year vs the index, net"),
            ("92%", "deal hit rate (35 of 38)"),
            ("0%", "investor loss over a holding period")], top)
bullets(s, [
    ("Downside protection.", "two market drawdowns fully recovered (2010, 2020); "
     "Sortino ratio 2.40 — the return came with genuine resilience."),
    ("Entry timing is a non-risk.", "even the worst entry month in 20 years still "
     "compounded at +7% net — there is no wrong moment to start."),
    ("Independently verified.", "custody by SEB, administration by MUFG, audit by "
     "KPMG — an institutional paper trail."),
], Emu(int(top) + int(Inches(1.55))), size=13.5, gap=10)
para(tbox(s, Inches(0.6), Inches(6.06), Inches(12.1), Inches(0.3)),
     "Net of fees, 2006–2025, independently reconciled. Past performance is not "
     "a guide to future returns.", 8, FOOT, first=True, italic=True, after=0)
closer(s, "Returns from operating improvement — repeatable, and verified by "
          "tier-one institutions.", y=Inches(6.42), h=Inches(0.56), size=12)

# ===========================================================================
# 5 · A GENUINE DIVERSIFIER
# ===========================================================================
s, top = content(
    "Diversification",
    "The diversifier that defends the book when markets fall",
    "Low-correlation equity exposure that earns its place beside passive and "
    "private equity.")
bullets(s, [
    ("0.44 correlation to global equities.", "real diversification — not more of "
     "what your clients already own through passive."),
    ("A board seat floors the fundamental risk.", "we own the cash flow, capex "
     "and pay — so the asset is de-risked whatever the share price does."),
    ("Uncorrelated by construction.", "returns come from earnings and governance "
     "change, not from the index re-rating — a different engine entirely."),
], top, size=14.5, gap=14)
closer(s, "When the index falls, this is the holding that is built to defend the "
          "portfolio — not amplify the fall.")

# ===========================================================================
# 6 · WHY AN INVESTMENT TRUST
# ===========================================================================
s, top = content(
    "The structure",
    "Why a listed trust is the ideal home for this strategy",
    "A closed-end investment trust matches a long-horizon engaged-ownership "
    "strategy better than any open-ended wrapper.")
bullets(s, [
    ("Permanent capital = patient capital.", "a closed-end structure means we "
     "are never a forced seller — the multi-year engagement horizon is protected, "
     "and dislocations become buying opportunities, not redemptions."),
    ("Daily liquidity for your clients.", "shares trade on the London Stock "
     "Exchange — an institutional strategy with daily dealing, no lock-ups and no "
     "capital calls."),
    ("Governance your clients can trust.", "an independent board, plus operations "
     "and valuation segregated from investment — SEB · MUFG · KPMG."),
    ("NAV discipline built in.", "a buyback mechanism to keep the share price "
     "tracking close to net asset value."),
], top, size=13.5, gap=11)
closer(s, "An institutional strategy your clients can buy and sell like any "
          "listed share — without giving up the patient capital it needs.")

# ===========================================================================
# 7 · WHY NOW (1) — THE REGIME
# ===========================================================================
s, top = content(
    "Why now · the regime",
    "When valuations crack, where will your clients be holding?",
    "At record valuations and concentration, the cycle will turn — and the time "
    "to own the hedge is before it does.")
bullets(s, [
    ("Global equities — fully exposed.", "your clients own the most-expensive, "
     "most-concentrated names, with nothing to do when they fall."),
    ("Private equity — the trap.", "marks lag then catch down, while illiquidity "
     "blocks the exit exactly when it is needed."),
    ("This trust — the hedge.", "board control floors the risk, the shares stay "
     "liquid, and the dislocation becomes the entry point — not the trap."),
], top, size=14.5, gap=14)
closer(s, "The hedge can only be bought before the crack, not after — which is "
          "what makes the timing matter now.")

# ===========================================================================
# 8 · WHY NOW (2) — AI
# ===========================================================================
s, top = content(
    "Why now · AI",
    "AI is commoditizing the alpha your clients pay for",
    "Long-only equity alpha rests on financial analysis — exactly what AI now "
    "does instantly, and universally.")
bullets(s, [
    ("The easy edge is decaying.", "as every manager runs the same AI-driven "
     "analysis, the screenable names get bid up and the analytical edge erodes "
     "toward zero."),
    ("Our return is created, not discovered.", "board control and operational "
     "change cannot be downloaded or arbitraged away."),
    ("Scarcer means more valuable.", "as analytical alpha commoditizes, "
     "engaged-ownership alpha becomes rarer — and worth more."),
], top, size=14.5, gap=14)
para(tbox(s, Inches(0.75), Inches(5.95), Inches(11.9), Inches(0.3)),
     "Mechanism documented: LLMs already perform financial-statement analysis at "
     "analyst level (Kim, Muhn & Nikolaev, 2024); predictability decays once a "
     "strategy is replicable (McLean & Pontiff, 2016).", 8, FOOT, first=True,
     italic=True, after=0)
closer(s, "Own the edge AI cannot copy.", y=Inches(6.42), h=Inches(0.56))

# ===========================================================================
# 9 · WHY ATHANASE
# ===========================================================================
s, top = content(
    "Why Athanase",
    "An institutional team, inside an institutional fortress",
    "Boutique conviction with tier-one governance — defensible to your clients "
    "and your committee.")
bullets(s, [
    ("A 20-year team.", "the same people, compounding through multiple cycles — "
     "deep operating and boardroom experience, not a star manager."),
    ("A proprietary edge.", "a database of ~20,000 CEOs and a behavioural "
     "diligence model — the discipline behind a 92% deal-level hit rate."),
    ("An institutional fortress.", "operations, risk and valuation independent of "
     "investment decisions; custody, administration and audit by SEB, MUFG and "
     "KPMG."),
], top, size=14.5, gap=14)
closer(s, "Boutique conviction, tier-one governance — the combination wealth "
          "committees ask for.")

# ===========================================================================
# 10 · THE ASK / CLOSE
# ===========================================================================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, fill=NAVY)
place_logo_white(s, Inches(0.6), Inches(0.55), Inches(0.5))
rect(s, Inches(0.62), Inches(2.35), Inches(0.5), Pt(2.4), fill=SLATE_LT)
para(tbox(s, Inches(0.6), Inches(2.55), Inches(12), Inches(1.0)),
     "Bringing institutional engaged ownership to UK wealth", 30, WHITE,
     first=True, after=0, font=SERIF)
tf = tbox(s, Inches(0.6), Inches(3.6), Inches(12), Inches(2.4))
for lead, b in [
    ("For your clients:", " a liquid, listed, differentiated equity holding — a "
     "20-year net track record, real diversification and downside protection, "
     "with daily dealing."),
    ("Why now:", " expensive, concentrated markets and decaying analytical alpha "
     "make a created-alpha diversifier timely — and the hedge is bought before "
     "the crack."),
    ("The proposal:", " partner to launch a UK investment trust built on this "
     "strategy.")]:
    p = tf.add_paragraph() if tf.paragraphs[0].runs else tf.paragraphs[0]
    p.space_after = Pt(14); p.line_spacing = 1.2
    r0 = p.add_run(); r0.text = lead; r0.font.size = Pt(15); r0.font.bold = True
    r0.font.color.rgb = WHITE; r0.font.name = SANS
    r1 = p.add_run(); r1.text = b; r1.font.size = Pt(15)
    r1.font.color.rgb = DIVIDER; r1.font.name = SANS
para(tbox(s, Inches(0.6), Inches(6.35), Inches(12), Inches(0.6)),
     "Stefan Charette   ·   Athanase Industrial Partner   ·   "
     "charette@athanase.se", 13, SLATE_LT, first=True, after=0)

# ===========================================================================
# final 4:3 rescale (match the AIP deck output)
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
print(f"saved Athanase_Investment_Trust_Pitch.pptx ({len(prs.slides.__iter__.__self__._sldIdLst)} slides)")
