"""Standalone single page: DistIT case study (acting on a thesis shift; the
origin of the determined-seller framework). NOT part of the combined deck —
built for review, styled to match build_combined_deck.py exactly (incl. the
final 4:3 rescale)."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn as _qn
from pptx.enum.shapes import MSO_SHAPE_TYPE as _MST
from PIL import Image

# ---- palette / fonts (AIP Blue group) -------------------------------------
BLUE1 = RGBColor(0x0E, 0x16, 0x20); BLUE2 = RGBColor(0x15, 0x21, 0x30)
BLUE3 = RGBColor(0x31, 0x43, 0x59); BLUE4 = RGBColor(0x55, 0x6A, 0x83)
BLUE5 = RGBColor(0xE2, 0xE5, 0xE9); BLUE6 = RGBColor(0xF6, 0xF7, 0xF9)
NAVY, NAVY_TX, SLATE, SLATE_LT = BLUE2, BLUE3, BLUE3, BLUE4
HEADERBG, BODY, SUBTLE, DIVIDER = BLUE6, BLUE1, BLUE4, BLUE5
WHITE = RGBColor(0xFF, 0xFF, 0xFF); FOOT = BLUE4
SERIF, SANS = "Times New Roman", "Arial"
MARK_DARK = "assets/mark_dark.png"
_MD_AR = (lambda s: s[0] / s[1])(Image.open(MARK_DARK).size)

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def rect(slide, l, t, w, h, fill=None, line=None, line_w=1.0):
    sp = slide.shapes.add_shape(1, l, t, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(line_w)
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


def place_mark(slide, left, top, height):
    slide.shapes.add_picture(MARK_DARK, left, top, height=height,
                             width=Emu(int(int(height) * _MD_AR)))


def content(section, title, subtitle=None):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=WHITE)
    place_mark(s, Inches(0.55), Inches(0.24), Inches(0.26))
    para(tbox(s, Inches(0.98), Inches(0.27), Inches(7.0), Inches(0.3)),
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
    para(tbox(s, Inches(9.0), Inches(7.05), Inches(4.0), Inches(0.3)),
         "Strictly confidential          proposed", 8.5, FOOT, first=True,
         align=PP_ALIGN.RIGHT, after=0)
    return s, body_top


# ---------------------------------------------------------------------------
s, top = content(
    "Case studies · Risk discipline",
    "Two cases, one discipline: when to walk away, and when to sell",
    "When we cannot secure alignment we walk away (Robit); when the thesis "
    "breaks we sell on the facts, not the share price (DistIT) — and the one "
    "mistake we made built the system.")

colL = tbox(s, Inches(0.75), top, Inches(5.85), Inches(4.3))
para(colL, "DISTIT — A THESIS SHIFT, ACTED ON", 12.5, SLATE, first=True,
     bold=True, after=8)
for t in ["Built it (from 2015): bought in and installed a new management team "
          "and board, compounding the platform over years.",
          "Couldn’t block it, so we governed it (2021): we warned against the "
          "EFUEL acquisition and held a board seat but could not stop it — so we "
          "forced the price down ~40%, from SEK 300m upfront to SEK 180m plus an "
          "earn-out, which never paid.",
          "Acted on the shift: judging the company overvalued, we sold across two "
          "blocks — into a rising price, not a falling one — and still made money.",
          "The honest mistake: we were not determined enough; we should have "
          "exited in full and taken the ~20% hit at once."]:
    para(colL, t, 12.5, BODY, after=8, lead=1.14)
colR = tbox(s, Inches(7.0), top, Inches(5.85), Inches(4.3))
para(colR, "ROBIT — AGREEMENT BEFORE CAPITAL", 12.5, SLATE, first=True,
     bold=True, after=8)
for t in ["Built a position (from May 2015): invested in stages in a niche "
          "drilling-consumables franchise trading below intrinsic value.",
          "Sought engagement (from 2019): pursued a board seat to fix the cost "
          "base and capital allocation — a board-level plan.",
          "The breaking point: we could not secure board alignment.",
          "Walked away: unwound rather than stay without influence — exiting at "
          "~14% IRR (>2× the index). The shares have since roughly halved, so the "
          "discipline avoided the loss."]:
    para(colR, t, 12.5, BODY, after=8, lead=1.14)
rect(s, Inches(0.6), Inches(5.28), Inches(12.16), Inches(0.92), fill=HEADERBG)
rect(s, Inches(0.6), Inches(5.28), Inches(0.07), Inches(0.92), fill=NAVY)
para(tbox(s, Inches(0.85), Inches(5.36), Inches(11.6), Inches(0.3)),
     "THE DISCIPLINE IT BUILT", 12, NAVY, first=True, bold=True, after=4)
para(tbox(s, Inches(0.85), Inches(5.66), Inches(11.7), Inches(0.5)),
     "Two routes, one rule — act on the facts, not the share price. Robit set "
     "the precondition (no board alignment, no capital); DistIT set the exit "
     "(when the thesis breaks, sell completely). Selling DistIT too gradually is "
     "the origin of our determined-seller framework.", 11, BODY, first=True,
     after=0, lead=1.14)
rect(s, Inches(0.6), Inches(6.46), Inches(12.16), Inches(0.6), fill=NAVY)
para(tbox(s, Inches(0.8), Inches(6.46), Inches(11.8), Inches(0.6),
          anchor=MSO_ANCHOR.MIDDLE),
     "Conviction to build, the discipline to exit, and the honesty to "
     "systematise our own mistakes — both cases protected capital, and made the "
     "risk system better.", 12.5, WHITE, first=True, italic=True, after=0,
     track=0)

# ---- final 4:3 rescale (match the combined deck output) --------------------
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
                for attr in ("marL", "indent"):
                    v = pPr.get(attr)
                    if v is not None:
                        pPr.set(attr, str(int(round(int(v) * sx))))


_sx = TARGET_W / int(prs.slide_width); _sy = TARGET_H / int(prs.slide_height)
for _sl in prs.slides:
    for _sh in _sl.shapes:
        _rescale_shape(_sh, _sx, _sy)
prs.slide_width = TARGET_W; prs.slide_height = TARGET_H

prs.save("Athanase_DistIT_page.pptx")
print("saved Athanase_DistIT_page.pptx")
