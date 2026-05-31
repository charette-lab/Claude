"""Standalone, brand-compliant slide: switching a mid-market private-equity
allocation for Athanase.

Athanase stats are computed live from data/Comparison_returns.xlsx (real).
Mid-market PE figures are clearly-labelled ASSUMPTIONS (no PE return series
supplied): ~13% net IRR; reported (appraisal-smoothed) downside risk with a
de-smoothed 'true economic' estimate. Built 4:3 per AIP Brand Guidelines.
"""
import math
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
from lxml import etree
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as _fmgr

# ---- AIP Brand palette (Blue group only) -----------------------------------
BLUE1 = RGBColor(0x0E, 0x16, 0x20); BLUE2 = RGBColor(0x15, 0x21, 0x30)
BLUE3 = RGBColor(0x31, 0x43, 0x59); BLUE4 = RGBColor(0x55, 0x6A, 0x83)
BLUE5 = RGBColor(0xE2, 0xE5, 0xE9); BLUE6 = RGBColor(0xF6, 0xF7, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NAVY, NAVY_TX, SLATE, SLATE_LT = BLUE2, BLUE3, BLUE3, BLUE4
HEADERBG, BODY, SUBTLE, FOOT = BLUE6, BLUE1, BLUE4, BLUE4
SERIF, SANS = "Times New Roman", "Arial"
H_NAVY, H_BLUE3, H_BLUE4 = "#152130", "#314359", "#556A83"
H_BLUE5, H_BODY = "#E2E5E9", "#0E1620"
MARK_DARK = "assets/mark_dark.png"
_MD_AR = (lambda s: s[0] / s[1])(Image.open(MARK_DARK).size)
for _ff in ("Arial", "Liberation Sans", "DejaVu Sans"):
    try:
        _fmgr.findfont(_ff, fallback_to_default=False)
        plt.rcParams["font.family"] = _ff
        break
    except Exception:
        continue

# ===========================================================================
# Data: Athanase (real) + mid-market PE (assumptions)
# ===========================================================================
ws = openpyxl.load_workbook("data/Comparison_returns.xlsx", data_only=True)["Sheet1"]
_ATH = []
for c in range(2, ws.max_column + 1):
    for r in range(22, 34):
        v = ws.cell(r, c).value
        if isinstance(v, (int, float)):
            _ATH.append(v)
_NM = len(_ATH)


def _ann(xs):
    g = 1.0
    for x in xs:
        g *= (1 + x)
    return g ** (12 / len(xs)) - 1


def _dvol(xs):
    return math.sqrt(sum(min(x, 0) ** 2 for x in xs) / len(xs)) * math.sqrt(12)


ATH_RET, ATH_DD = _ann(_ATH), _dvol(_ATH)        # ~16.0%, ~10.9%  (real)

# Mid-market PE — ASSUMPTIONS (no return series supplied)
PE_RET = 0.13                 # net-of-fees IRR to LPs
PE_DD_REPORTED = 0.07         # appraisal-smoothed downside volatility (as seen)
PE_DESMOOTH = 1.8             # de-smoothing factor (autocorrelation adjustment)
PE_DD_TRUE = PE_DD_REPORTED * PE_DESMOOTH        # ~12.6% true economic downside

# ===========================================================================
# Build the slide
# ===========================================================================
prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
s = prs.slides.add_slide(prs.slide_layouts[6])


def rect(sl, l, t, w, h, fill=None):
    shp = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    if fill is not None:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    shp.line.fill.background(); shp.shadow.inherit = False
    return shp


def tbox(sl, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = sl.shapes.add_textbox(l, t, w, h); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


def para(tf, text, size, color, bold=False, italic=False, first=False,
         align=PP_ALIGN.LEFT, after=8, lead=1.1, font=SANS, track=0):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align; p.space_after = Pt(after); p.line_spacing = lead
    r = p.add_run(); r.text = text
    f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic
    f.color.rgb = color; f.name = font
    if track:
        r._r.get_or_add_rPr().set("spc", str(int(round(size * track / 100 * 100))))
    return p


# header
s.shapes.add_picture(MARK_DARK, Inches(0.55), Inches(0.24), height=Inches(0.26),
                     width=Emu(int(Inches(0.26) * _MD_AR)))
para(tbox(s, Inches(0.98), Inches(0.27), Inches(7), Inches(0.3)),
     "Risk-Adjusted Outcomes", 11, SLATE_LT, first=True, after=0)
para(tbox(s, Inches(9.9), Inches(0.27), Inches(2.85), Inches(0.3)),
     "Figure 1", 11, SLATE_LT, first=True, bold=True, align=PP_ALIGN.RIGHT, after=0)
rect(s, 0, Inches(0.62), SW, Inches(1.45), fill=HEADERBG)
para(tbox(s, Inches(0.6), Inches(0.74), Inches(12.1), Inches(0.85)),
     "Switching a private-equity sleeve into Athanase", 30, NAVY_TX,
     first=True, after=2, font=SERIF)
para(tbox(s, Inches(0.62), Inches(1.55), Inches(11.9), Inches(0.5)),
     "Same operational value creation — at a higher return, with liquidity, "
     "lower fees and daily transparency the private-equity wrapper gives up.",
     13, SUBTLE, first=True, italic=True, after=0)


# ---- left: return-vs-risk scatter ------------------------------------------
def scatter_chart(path_png):
    fig, ax = plt.subplots(figsize=(6.0, 4.25), dpi=200)
    # PE reported -> true: arrow showing the smoothing illusion
    ax.annotate("", (PE_DD_TRUE * 100, PE_RET * 100),
                (PE_DD_REPORTED * 100, PE_RET * 100),
                arrowprops=dict(arrowstyle="->", color=H_BLUE4, lw=1.6,
                                linestyle=(0, (3, 2))))
    ax.scatter([PE_DD_REPORTED * 100], [PE_RET * 100], s=150, color="white",
               edgecolor=H_BLUE4, linewidth=2.2, zorder=5)
    ax.annotate("Mid-market PE\n(as reported*)", (PE_DD_REPORTED * 100, PE_RET * 100),
                color=H_BLUE4, fontsize=10, fontweight="bold", xytext=(-6, 16),
                textcoords="offset points", ha="center")
    ax.scatter([PE_DD_TRUE * 100], [PE_RET * 100], s=150, color=H_BLUE4, zorder=5,
               edgecolor="white", linewidth=1.5)
    ax.annotate("PE (de-smoothed,\ntrue economic*)", (PE_DD_TRUE * 100, PE_RET * 100),
                color=H_BLUE4, fontsize=10, xytext=(8, -22),
                textcoords="offset points", ha="center")
    # Athanase
    ax.scatter([ATH_DD * 100], [ATH_RET * 100], s=210, color=H_NAVY, zorder=6,
               edgecolor="white", linewidth=1.5)
    ax.annotate(f"Athanase\n{ATH_RET*100:.0f}% return", (ATH_DD * 100, ATH_RET * 100),
                color=H_NAVY, fontsize=11, fontweight="bold", xytext=(10, 6),
                textcoords="offset points", va="center")
    ax.set_xlabel("Downside volatility (annualised)", color=H_BODY, fontsize=11)
    ax.set_ylabel("Net return to investors", color=H_BODY, fontsize=11)
    ax.set_xlim(4, 16); ax.set_ylim(8, 19)
    ax.xaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax.tick_params(colors=H_BODY, labelsize=10)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color(H_BLUE5)
    ax.grid(True, color=H_BLUE5, lw=0.6, alpha=0.6)
    ax.set_title("Net return vs true downside risk", color=H_NAVY, fontsize=13,
                 fontweight="bold", pad=10)
    fig.tight_layout(); fig.savefig(path_png, dpi=200, bbox_inches="tight",
                                    facecolor="white"); plt.close(fig)


scatter_chart("/tmp/pe_scatter.png")
s.shapes.add_picture("/tmp/pe_scatter.png", Inches(0.55), Inches(2.28),
                     height=Inches(3.95))
para(tbox(s, Inches(0.7), Inches(6.25), Inches(6.2), Inches(0.5)),
     "Marked daily like Athanase, PE’s appraisal-smoothed risk roughly doubles "
     "— its “low volatility” is largely an accounting artefact.",
     10, SUBTLE, first=True, italic=True, after=0, lead=1.12)

# ---- right: what you gain by switching (cards) -----------------------------
cardx = Inches(7.05); cardw = Inches(5.7); cardh = Inches(0.86); cy = Inches(2.3)
gains = [
    ("RETURN", f"+{(ATH_RET-PE_RET)*100:.0f} pts net",
     "16% vs ~13% net IRR — more return, not less, for the switch."),
    ("LIQUIDITY", "Daily vs 10-yr lock-up",
     "Listed positions you can exit; no blind-pool capital calls."),
    ("FEES & CAPITAL", "Lower drag, fully invested",
     "No 2&20 on undeployed dry powder waiting in a queue."),
    ("TRANSPARENCY", "Daily marks, real control",
     "A public scoreboard plus board control of cash flow, capex and pay."),
    ("RISK, HONESTLY", "Lower true downside",
     "De-smoothed, PE’s economic downside is ~13% vs Athanase’s ~11%."),
]
for title, big, body in gains:
    rect(s, cardx, cy, cardw, cardh, fill=HEADERBG)
    para(tbox(s, Emu(int(cardx) + int(Inches(0.2))), cy + Inches(0.1),
              Inches(2.0), Inches(0.25)),
         title, 9.5, SLATE, first=True, bold=True, after=0)
    para(tbox(s, Emu(int(cardx) + int(Inches(2.05))), cy + Inches(0.09),
              Inches(3.4), Inches(0.3)),
         big, 13, NAVY_TX, first=True, bold=True, after=0, font=SERIF)
    para(tbox(s, Emu(int(cardx) + int(Inches(0.2))), cy + Inches(0.4),
              Emu(int(cardw) - int(Inches(0.4))), Inches(0.42)),
         body, 10, BODY, first=True, after=0, lead=1.08)
    cy = Emu(int(cy) + int(cardh) + int(Inches(0.1)))

# takeaway strip
rect(s, Inches(0.6), Inches(6.62), Inches(12.13), Inches(0.46), fill=NAVY)
para(tbox(s, Inches(0.78), Inches(6.62), Inches(11.8), Inches(0.46),
          anchor=MSO_ANCHOR.MIDDLE),
     "Switching a PE sleeve into Athanase keeps the operational upside but adds "
     "return, liquidity and transparency — and removes the smoothing that hides "
     "PE’s true risk.", 12, WHITE, first=True, italic=True, after=0)
para(tbox(s, Inches(0.6), Inches(7.16), Inches(12.6), Inches(0.42)),
     f"Athanase: net monthly returns, {_NM} months (2006–2025), real. *Mid-market "
     "PE shown from assumptions (≈13% net IRR; reported downside ≈7%, de-smoothed "
     "≈13% for appraisal autocorrelation), not a supplied return series. "
     "Illustrative; past performance is not indicative of future results.",
     7.5, FOOT, first=True, after=0, lead=1.1)

# ===========================================================================
# Convert to 4:3 + brand the theme
# ===========================================================================
TARGET_W, TARGET_H = 9753600, 7315200
from pptx.oxml.ns import qn as _qn
from pptx.enum.shapes import MSO_SHAPE_TYPE as _MST


def _rescale(sh, sx, sy):
    try:
        L, T, W, H = sh.left, sh.top, sh.width, sh.height
    except Exception:
        L = None
    if L is not None:
        sh.left, sh.top = int(L * sx), int(T * sy)
        if sh.shape_type == _MST.PICTURE:
            sh.width, sh.height = int(W * sx), int(W * sx * H / W)
        else:
            sh.width, sh.height = int(W * sx), int(H * sy)
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size is not None:
                    r.font.size = Pt(round(r.font.size.pt * sx, 1))


sx, sy = TARGET_W / int(prs.slide_width), TARGET_H / int(prs.slide_height)
for sh in s.shapes:
    _rescale(sh, sx, sy)
prs.slide_width, prs.slide_height = TARGET_W, TARGET_H


def _brand_theme(prs):
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    REL = ("http://schemas.openxmlformats.org/officeDocument/2006/"
           "relationships/theme")
    scheme = {"dk1": "152130", "lt1": "FFFFFF", "dk2": "0E1620", "lt2": "F6F7F9",
              "accent1": "152130", "accent2": "314359", "accent3": "556A83",
              "accent4": "0E1620", "accent5": "E2E5E9", "accent6": "F6F7F9",
              "hlink": "314359", "folHlink": "556A83"}
    seen = set()
    for master in prs.slide_masters:
        th = master.part.part_related_by(REL)
        if th.partname in seen:
            continue
        seen.add(th.partname)
        root = etree.fromstring(th.blob)
        clr = root.find(f"{a}themeElements/{a}clrScheme")
        for nm, hx in scheme.items():
            el = clr.find(f"{a}{nm}")
            if el is None:
                continue
            srgb = el.find(f"{a}srgbClr"); sysc = el.find(f"{a}sysClr")
            if srgb is not None:
                srgb.set("val", hx)
            elif sysc is not None:
                el.remove(sysc); etree.SubElement(el, f"{a}srgbClr", {"val": hx})
        fonts = root.find(f"{a}themeElements/{a}fontScheme")
        for grp, face in (("majorFont", "Times New Roman"), ("minorFont", "Arial")):
            latin = fonts.find(f"{a}{grp}/{a}latin")
            if latin is not None:
                latin.set("typeface", face)
        th._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                  standalone=True)


_brand_theme(prs)
out = "Athanase_vs_Private_Equity.pptx"
prs.save(out)
print("Saved", out)
print(f"Athanase (real): {ATH_RET*100:.1f}% ret, {ATH_DD*100:.1f}% downside")
print(f"PE (assumed): {PE_RET*100:.0f}% ret, reported {PE_DD_REPORTED*100:.0f}% "
      f"-> de-smoothed {PE_DD_TRUE*100:.1f}% downside")
