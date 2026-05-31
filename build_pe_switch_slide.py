"""Standalone, brand-compliant slides: Athanase vs mid-market private equity,
using REAL (committed-capital) PE returns rather than headline IRR.

Slide 1 — reconciliation bridge: why the ~20% headline IRR an LP is shown is
          really ~8.6% once measured honestly (committed-capital / PME basis).
Slide 2 — switching a PE sleeve into Athanase, on the real numbers.

Athanase stats are computed live from data/Comparison_returns.xlsx (real).
Private-equity figures are sourced from published research (see footnotes):
  - Top-quartile mid-market buyout headline net IRR ~18-22%  (Cambridge Assoc.)
  - Real committed-basis return ~8.6%, ~matching the S&P 500  (Phalippou 2020)
  - IRR reinvestment assumption inflates by ~3 pts  (evidence-based research)
  - Reported vol ~11% de-smooths to ~22% total / ~13% downside  (MSCI/Morningstar)
Built 4:3 per AIP Brand Guidelines.
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
# Data
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


ATH_RET, ATH_DD = _ann(_ATH), _dvol(_ATH)        # ~16.0%, ~10.9% (real)

# Private equity — from published research (see footnotes)
PE_HEADLINE = 0.20            # top-quartile mid-market buyout headline net IRR
PE_IRR_MATH = 0.03            # IRR reinvestment-assumption inflation (~3 pts)
PE_REAL = 0.086              # real committed-capital / PME basis (Phalippou)
PE_DD_REPORTED = 0.07         # reported (smoothed) downside vol
PE_DD_TRUE = 0.13             # de-smoothed (true economic) downside vol

# ===========================================================================
# Helpers
# ===========================================================================
prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height


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
         align=PP_ALIGN.LEFT, after=8, lead=1.1, font=SANS):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align; p.space_after = Pt(after); p.line_spacing = lead
    r = p.add_run(); r.text = text
    f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic
    f.color.rgb = color; f.name = font
    return p


def header(sl, kicker, ref, title, subtitle):
    sl.shapes.add_picture(MARK_DARK, Inches(0.55), Inches(0.24), height=Inches(0.26),
                          width=Emu(int(Inches(0.26) * _MD_AR)))
    para(tbox(sl, Inches(0.98), Inches(0.27), Inches(7), Inches(0.3)),
         kicker, 11, SLATE_LT, first=True, after=0)
    para(tbox(sl, Inches(9.9), Inches(0.27), Inches(2.85), Inches(0.3)),
         ref, 11, SLATE_LT, first=True, bold=True, align=PP_ALIGN.RIGHT, after=0)
    rect(sl, 0, Inches(0.62), SW, Inches(1.45), fill=HEADERBG)
    para(tbox(sl, Inches(0.6), Inches(0.74), Inches(12.1), Inches(0.85)),
         title, 30, NAVY_TX, first=True, after=2, font=SERIF)
    para(tbox(sl, Inches(0.62), Inches(1.55), Inches(11.9), Inches(0.5)),
         subtitle, 13, SUBTLE, first=True, italic=True, after=0)


# ===========================================================================
# SLIDE 1 — reconciliation bridge: headline IRR -> real committed return
# ===========================================================================
def bridge_chart(path_png):
    fig, ax = plt.subplots(figsize=(8.7, 4.35), dpi=200)
    # waterfall: start full, two drops, end full
    labels = ["Headline\ntop-quartile\nIRR",
              "IRR reinvestment\nassumption",
              "Cash drag on committed\ncapital, fees & timing",
              "Real return on\ncommitted capital\n(PME basis)"]
    start = PE_HEADLINE * 100
    after_math = start - PE_IRR_MATH * 100
    end = PE_REAL * 100
    drop2 = after_math - end
    xs = [0, 1, 2, 3]
    # bar 1 (headline, solid navy)
    ax.bar(0, start, 0.62, color=H_NAVY, zorder=3)
    # drop 1 (IRR math) — floating
    ax.bar(1, PE_IRR_MATH * 100, 0.62, bottom=after_math, color=H_BLUE4,
           alpha=0.75, zorder=3)
    # drop 2 (cash drag) — floating
    ax.bar(2, drop2, 0.62, bottom=end, color=H_BLUE4, alpha=0.75, zorder=3)
    # bar 4 (real, solid navy)
    ax.bar(3, end, 0.62, color=H_NAVY, zorder=3)
    # connector lines
    for x0, y in [(0, start), (1, after_math), (2, end)]:
        ax.plot([x0 + 0.31, x0 + 1 - 0.31], [y, y], color=H_BLUE4, lw=1.0,
                ls=(0, (3, 2)), zorder=2)
    # value labels
    ax.annotate(f"{start:.0f}%", (0, start), ha="center", va="bottom",
                fontsize=15, fontweight="bold", color=H_NAVY,
                xytext=(0, 4), textcoords="offset points")
    ax.annotate(f"–{PE_IRR_MATH*100:.0f} pts", (1, after_math + PE_IRR_MATH * 50),
                ha="center", va="center", fontsize=11, color=H_BLUE4,
                fontweight="bold")
    ax.annotate(f"–{drop2:.0f} pts", (2, end + drop2 / 2), ha="center",
                va="center", fontsize=11, color=H_BLUE4, fontweight="bold")
    ax.annotate(f"{end:.1f}%", (3, end), ha="center", va="bottom", fontsize=15,
                fontweight="bold", color=H_NAVY, xytext=(0, 4),
                textcoords="offset points")
    # Athanase reference line
    ax.axhline(ATH_RET * 100, color=H_NAVY, lw=1.6, ls=(0, (5, 3)), zorder=4)
    ax.annotate(f"Athanase {ATH_RET*100:.0f}% (real, fully invested)",
                (3.55, ATH_RET * 100), ha="right", va="bottom", fontsize=10.5,
                color=H_NAVY, fontweight="bold")
    ax.set_xticks(xs); ax.set_xticklabels(labels, fontsize=9.5, color=H_BODY)
    ax.set_ylabel("Annualised return to investors", color=H_BODY, fontsize=11)
    ax.set_ylim(0, 23); ax.set_xlim(-0.6, 3.7)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax.tick_params(colors=H_BODY, labelsize=10)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color(H_BLUE5)
    ax.grid(True, axis="y", color=H_BLUE5, lw=0.6, alpha=0.6)
    ax.set_title("What the LP is shown vs what the LP earns",
                 color=H_NAVY, fontsize=13, fontweight="bold", pad=10)
    fig.tight_layout(); fig.savefig(path_png, dpi=200, bbox_inches="tight",
                                    facecolor="white"); plt.close(fig)


bridge_chart("/tmp/pe_bridge.png")
s1 = prs.slides.add_slide(prs.slide_layouts[6])
header(s1, "Risk-Adjusted Outcomes", "Figure 1",
       "Private equity’s headline return is not what you earn",
       "Headline IRR is struck on drawn capital only. Measured on the capital "
       "you commit — the honest basis — the return is far lower.")
s1.shapes.add_picture("/tmp/pe_bridge.png", Inches(0.55), Inches(2.28),
                      height=Inches(3.95))
# right: explanation cards
cx = Inches(8.95); cw = Inches(3.8); chh = Inches(1.2); cy = Inches(2.32)
expl = [
    ("THE REINVESTMENT TRICK",
     "IRR assumes uncalled cash earns the IRR — inflating the figure by ~3 pts."),
    ("THE COMMITTED-CAPITAL DRAG",
     "You commit 100 but ~half sits in cash earning little, while paying fees — "
     "overstating returns by up to ~50%."),
    ("THE LIKE-FOR-LIKE TRUTH",
     "On a public-market-equivalent basis, 2006–18 PE returned ~8.6% — about the "
     "same as the S&P 500, after $400bn+ in fees."),
]
for title, body in expl:
    rect(s1, cx, cy, cw, chh, fill=HEADERBG)
    para(tbox(s1, Emu(int(cx) + int(Inches(0.2))), cy + Inches(0.12),
              Emu(int(cw) - int(Inches(0.4))), Inches(0.25)),
         title, 10.5, SLATE, first=True, bold=True, after=0)
    para(tbox(s1, Emu(int(cx) + int(Inches(0.2))), cy + Inches(0.4),
              Emu(int(cw) - int(Inches(0.4))), Inches(0.75)),
         body, 11, BODY, first=True, after=0, lead=1.14)
    cy = Emu(int(cy) + int(chh) + int(Inches(0.16)))
rect(s1, Inches(0.6), Inches(6.62), Inches(12.13), Inches(0.46), fill=NAVY)
para(tbox(s1, Inches(0.78), Inches(6.62), Inches(11.8), Inches(0.46),
          anchor=MSO_ANCHOR.MIDDLE),
     "Athanase’s 16% is already a committed-capital return — fully invested from "
     "day one. Compared like-for-like, it is roughly double real PE.",
     12, WHITE, first=True, italic=True, after=0)
para(tbox(s1, Inches(0.6), Inches(7.1), Inches(12.6), Inches(0.5)),
     "Sources: top-quartile headline IRR ~18–22% (Cambridge Associates US PE "
     "benchmark, 2024); real ~8.6% public-market-equivalent return, 2006–18 "
     "(Phalippou, 2020); committed-capital / undrawn-commitment drag (Meyer, "
     "2020, J. Alternative Investments; Hayley & Sefiloglu, 2022; Buchner, "
     "Kaserer & Wagner, 2010). Bridge magnitudes illustrative, reconciling cited "
     "endpoints. Athanase real, net, 2006–2025.",
     7.5, FOOT, first=True, after=0, lead=1.1)


# ===========================================================================
# SLIDE 2 — switching a PE sleeve into Athanase (on the real numbers)
# ===========================================================================
def scatter_chart(path_png):
    fig, ax = plt.subplots(figsize=(6.0, 4.25), dpi=200)
    # PE marketed -> real: arrow down and to the right
    ax.annotate("", (PE_DD_TRUE * 100, PE_REAL * 100),
                (PE_DD_REPORTED * 100, PE_HEADLINE * 100),
                arrowprops=dict(arrowstyle="->", color=H_BLUE4, lw=1.8,
                                linestyle=(0, (3, 2))))
    ax.scatter([PE_DD_REPORTED * 100], [PE_HEADLINE * 100], s=150, color="white",
               edgecolor=H_BLUE4, linewidth=2.2, zorder=5)
    ax.annotate("PE as marketed\n(headline IRR,\nsmoothed risk)",
                (PE_DD_REPORTED * 100, PE_HEADLINE * 100), color=H_BLUE4,
                fontsize=9.5, fontweight="bold", xytext=(6, 10),
                textcoords="offset points", ha="left")
    ax.scatter([PE_DD_TRUE * 100], [PE_REAL * 100], s=150, color=H_BLUE4, zorder=5,
               edgecolor="white", linewidth=1.5)
    ax.annotate("PE as earned\n(committed basis,\ntrue risk)",
                (PE_DD_TRUE * 100, PE_REAL * 100), color=H_BLUE4, fontsize=9.5,
                xytext=(8, -6), textcoords="offset points", ha="left", va="top")
    # Athanase
    ax.scatter([ATH_DD * 100], [ATH_RET * 100], s=210, color=H_NAVY, zorder=6,
               edgecolor="white", linewidth=1.5)
    ax.annotate(f"Athanase\n{ATH_RET*100:.0f}% return", (ATH_DD * 100, ATH_RET * 100),
                color=H_NAVY, fontsize=11, fontweight="bold", xytext=(-10, 10),
                textcoords="offset points", va="center", ha="right")
    ax.set_xlabel("Downside volatility (annualised)", color=H_BODY, fontsize=11)
    ax.set_ylabel("Real net return to investors", color=H_BODY, fontsize=11)
    ax.set_xlim(4, 16); ax.set_ylim(6, 23)
    ax.xaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax.tick_params(colors=H_BODY, labelsize=10)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color(H_BLUE5)
    ax.grid(True, color=H_BLUE5, lw=0.6, alpha=0.6)
    ax.set_title("Real return vs true downside risk", color=H_NAVY, fontsize=13,
                 fontweight="bold", pad=10)
    fig.tight_layout(); fig.savefig(path_png, dpi=200, bbox_inches="tight",
                                    facecolor="white"); plt.close(fig)


scatter_chart("/tmp/pe_scatter.png")
s2 = prs.slides.add_slide(prs.slide_layouts[6])
header(s2, "Risk-Adjusted Outcomes", "Figure 2",
       "Switching a private-equity sleeve into Athanase",
       "On the honest numbers, Athanase out-returns real PE with more liquidity, "
       "lower fees and daily transparency.")
s2.shapes.add_picture("/tmp/pe_scatter.png", Inches(0.55), Inches(2.28),
                      height=Inches(3.95))
para(tbox(s2, Inches(0.7), Inches(6.28), Inches(6.2), Inches(0.4)),
     "Measured honestly, PE moves down and to the right — lower return, higher "
     "true risk — while Athanase sits up and to the left.",
     10, SUBTLE, first=True, italic=True, after=0, lead=1.12)
cardx = Inches(7.05); cardw = Inches(5.7); cardh = Inches(0.86); cy = Inches(2.3)
gains = [
    ("RETURN", f"~16% vs ~9% real",
     "Roughly double real PE on a committed-capital basis."),
    ("LIQUIDITY", "Daily vs 10-yr lock-up",
     "Listed positions you can exit; no blind-pool capital calls."),
    ("FEES & CAPITAL", "Lower drag, fully invested",
     "No fees on undeployed dry powder waiting in a queue."),
    ("TRANSPARENCY", "Daily marks, real control",
     "A public scoreboard plus board control of cash flow, capex and pay."),
    ("RISK, HONESTLY", "Lower true downside",
     "De-smoothed, PE’s economic downside is ~13% vs Athanase’s ~11%."),
]
for title, big, body in gains:
    rect(s2, cardx, cy, cardw, cardh, fill=HEADERBG)
    para(tbox(s2, Emu(int(cardx) + int(Inches(0.2))), cy + Inches(0.1),
              Inches(2.0), Inches(0.25)), title, 9.5, SLATE, first=True,
         bold=True, after=0)
    para(tbox(s2, Emu(int(cardx) + int(Inches(2.05))), cy + Inches(0.09),
              Inches(3.4), Inches(0.3)), big, 13, NAVY_TX, first=True, bold=True,
         after=0, font=SERIF)
    para(tbox(s2, Emu(int(cardx) + int(Inches(0.2))), cy + Inches(0.4),
              Emu(int(cardw) - int(Inches(0.4))), Inches(0.42)), body, 10, BODY,
         first=True, after=0, lead=1.08)
    cy = Emu(int(cy) + int(cardh) + int(Inches(0.1)))
rect(s2, Inches(0.6), Inches(6.62), Inches(12.13), Inches(0.46), fill=NAVY)
para(tbox(s2, Inches(0.78), Inches(6.62), Inches(11.8), Inches(0.46),
          anchor=MSO_ANCHOR.MIDDLE),
     "Switching a PE sleeve into Athanase keeps the operational upside but adds "
     "return, liquidity and transparency — once PE is measured the way you "
     "actually experience it.", 12, WHITE, first=True, italic=True, after=0)
para(tbox(s2, Inches(0.6), Inches(7.16), Inches(12.6), Inches(0.42)),
     f"Athanase: net monthly returns, {_NM} months (2006–2025), real. PE real "
     "return ~8.6% committed basis (Phalippou, 2020); downside vol reported ~7% "
     "de-smoothed to ~13% (MSCI; Morningstar). Illustrative; past performance is "
     "not indicative of future results.",
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
for _sl in prs.slides:
    for sh in _sl.shapes:
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
print("Saved", out, "(2 slides)")
print(f"Athanase real: {ATH_RET*100:.1f}% ret, {ATH_DD*100:.1f}% downside")
print(f"PE headline {PE_HEADLINE*100:.0f}% -> real {PE_REAL*100:.1f}% committed basis")
