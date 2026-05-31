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

# Private equity — REAL allocator-realised returns from institutional datasets
# (PME / cash-matched basis). PE genuinely earns a premium over public markets;
# the honest benchmark is what allocators realise (~11-14%), not headline IRR.
PE_HEADLINE = 0.20            # top-quartile mid-market buyout headline net IRR
PE_REAL = 0.114              # Cliffwater: US state-pension PE realised, 2000-22
PE_REAL_LO, PE_REAL_HI = 0.114, 0.141   # Cliffwater (US) .. BVCA (Europe)
# evidence pairs: (label, PE realised %, public-market-equivalent %)
PE_EVIDENCE = [
    ("Cliffwater\nUS state pensions\n(2000–2022)", 11.4, 5.8),
    ("BVCA / Capital Dynamics\nUK & Europe\n(2001–2023)", 14.1, 7.7),
]
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
# SLIDE 1 — what allocators ACTUALLY earn: realised PE vs its public-market
# equivalent, across institutional datasets, with Athanase for reference
# ===========================================================================
def realised_chart(path_png):
    fig, ax = plt.subplots(figsize=(8.7, 4.35), dpi=200)
    groups = [e[0] for e in PE_EVIDENCE]
    pe_vals = [e[1] for e in PE_EVIDENCE]
    pme_vals = [e[2] for e in PE_EVIDENCE]
    x = np.arange(len(groups)); w = 0.38
    ax.bar(x - w / 2, pe_vals, w, color=H_BLUE4, label="PE realised (net)")
    ax.bar(x + w / 2, pme_vals, w, color=H_BLUE5, edgecolor=H_BLUE4, linewidth=1,
           label="Public-market equivalent")
    for xi, (pe, pme) in enumerate(zip(pe_vals, pme_vals)):
        ax.annotate(f"{pe:.1f}%", (xi - w / 2, pe), ha="center", va="bottom",
                    fontsize=12, fontweight="bold", color=H_BLUE3,
                    xytext=(0, 3), textcoords="offset points")
        ax.annotate(f"{pme:.1f}%", (xi + w / 2, pme), ha="center", va="bottom",
                    fontsize=11, color=H_BLUE4, xytext=(0, 3),
                    textcoords="offset points")
        # place premium label between the two bars, clear of the Athanase line
        ax.annotate(f"+{pe-pme:.1f} pts\npremium", (xi, (pe + pme) / 2),
                    ha="center", va="center", fontsize=9.5, color=H_BLUE3,
                    fontweight="bold")
    # Athanase reference line
    ax.axhline(ATH_RET * 100, color=H_NAVY, lw=1.8, ls=(0, (5, 3)), zorder=5)
    ax.annotate(f"Athanase {ATH_RET*100:.0f}%  (real, net, fully invested)",
                (len(groups) - 0.5, ATH_RET * 100), ha="right", va="bottom",
                fontsize=10.5, color=H_NAVY, fontweight="bold")
    ax.set_xticks(x); ax.set_xticklabels(groups, fontsize=9.5, color=H_BODY)
    ax.set_ylabel("Annualised return to investors", color=H_BODY, fontsize=11)
    ax.set_ylim(0, 19); ax.set_xlim(-0.6, len(groups) - 0.4)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax.tick_params(colors=H_BODY, labelsize=10)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color(H_BLUE5)
    ax.grid(True, axis="y", color=H_BLUE5, lw=0.6, alpha=0.6)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=2,
              fontsize=10, frameon=False, labelcolor=H_BODY)
    ax.set_title("What allocators actually realise from private equity",
                 color=H_NAVY, fontsize=13, fontweight="bold", pad=10)
    fig.tight_layout(); fig.savefig(path_png, dpi=200, bbox_inches="tight",
                                    facecolor="white"); plt.close(fig)


realised_chart("/tmp/pe_realised.png")
s1 = prs.slides.add_slide(prs.slide_layouts[6])
header(s1, "Risk-Adjusted Outcomes", "Figure 1",
       "What allocators actually earn from private equity",
       "Stripping out headline-IRR marketing, cash-matched (PME) data shows PE "
       "earns a real ~3–6 pt premium over public markets — but lands below "
       "Athanase.")
s1.shapes.add_picture("/tmp/pe_realised.png", Inches(0.55), Inches(2.28),
                      height=Inches(3.95))
# right: explanation cards
cx = Inches(8.95); cw = Inches(3.8); chh = Inches(1.2); cy = Inches(2.32)
expl = [
    ("HEADLINE ≠ REALISED",
     "Top-quartile IRR is shown as ~18–22%, struck on drawn capital. Cash-matched "
     "to public markets, the real figure is far lower."),
    ("THE REAL PREMIUM IS GENUINE",
     "PME studies agree: ~1.2× wealth multiple, a real 3–5 pt premium over "
     "public equities (Cambridge; Kaplan-Schoar / Burgiss; Cliffwater; BVCA)."),
    ("BUT BELOW ATHANASE",
     "Allocators realise ~11–14% from PE. Athanase has delivered ~16% — net, "
     "real, and fully invested from day one."),
]
for title, body in expl:
    rect(s1, cx, cy, cw, chh, fill=HEADERBG)
    para(tbox(s1, Emu(int(cx) + int(Inches(0.2))), cy + Inches(0.1),
              Emu(int(cw) - int(Inches(0.4))), Inches(0.25)),
         title, 10.5, SLATE, first=True, bold=True, after=0)
    para(tbox(s1, Emu(int(cx) + int(Inches(0.2))), cy + Inches(0.36),
              Emu(int(cw) - int(Inches(0.4))), Inches(0.8)),
         body, 10.5, BODY, first=True, after=0, lead=1.12)
    cy = Emu(int(cy) + int(chh) + int(Inches(0.16)))
rect(s1, Inches(0.6), Inches(6.62), Inches(12.13), Inches(0.46), fill=NAVY)
para(tbox(s1, Inches(0.78), Inches(6.62), Inches(11.8), Inches(0.46),
          anchor=MSO_ANCHOR.MIDDLE),
     "PE’s premium over public markets is real — but Athanase’s ~16% sits above "
     "even what the best allocators realise, with daily liquidity.",
     12, WHITE, first=True, italic=True, after=0)
para(tbox(s1, Inches(0.6), Inches(7.1), Inches(12.6), Inches(0.5)),
     "Sources: Cliffwater (US state pensions, 2000–2022): PE 11.4% vs PME 5.8%. "
     "BVCA / Capital Dynamics (UK & Europe, 2001–2023): 14.1% vs 7.7% (PME+). "
     "Corroborated by Cambridge Associates mPME (1.15–1.25×) and "
     "Kaplan-Harris-Jenkinson KS-PME on Burgiss (1.20–1.27×), ≈ +3 pts/yr. "
     "Athanase real, net, 2006–2025.",
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
                fontsize=9.5, fontweight="bold", xytext=(-6, 8),
                textcoords="offset points", ha="left")
    ax.scatter([PE_DD_TRUE * 100], [PE_REAL * 100], s=150, color=H_BLUE4, zorder=5,
               edgecolor="white", linewidth=1.5)
    ax.annotate("PE as realised\n(PME basis,\ntrue risk)",
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
    ("RETURN", "~16% vs ~11–14% realised",
     "Above what even the best allocators realise from PE (Cliffwater; BVCA)."),
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
     f"Athanase: net monthly returns, {_NM} months (2006–2025), real. PE realised "
     "~11–14% (Cliffwater US state pensions 11.4%, 2000–22; BVCA UK/Europe 14.1%, "
     "2001–23) vs headline IRR ~18–22%; downside vol reported ~7% de-smoothed to "
     "~13% (MSCI; Morningstar). Illustrative; past performance ≠ future results.",
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
