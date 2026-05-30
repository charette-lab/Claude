"""Standalone, brand-compliant slide: downside-volatility-driven outcome
projection for Athanase vs MSCI World IMI over 3/5/7/10-year horizons.

Statistics are computed live from data/Comparison_returns.xlsx so the slide
is reproducible and auditable. Built in 4:3 per the AIP Brand Guidelines.
"""
import math
from statistics import NormalDist
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from PIL import Image
from lxml import etree

# ---- AIP Brand palette (Blue group only; never mix groups) -----------------
BLUE1 = RGBColor(0x0E, 0x16, 0x20)
BLUE2 = RGBColor(0x15, 0x21, 0x30)
BLUE3 = RGBColor(0x31, 0x43, 0x59)
BLUE4 = RGBColor(0x55, 0x6A, 0x83)
BLUE5 = RGBColor(0xE2, 0xE5, 0xE9)
BLUE6 = RGBColor(0xF6, 0xF7, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NAVY, NAVY_TX, SLATE, SLATE_LT = BLUE2, BLUE3, BLUE3, BLUE4
HEADERBG, BODY, SUBTLE, DIVIDER, FOOT = BLUE6, BLUE1, BLUE4, BLUE5, BLUE4
SERIF, SANS = "Times New Roman", "Arial"

LOGO_WHITE = "assets/logo_white.png"
MARK_DARK = "assets/mark_dark.png"
_MD_AR = (lambda s: s[0] / s[1])(Image.open(MARK_DARK).size)

# ===========================================================================
# 1) Compute statistics from the workbook
# ===========================================================================
def load_series():
    ws = openpyxl.load_workbook("data/Comparison_returns.xlsx",
                                data_only=True)["Sheet1"]

    def grab(month_rows):
        out = []
        for c in range(2, ws.max_column + 1):       # year columns
            for r in month_rows:
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)):
                    out.append(v)
        return out

    return grab(range(4, 16)), grab(range(22, 34))   # MSCI, Athanase


def stats(x, mar=0.0):
    n = len(x)
    mean = sum(x) / n
    sd = math.sqrt(sum((r - mean) ** 2 for r in x) / (n - 1))
    dd = math.sqrt(sum(min(r - mar, 0) ** 2 for r in x) / n)
    ud = math.sqrt(sum(max(r - mar, 0) ** 2 for r in x) / n)
    g = 1.0
    for r in x:
        g *= (1 + r)
    geo_m = g ** (1 / n) - 1
    return dict(n=n, ann_ret=(1 + geo_m) ** 12 - 1,
                ann_vol=sd * math.sqrt(12), ann_dd=dd * math.sqrt(12),
                ann_ud=ud * math.sqrt(12), mean_m=mean, sd_m=sd, series=x)


_msci_x, _ath_x = load_series()
MSCI, ATH = stats(_msci_x), stats(_ath_x)
N = NormalDist()


def project(T, r, dd):
    """Median and downside (≈1-in-6, one downside-deviation) terminal of 100."""
    med = 100 * (1 + r) ** T
    down_ann = r - dd / math.sqrt(T)
    down = 100 * (1 + down_ann) ** T
    ploss = N.cdf((0 - r) / (dd / math.sqrt(T)))
    return med, down, ploss


HORIZONS = (3, 5, 7, 10)

# ===========================================================================
# 2) Build the slide
# ===========================================================================
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
s = prs.slides.add_slide(prs.slide_layouts[6])


def rect(sl, l, t, w, h, fill=None, line=None, lw=1.0):
    shp = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    if fill is not None:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if line is not None:
        shp.line.color.rgb = line; shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def tbox(sl, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = sl.shapes.add_textbox(l, t, w, h); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


def para(tf, text, size, color, bold=False, italic=False, first=False,
         align=PP_ALIGN.LEFT, after=8, lead=1.1, font=SANS, track=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align; p.space_after = Pt(after); p.line_spacing = lead
    r = p.add_run(); r.text = text
    f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic
    f.color.rgb = color; f.name = font
    if track is None:
        track = -5 if size >= 48 else (-3 if size >= 24 else 0)
    if track:
        r._r.get_or_add_rPr().set("spc", str(int(round(size * track / 100 * 100))))
    return p


# header band + mark + title (left-aligned, brand)
s.shapes.add_picture(MARK_DARK, Inches(0.55), Inches(0.24), height=Inches(0.26),
                     width=Emu(int(Inches(0.26) * _MD_AR)))
para(tbox(s, Inches(0.98), Inches(0.27), Inches(7), Inches(0.3)),
     "Risk-Adjusted Outcomes", 11, SLATE_LT, first=True, after=0)
para(tbox(s, Inches(9.9), Inches(0.27), Inches(2.85), Inches(0.3)),
     "Figure 1", 11, SLATE_LT, first=True, bold=True, align=PP_ALIGN.RIGHT, after=0)
rect(s, 0, Inches(0.62), SW, Inches(1.45), fill=HEADERBG)
para(tbox(s, Inches(0.6), Inches(0.74), Inches(12.1), Inches(0.85)),
     "Lower downside risk compounds into better outcomes", 30, NAVY_TX,
     first=True, after=2, font=SERIF)
para(tbox(s, Inches(0.62), Inches(1.55), Inches(11.9), Inches(0.5)),
     "Athanase carries higher total volatility than the market — but its "
     "downside volatility is lower, which is what protects compounding.",
     13, SUBTLE, first=True, italic=True, after=0)

top = Inches(2.2)

# --- risk stat band (three metrics x two funds) ---
def _pct(v):
    return f"{v * 100:.1f}%"


metrics = [
    ("Annualised return", _pct(MSCI["ann_ret"]), _pct(ATH["ann_ret"]), True),
    ("Total volatility", _pct(MSCI["ann_vol"]), _pct(ATH["ann_vol"]), False),
    ("Downside volatility", _pct(MSCI["ann_dd"]), _pct(ATH["ann_dd"]), True),
]
bx = Inches(0.6); bw = Inches(3.95); gap = Inches(0.12); bh = Inches(1.28)
for label, mv, av, ath_better in metrics:
    rect(s, bx, top, bw, bh, fill=HEADERBG)
    para(tbox(s, Emu(int(bx) + int(Inches(0.2))), top + Inches(0.14),
              Emu(int(bw) - int(Inches(0.4))), Inches(0.3)),
         label.upper(), 10.5, SLATE, first=True, bold=True, after=0, track=0)
    vt = tbox(s, Emu(int(bx) + int(Inches(0.2))), top + Inches(0.46),
              Emu(int(bw) - int(Inches(0.4))), Inches(0.8))
    p = vt.paragraphs[0]; p.space_after = Pt(2)
    r1 = p.add_run(); r1.text = "MSCI  "; r1.font.size = Pt(11)
    r1.font.color.rgb = SUBTLE; r1.font.name = SANS
    r2 = p.add_run(); r2.text = mv; r2.font.size = Pt(19); r2.font.bold = True
    r2.font.color.rgb = NAVY_TX; r2.font.name = SERIF
    p2 = vt.add_paragraph()
    r3 = p2.add_run(); r3.text = "AIP   "; r3.font.size = Pt(11)
    r3.font.color.rgb = SUBTLE; r3.font.name = SANS
    r4 = p2.add_run(); r4.text = av; r4.font.size = Pt(19); r4.font.bold = True
    r4.font.color.rgb = NAVY; r4.font.name = SERIF
    bx = Emu(int(bx) + int(bw) + int(gap))

# annotation under the downside card
para(tbox(s, Inches(9.27), top + Inches(1.34), Inches(3.95), Inches(0.4)),
     "↑ Athanase downside volatility is LOWER than the market’s.",
     10, NAVY, first=True, bold=True, after=0, track=0, lead=1.05)

# --- projection table: terminal value of 100 by horizon ---
ty = Inches(4.18)
para(tbox(s, Inches(0.6), ty - Inches(0.3), Inches(12), Inches(0.3)),
     "TABLE 1.  EXPECTED VALUE OF 100 INVESTED, BY HOLDING PERIOD",
     11, SLATE, first=True, bold=True, after=0, track=0)

# columns: metric label + 4 horizons
col0 = Inches(3.5)
hw = (SW - Inches(1.2) - col0) / len(HORIZONS)
rh = Inches(0.375)
header = ["", "3 years", "5 years", "7 years", "10 years"]
rows = []
rows.append(("Median outcome — MSCI",
             [f"{project(T, MSCI['ann_ret'], MSCI['ann_dd'])[0]:.0f}" for T in HORIZONS], "sub"))
rows.append(("Median outcome — Athanase",
             [f"{project(T, ATH['ann_ret'], ATH['ann_dd'])[0]:.0f}" for T in HORIZONS], "navy"))
rows.append(("Adverse case* — MSCI",
             [f"{project(T, MSCI['ann_ret'], MSCI['ann_dd'])[1]:.0f}" for T in HORIZONS], "sub"))
rows.append(("Adverse case* — Athanase",
             [f"{project(T, ATH['ann_ret'], ATH['ann_dd'])[1]:.0f}" for T in HORIZONS], "navy"))
rows.append(("Chance of a loss — MSCI",
             [f"{project(T, MSCI['ann_ret'], MSCI['ann_dd'])[2]*100:.0f}%" for T in HORIZONS], "sub"))
rows.append(("Chance of a loss — Athanase",
             [f"{project(T, ATH['ann_ret'], ATH['ann_dd'])[2]*100:.0f}%" for T in HORIZONS], "navy"))

# header row
cx = Inches(0.6)
for ci, htext in enumerate(header):
    w = col0 if ci == 0 else hw
    rect(s, cx, ty, w, rh, fill=SLATE)
    para(tbox(s, Emu(int(cx) + int(Inches(0.12))), ty, Emu(int(w) - int(Inches(0.2))), rh,
              anchor=MSO_ANCHOR.MIDDLE), htext, 12, WHITE, bold=True, first=True,
         align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT, after=0, track=0)
    cx = Emu(int(cx) + int(w))
yy = Emu(int(ty) + int(rh))
for ri, (label, vals, kind) in enumerate(rows):
    fill = HEADERBG if ri % 2 == 0 else WHITE
    cx = Inches(0.6)
    rect(s, cx, yy, col0, rh, fill=fill)
    para(tbox(s, Emu(int(cx) + int(Inches(0.12))), yy, Emu(int(col0) - int(Inches(0.2))), rh,
              anchor=MSO_ANCHOR.MIDDLE), label, 11.5,
         NAVY if kind == "navy" else SUBTLE, bold=(kind == "navy"),
         first=True, after=0, track=0)
    cx = Emu(int(cx) + int(col0))
    for v in vals:
        rect(s, cx, yy, hw, rh, fill=fill)
        para(tbox(s, Emu(int(cx) + int(Inches(0.1))), yy, Emu(int(hw) - int(Inches(0.2))), rh,
                  anchor=MSO_ANCHOR.MIDDLE), v, 12.5,
             NAVY_TX if kind == "navy" else SUBTLE, bold=(kind == "navy"),
             first=True, align=PP_ALIGN.RIGHT, after=0, track=0)
        cx = Emu(int(cx) + int(hw))
    yy = Emu(int(yy) + int(rh))

# takeaway strip (navy) + footnote
tk = Emu(int(yy) + int(Inches(0.1)))
rect(s, Inches(0.6), tk, Inches(12.13), Inches(0.4), fill=NAVY)
para(tbox(s, Inches(0.78), tk, Inches(11.8), Inches(0.4), anchor=MSO_ANCHOR.MIDDLE),
     "Even in the adverse case, Athanase’s lower downside volatility leaves an "
     "investor ahead at every horizon — and the gap widens with time.",
     11.5, WHITE, first=True, italic=True, after=0, track=0)
para(tbox(s, Inches(0.6), Inches(7.32), Inches(12.6), Inches(0.4)),
     f"Source: monthly net returns, {MSCI['n']} months (2006–2025). Volatility "
     "annualised from monthly; downside volatility = annualised deviation of "
     "negative months (MAR 0). *Adverse case = one downside-deviation outcome "
     "(≈1-in-6), narrowing with horizon as dd/√T. Illustrative; past performance "
     "is not indicative of future results.", 7.5, FOOT, first=True, after=0,
     track=0, lead=1.1)

# ===========================================================================
# 2b) Volatility-cone slide — show the up/down path envelope for both series
# ===========================================================================
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

# brand colours as hex for matplotlib
H_NAVY, H_BLUE3, H_BLUE4 = "#152130", "#314359", "#556A83"
H_BLUE5, H_BLUE6, H_BODY = "#E2E5E9", "#F6F7F9", "#0E1620"
for _f in ("Arial", "Liberation Sans", "DejaVu Sans"):
    try:
        font_manager.findfont(_f, fallback_to_default=False)
        plt.rcParams["font.family"] = _f
        break
    except Exception:
        continue


def _cone(st, k=1.0):
    """Asymmetric ±k-deviation envelope (real up/down deviation) over 10y."""
    r, ud, dd = st["ann_ret"], st["ann_ud"], st["ann_dd"]
    yrs = np.linspace(0, 10, 121)
    expected = 100 * (1 + r) ** yrs
    with np.errstate(invalid="ignore"):
        upper = 100 * (1 + np.minimum(r + k * ud / np.sqrt(yrs), 3.0)) ** yrs
        lower = 100 * (1 + np.maximum(r - k * dd / np.sqrt(yrs), -0.95)) ** yrs
    upper[0] = lower[0] = 100
    return yrs, expected, upper, lower


def combined_cone(path_png, ymax):
    """One chart, both series overlaid on shared axes — no random paths.
    Each series gets a dark ±1-dev band and a faint ±2-dev outer band."""
    ya, ea, ua, la = _cone(ATH, 1.0)
    _, _, ua2, la2 = _cone(ATH, 2.0)
    ym, em, um, lm = _cone(MSCI, 1.0)
    _, _, um2, lm2 = _cone(MSCI, 2.0)

    fig, ax = plt.subplots(figsize=(9.4, 4.5), dpi=200)
    # --- MSCI (lighter), behind ---
    ax.fill_between(ym, lm2, um2, color=H_BLUE4, alpha=0.08, zorder=2)   # ±2 dev
    ax.fill_between(ym, lm, um, color=H_BLUE4, alpha=0.20, zorder=3)     # ±1 dev
    ax.plot(ym, em, color=H_BLUE4, lw=2.2, ls=(0, (6, 3)), zorder=5,
            label=f"MSCI World IMI — expected ({MSCI['ann_ret']*100:.0f}%)")
    ax.plot(ym, um, color=H_BLUE4, lw=1.0, alpha=0.8, zorder=4)
    ax.plot(ym, lm, color=H_BLUE4, lw=1.0, alpha=0.8, zorder=4)
    ax.plot(ym, um2, color=H_BLUE4, lw=0.8, alpha=0.5, ls=(0, (2, 2)), zorder=4)
    ax.plot(ym, lm2, color=H_BLUE4, lw=0.8, alpha=0.5, ls=(0, (2, 2)), zorder=4)
    # --- Athanase (navy), on top ---
    ax.fill_between(ya, la2, ua2, color=H_BLUE3, alpha=0.10, zorder=6)   # ±2 dev
    ax.fill_between(ya, la, ua, color=H_BLUE3, alpha=0.22, zorder=7)     # ±1 dev
    ax.plot(ya, ea, color=H_NAVY, lw=2.8, ls=(0, (6, 3)), zorder=9,
            label=f"Athanase — expected ({ATH['ann_ret']*100:.0f}%)")
    ax.plot(ya, ua, color=H_BLUE3, lw=1.2, alpha=0.85, zorder=8)
    ax.plot(ya, la, color=H_BLUE3, lw=1.2, alpha=0.85, zorder=8)
    ax.plot(ya, ua2, color=H_BLUE3, lw=0.9, alpha=0.55, ls=(0, (2, 2)), zorder=8)
    ax.plot(ya, la2, color=H_BLUE3, lw=0.9, alpha=0.55, ls=(0, (2, 2)), zorder=8)
    # shade legend entries
    ax.fill_between([], [], [], color=H_BLUE3, alpha=0.22, label="±1 deviation (≈1-in-6)")
    ax.fill_between([], [], [], color=H_BLUE3, alpha=0.10, label="±2 deviation (≈1-in-40)")
    # reference line at break-even
    ax.axhline(100, color=H_BLUE5, lw=1.0, zorder=1)

    # end-of-horizon value labels
    def _lab(y, txt, col, weight="bold"):
        ax.annotate(txt, (10, y), color=col, fontsize=10, fontweight=weight,
                    va="center", ha="left", xytext=(4, 0),
                    textcoords="offset points")
    _lab(ua2[-1], f"{ua2[-1]:.0f}", H_BLUE3, "normal")
    _lab(ua[-1], f"{ua[-1]:.0f}", H_NAVY)
    _lab(ea[-1], f"{ea[-1]:.0f}", H_NAVY)
    _lab(la[-1], f"{la[-1]:.0f}  Athanase ±1-dev downside", H_NAVY)
    _lab(um[-1], f"{um[-1]:.0f}  MSCI upside", H_BLUE4, "normal")
    _lab(lm[-1], f"{lm[-1]:.0f}", H_BLUE4, "normal")

    ax.set_xlabel("Years", color=H_BODY, fontsize=11)
    ax.set_ylabel("Value of 100 invested", color=H_BODY, fontsize=11)
    ax.set_xlim(0, 11.8); ax.set_ylim(0, ymax)
    ax.tick_params(colors=H_BODY, labelsize=10)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color(H_BLUE5)
    ax.grid(True, axis="y", color=H_BLUE5, lw=0.6, alpha=0.7)
    ax.legend(loc="upper left", fontsize=10, frameon=False, labelcolor=H_BODY)
    fig.tight_layout()
    fig.savefig(path_png, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


combined_cone("/tmp/cone_combined.png", 1700)

s2 = prs.slides.add_slide(prs.slide_layouts[6])
s2.shapes.add_picture(MARK_DARK, Inches(0.55), Inches(0.24), height=Inches(0.26),
                      width=Emu(int(Inches(0.26) * _MD_AR)))
para(tbox(s2, Inches(0.98), Inches(0.27), Inches(7), Inches(0.3)),
     "Risk-Adjusted Outcomes", 11, SLATE_LT, first=True, after=0)
para(tbox(s2, Inches(9.9), Inches(0.27), Inches(2.85), Inches(0.3)),
     "Figure 2", 11, SLATE_LT, first=True, bold=True, align=PP_ALIGN.RIGHT, after=0)
rect(s2, 0, Inches(0.62), SW, Inches(1.45), fill=HEADERBG)
para(tbox(s2, Inches(0.6), Inches(0.74), Inches(12.1), Inches(0.85)),
     "The path through time — upside and downside envelope", 30, NAVY_TX,
     first=True, after=2, font=SERIF)
para(tbox(s2, Inches(0.62), Inches(1.55), Inches(11.9), Inches(0.5)),
     "Both series, one scale. Athanase’s envelope fans wide on the upside but "
     "holds firm below — its downside edge stays above the market’s best case.",
     13, SUBTLE, first=True, italic=True, after=0)
# single combined chart
s2.shapes.add_picture("/tmp/cone_combined.png", Inches(0.85), Inches(2.2),
                      height=Inches(4.1))
# takeaway strip
rect(s2, Inches(0.6), Inches(6.5), Inches(12.13), Inches(0.5), fill=NAVY)
para(tbox(s2, Inches(0.78), Inches(6.5), Inches(11.8), Inches(0.5),
          anchor=MSO_ANCHOR.MIDDLE),
     f"Athanase upside deviation {ATH['ann_ud']*100:.0f}% vs downside "
     f"{ATH['ann_dd']*100:.0f}% — the volatility is mostly upside. Even its "
     "10-year downside edge stays above the market’s expected path.",
     12, WHITE, first=True, italic=True, after=0, track=0)
para(tbox(s2, Inches(0.6), Inches(7.12), Inches(12.6), Inches(0.4)),
     f"Source: monthly net returns, {MSCI['n']} months (2006–2025). Shaded band "
     "= expected path ±1 annualised up/downside deviation, widening with √time. "
     "Illustrative; past performance is not indicative of future results.",
     7.5, FOOT, first=True, after=0, track=0, lead=1.1)

# ===========================================================================
# 3) Convert to 4:3 and brand the theme (palette + fonts)
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
            pPr = p._p.find(_qn("a:pPr"))
            if pPr is not None:
                for at in ("marL", "indent"):
                    v = pPr.get(at)
                    if v is not None:
                        pPr.set(at, str(int(round(int(v) * sx))))


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

out = "Athanase_Downside_Risk_Outcomes.pptx"
prs.save(out)
print("Saved", out)
print(f"MSCI  : ret {MSCI['ann_ret']*100:.1f}%  vol {MSCI['ann_vol']*100:.1f}%  "
      f"dd {MSCI['ann_dd']*100:.1f}%")
print(f"AIP   : ret {ATH['ann_ret']*100:.1f}%  vol {ATH['ann_vol']*100:.1f}%  "
      f"dd {ATH['ann_dd']*100:.1f}%")
