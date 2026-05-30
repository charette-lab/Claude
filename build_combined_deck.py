"""Build a combined Athanase deck:
   Part I  - Why allocate to engaged ownership (activist) as an asset class
   Part II - Why choose Athanase among engaged owners
Styled to match the Athanase Industrial Partner brand.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

# ---- Brand palette (AIP Brand Guidelines v1.0 — BLUE group only) -----------
# Rule: use one base colour group together; never mix groups. Blue is primary.
BLUE1 = RGBColor(0x0E, 0x16, 0x20)
BLUE2 = RGBColor(0x15, 0x21, 0x30)   # primary logo / dark navy
BLUE3 = RGBColor(0x31, 0x43, 0x59)
BLUE4 = RGBColor(0x55, 0x6A, 0x83)
BLUE5 = RGBColor(0xE2, 0xE5, 0xE9)
BLUE6 = RGBColor(0xF6, 0xF7, 0xF9)

NAVY      = BLUE2   # cover / divider background, dark strips, dark bars
NAVY_TX   = BLUE3   # titles on white, table-header fills, lead/emphasis text
SLATE     = BLUE3   # primary accent: table headers, lead bullets, captions
SLATE_LT  = BLUE4   # section labels, secondary accent, lighter bars/lines
HEADERBG  = BLUE6   # light header band / card fill
BODY      = BLUE1   # body text (near-black with blue hint, in-group)
SUBTLE    = BLUE4   # subtitles / secondary
DIVIDER   = BLUE5   # light title text on navy
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
FOOT      = BLUE4   # footer
GOLD      = BLUE4   # accent rule
# Brand palette is monochrome Blue — no red. Losses are de-emphasised in a
# lighter in-group Blue (BLUE4); the negative sign carries the meaning.
LOSS      = BLUE4
HILITE    = BLUE3   # in-group highlight for callouts (was red)

SERIF = "Times New Roman"   # brand headline face
SANS  = "Arial"             # brand body face

import os
from PIL import Image
LOGO_WHITE = "assets/logo_white.png"
MARK_DARK  = "assets/mark_dark.png"
_LW_AR = (lambda s: s[0] / s[1])(Image.open(LOGO_WHITE).size)
_MD_AR = (lambda s: s[0] / s[1])(Image.open(MARK_DARK).size)


def place_logo_white(slide, left, top, height):
    slide.shapes.add_picture(LOGO_WHITE, left, top, height=height,
                             width=Emu(int(int(height) * _LW_AR)))


def place_mark(slide, left, top, height):
    slide.shapes.add_picture(MARK_DARK, left, top, height=height,
                             width=Emu(int(int(height) * _MD_AR)))


# ---- Track-record data (from data/AIP_Trackrecord.xlsx, Transactions tab) ---
import openpyxl


def _fmt_pct(v):
    if not isinstance(v, (int, float)) or abs(v) > 3.0:
        return "n.m."
    return f"{v * 100:+.0f}%"


def _fmt_irr(v):
    if not isinstance(v, (int, float)) or v > 3.0:
        return "n.m."
    return f"{v * 100:.0f}%"


def _fmt_moic(v):
    return f"{v:.1f}x" if isinstance(v, (int, float)) else str(v)


def load_transactions():
    wb = openpyxl.load_workbook("data/AIP_Trackrecord.xlsx", data_only=True)
    ws = wb["Transactions"]

    def grab(r0, r1):
        deals = []
        for r in range(r0, r1 + 1):
            comp = ws.cell(r, 3).value          # C
            if not comp:
                continue
            deals.append(dict(
                period=str(ws.cell(r, 2).value or "").strip(),   # B
                company=str(comp).strip(),
                irr=ws.cell(r, 6).value,         # F
                outp=ws.cell(r, 8).value,        # H
                moic=ws.cell(r, 11).value,       # K
            ))
        return deals

    fund2 = grab(4, 20)         # AIP Fund II (2015-2026)
    hist = grab(27, 48)         # Prior period (2006-2014)
    return fund2, hist


def deal_table(slide, x, y, deals, col_w, font=10.5, rh=Inches(0.265),
               cols=("company", "period", "irr", "moic", "outp"),
               headers=("Company", "Holding", "IRR", "MOIC", "vs Index")):
    """Compact transactions table with header + alternating rows."""
    # header
    cx = x
    for ci, htext in enumerate(headers):
        rect(slide, cx, y, col_w[ci], rh, fill=SLATE)
        htf = tbox(slide, Emu(int(cx) + int(Inches(0.08))), y,
                   Emu(int(col_w[ci]) - int(Inches(0.12))), rh,
                   anchor=MSO_ANCHOR.MIDDLE)
        para(htf, htext, font, WHITE, bold=True, first=True,
             align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT, after=0)
        cx = Emu(int(cx) + int(col_w[ci]))
    yy = Emu(int(y) + int(rh))
    for ri, d in enumerate(deals):
        loss = isinstance(d["moic"], (int, float)) and d["moic"] < 1.0
        fill = HEADERBG if ri % 2 == 0 else WHITE
        cx = x
        for ci, key in enumerate(cols):
            rect(slide, cx, yy, col_w[ci], rh, fill=fill)
            if key == "company":
                txt = d["company"]
            elif key == "period":
                txt = d["period"]
            elif key == "irr":
                txt = _fmt_irr(d["irr"])
            elif key == "moic":
                txt = _fmt_moic(d["moic"])
            else:
                txt = _fmt_pct(d["outp"])
            is_loss_cell = loss and key in ("irr", "moic", "outp")
            col = LOSS if is_loss_cell else (NAVY_TX if ci == 0 else BODY)
            ctf = tbox(slide, Emu(int(cx) + int(Inches(0.08))), yy,
                       Emu(int(col_w[ci]) - int(Inches(0.12))), rh,
                       anchor=MSO_ANCHOR.MIDDLE)
            # In-palette loss treatment: lighter Blue + italic (brand allows
            # italic for emphasis); the negative sign carries the meaning.
            para(ctf, txt, font, col, bold=(ci == 0), italic=is_loss_cell,
                 first=True, align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT,
                 after=0, track=0)
            cx = Emu(int(cx) + int(col_w[ci]))
        yy = Emu(int(yy) + int(rh))
    return yy

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

_state = {"n": 0}


def rect(slide, l, t, w, h, fill=None, line=None, line_w=1.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    if fill is not None:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if line is not None:
        shp.line.color.rgb = line; shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def tbox(slide, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0; tf.margin_bottom = 0
    return tf


def para(tf, text, size, color, bold=False, italic=False, first=False,
         align=PP_ALIGN.LEFT, after=8, lead=1.1, font=SANS, track=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(after)
    p.line_spacing = lead
    r = p.add_run(); r.text = text
    f = r.font
    f.size = Pt(size); f.bold = bold; f.italic = italic
    f.color.rgb = color; f.name = font
    # Brand type scale: letter-spacing (track in % of em, e.g. -5 => -0.05em).
    # Auto-apply brand defaults by size if not given: 48pt+ -> -5%, 24-47 -> -3%.
    if track is None:
        track = -5 if size >= 48 else (-3 if size >= 24 else 0)
    if track:
        spc = int(round(size * track / 100.0 * 100))  # EMU-less: points*100
        r._r.get_or_add_rPr().set("spc", str(spc))
    return p


def wordmark(slide, left, top, color=None, scale=1.0):
    """Place the real white Athanase wordmark (for navy backgrounds)."""
    place_logo_white(slide, left, top, Inches(0.52 * scale))


def footer(slide):
    _state["n"] += 1
    tf = tbox(slide, Inches(9.0), Inches(7.05), Inches(4.0), Inches(0.3))
    para(tf, f"Strictly confidential          {_state['n']}", 8.5, FOOT,
         first=True, align=PP_ALIGN.RIGHT, after=0)


# --- research-paper numbering: sections, figures, tables -------------------
_doc = {"part": 0, "sec": 0}
_exhibit = {"fig": 0, "tbl": 0}
TOC = []   # (number, title, ref) for the Contents page


def fig():
    _exhibit["fig"] += 1
    return f'Figure {_exhibit["fig"]}'


def tbl():
    _exhibit["tbl"] += 1
    return f'Table {_exhibit["tbl"]}'


def dlabels(holder, fmt='0.0"%"', size=9, color=None, pos=None):
    """Show numeric data labels on a chart plot or series."""
    holder.has_data_labels = True
    dl = holder.data_labels
    dl.number_format = fmt
    dl.number_format_is_linked = False
    dl.font.size = Pt(size)
    dl.font.color.rgb = color if color is not None else BODY
    if pos is not None:
        dl.position = pos


def divider(title, kicker=None):
    _doc["part"] += 1
    _doc["sec"] = 0
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=NAVY)
    wordmark(s, Inches(0.6), Inches(0.55), WHITE, scale=0.9)
    # brand rule: always left-align (never centre)
    rect(s, Inches(0.62), Inches(3.05), Inches(0.5), Pt(2.2), fill=SLATE_LT)
    if kicker:
        ktf = tbox(s, Inches(0.6), Inches(3.25), Inches(11.7), Inches(0.5))
        para(ktf, kicker.upper(), 13, SLATE_LT, first=True,
             align=PP_ALIGN.LEFT, after=0)
    tf = tbox(s, Inches(0.6), Inches(3.6), Inches(11.9), Inches(1.4))
    para(tf, title, 40, DIVIDER, italic=True, first=True,
         align=PP_ALIGN.LEFT, after=0, font=SERIF)
    return s


def content(section, title, subtitle=None, number=True, ref=None):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=WHITE)
    # top section label + real mark
    place_mark(s, Inches(0.55), Inches(0.24), Inches(0.26))
    lt = tbox(s, Inches(0.98), Inches(0.27), Inches(7.0), Inches(0.3))
    para(lt, section, 11, SLATE_LT, first=True, after=0)
    # top-right exhibit reference (Figure N / Table N)
    if ref:
        rtf = tbox(s, Inches(9.9), Inches(0.27), Inches(2.85), Inches(0.3))
        para(rtf, ref, 11, SLATE_LT, first=True, bold=True,
             align=PP_ALIGN.RIGHT, after=0)
    # numbered section header, research-paper style
    if number and _doc["part"] >= 1:
        _doc["sec"] += 1
        TOC.append((f'{_doc["part"]}.{_doc["sec"]}', title, ref))
        title = f'{_doc["part"]}.{_doc["sec"]} {title}'
    # grey header band
    band_h = Inches(1.45) if subtitle else Inches(1.05)
    rect(s, 0, Inches(0.62), SW, band_h, fill=HEADERBG)
    tt = tbox(s, Inches(0.6), Inches(0.74), Inches(12.1), Inches(0.85))
    para(tt, title, 30, NAVY_TX, first=True, after=2, font=SERIF)
    body_top = Inches(1.95)
    if subtitle:
        st = tbox(s, Inches(0.62), Inches(1.55), Inches(11.9), Inches(0.5))
        para(st, subtitle, 13, SUBTLE, first=True, italic=True, after=0)
        body_top = Inches(2.35)
    footer(s)
    return s, body_top


def checklist(s, items, top, left=Inches(0.75), width=Inches(11.9),
              size=16, gap=14, lead_bold=True):
    tf = tbox(s, left, top, width, Inches(4.7))
    for i, it in enumerate(items):
        if isinstance(it, tuple):
            lead, body = it
        else:
            lead, body = it, None
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.line_spacing = 1.12
        r0 = p.add_run(); r0.text = "›  "
        r0.font.size = Pt(size); r0.font.bold = True
        r0.font.color.rgb = SLATE; r0.font.name = SANS
        r1 = p.add_run(); r1.text = lead
        r1.font.size = Pt(size); r1.font.bold = lead_bold
        r1.font.color.rgb = NAVY_TX; r1.font.name = SANS
        if body:
            r2 = p.add_run(); r2.text = "  " + body
            r2.font.size = Pt(size); r2.font.color.rgb = BODY; r2.font.name = SANS
    return tf


# ===========================================================================
# COVER
# ===========================================================================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, fill=NAVY)
wordmark(s, Inches(0.6), Inches(0.55), WHITE, scale=1.0)
rect(s, 0, Inches(4.5), Inches(0.5), Pt(2.2), fill=SLATE_LT)
tf = tbox(s, Inches(0.6), Inches(4.75), Inches(11.8), Inches(2.0))
para(tf, "Engaged Ownership as an Asset Class", 42, WHITE, bold=False,
     first=True, after=6, font=SERIF)
para(tf, "Why allocators should consider it — and why Athanase", 24, DIVIDER,
     italic=True, after=0, font=SERIF)
tf2 = tbox(s, Inches(0.62), Inches(6.6), Inches(8), Inches(0.5))
para(tf2, "An allocator briefing  ·  May 2026", 13, SLATE_LT, first=True, after=0)

# ===========================================================================
# AGENDA
# ===========================================================================
agenda_s, agenda_top = content("Overview", "Contents")
# TOC body rendered at end, once all sections are numbered.

# ===========================================================================
# PART I DIVIDER
# ===========================================================================
divider("Why allocate to engaged ownership", kicker="Part I  ·  The asset class")

# ---- I.1 Return & alpha ----
s, top = content("The Proposition",
                 "A catalyst-driven source of return",
                 "Engaged owners change the outcome — they do not merely absorb "
                 "underperformance.")
checklist(s, [
    ("~7% announcement return.", "Average abnormal return around an activist "
     "filing, with no reversal over the following year; success or partial "
     "success in ~two-thirds of campaigns (Brav, Jiang, Partnoy & Thomas, 2008)."),
    ("~10.2% around the 13D,", "plus a further ~11.4% over the next year for "
     "engaged-owner targets (Klein & Zur, 2009)."),
    ("No long-term give-back.", "~2,000 interventions show no negative drift and "
     "no “pump-and-dump” (Bebchuk, Brav & Jiang, 2015)."),
    ("Real, not financial, value.", "Production efficiency improves in the three "
     "years after intervention (Brav, Jiang & Kim, 2015)."),
    ("Idiosyncratic alpha.", "Returns tied to a specific change at a specific "
     "company — scarce in a beta-dominated portfolio."),
], top, size=15.5, gap=12)

# ---- I.2 Horizon & fiduciary ----
s, top = content("Fit", "Aligned with horizon and fiduciary duty")
checklist(s, [
    ("Patient capital wins.", "Board seats, strategic reviews and turnarounds "
     "play out over years — perpetual endowments and long-dated pensions can "
     "hold through to full value realisation."),
    ("Stewardship is the duty.", "Engaged ownership expresses fiduciary "
     "obligation (UK Stewardship Code; UN PRI) rather than conflicting with it."),
    ("Mission alignment.", "Demanding accountability and good governance reflects "
     "institutional values and is defensible to beneficiaries."),
    ("Engagement as risk management.", "Turns hidden governance risk in the "
     "portfolio into a remediable, return-generating exposure."),
], top, size=16.5, gap=16)

# ---- I.3 Diversification ----
s, top = content("Portfolio", "Diversification and the total-portfolio view")
checklist(s, [
    ("A distinct return engine.", "Alpha driven by proxy contests, spin-offs and "
     "capital-return programmes — largely independent of rates, growth and "
     "inflation."),
    ("Diversifies the equity book.", "Small- and mid-cap value and special "
     "situations, typically underweight in institutional portfolios."),
    ("Counter-cyclical dry powder.", "Dislocations create the next targets; "
     "patient capital deploys into weakness rather than selling it."),
    ("PE-style upside, public-market terms.", "Operational engagement with daily "
     "pricing, transparency, lower fees and better liquidity."),
], top, size=16.5, gap=16)

# ---- I.4 Governance / ESG ----
s, top = content("Stewardship", "Governance and the ESG synergy")
checklist(s, [
    ("Applied governance.", "Board independence, pay-for-performance and capital "
     "discipline are the “G” of ESG put into practice."),
    ("Constructive engagement.", "Return and stewardship objectives expressed "
     "through one allocation, disciplined by a value thesis."),
    ("A positive externality.", "Governance precedents improve the quality of the "
     "beta the whole portfolio owns."),
    ("Defensible to stakeholders.", "“We own companies and make them better-"
     "run and more valuable” is an easy story to tell."),
], top, size=16.5, gap=16)

# ---- I.5 What to look for (bridge) ----
s, top = content("Selection", "What separates a great engaged owner",
                 "Dispersion is wide — manager selection matters more here "
                 "than in almost any other equity strategy.")
checklist(s, [
    ("A repeatable process", "for sourcing and underwriting correctable, "
     "undervalued companies."),
    ("Genuine operating capability,", "not just financial engineering — the "
     "experience to control the outcome from the boardroom."),
    ("A track record of outcomes,", "board seats won and value delivered across "
     "multiple market cycles."),
    ("Aligned economics", "and a stable, long-tenured team that invests its own "
     "capital."),
    ("Disciplined entry and risk control", "so the strategy delivers without "
     "uncompensated concentration risk."),
], top, size=15.5, gap=12)
# bridge line
bt = tbox(s, Inches(0.75), Inches(6.45), Inches(11.8), Inches(0.5))
para(bt, "Part II shows how Athanase meets each of these criteria.", 14,
     SLATE, first=True, italic=True, after=0)

# ===========================================================================
# PART II DIVIDER
# ===========================================================================
divider("Why Athanase", kicker="Part II  ·  Among engaged owners")

# ---- II.1 At a glance ----
s, top = content("Overview", "Athanase at a glance")
# four big stats
stats = [("40+", "investments as\nengaged owners"),
         ("18%", "net annual IRR\nsince inception"),
         ("20 yrs", "team working\ntogether"),
         ("7.1%", "minimum annual\nnet return*")]
bw = Inches(2.95); gapx = Inches(0.18); x = Inches(0.75); y = top
for big, lab in stats:
    card = rect(s, x, y, bw, Inches(1.85), fill=HEADERBG)
    ctf = tbox(s, x, y + Inches(0.28), bw, Inches(1.4), anchor=MSO_ANCHOR.TOP)
    para(ctf, big, 40, NAVY_TX, first=True, align=PP_ALIGN.CENTER, after=4,
         font=SERIF)
    for j, line in enumerate(lab.split("\n")):
        para(ctf, line, 12, SUBTLE, align=PP_ALIGN.CENTER, after=0)
    x = Emu(int(x) + int(bw) + int(gapx))
checklist(s, [
    ("Average net return of 18.6%", "(17.1% vs MSCI IMI 7.6% since inception), "
     "regardless of entry month."),
    ("+14.5% EBITA growth", "in portfolio companies during the ownership period."),
    ("+0.2% in down markets", "versus the market at –1.7% — downside "
     "protection when it matters most."),
], top + Inches(2.15), size=15, gap=10)
nt = tbox(s, Inches(0.75), Inches(6.75), Inches(11.8), Inches(0.4))
para(nt, "*Invested any month since 2006 and held to Dec 2025. Inception 2015-02-23.",
     9, FOOT, first=True, after=0)

# ---- II.2 The team ----
s, top = content("Experience", "An integrated team, 20 years together",
                 "In a world of specialists, a rare unit that has the operating "
                 "experience to control the outcome.")
# heritage timeline visual
nodes = [("1992", "Öresund"), ("1996", "Custos"),
         ("2006", "Creades"), ("2015", "Athanase")]
tl_y = Emu(int(top) + int(Inches(0.42)))           # the line
x0, x1 = Inches(1.6), Inches(11.5)
rect(s, x0, tl_y, Emu(int(x1) - int(x0)), Pt(1.6), fill=SLATE_LT)
n = len(nodes)
for i, (yr, name) in enumerate(nodes):
    cx = Emu(int(x0) + int((int(x1) - int(x0)) * i / (n - 1)))
    # year above
    yt = tbox(s, Emu(int(cx) - int(Inches(1.0))), Emu(int(tl_y) - int(Inches(0.42))),
              Inches(2.0), Inches(0.35))
    para(yt, yr, 15, NAVY_TX, first=True, bold=True, align=PP_ALIGN.CENTER,
         after=0, font=SERIF)
    # node dot
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(cx) - int(Inches(0.09))),
                             Emu(int(tl_y) - int(Inches(0.07))),
                             Inches(0.18), Inches(0.18))
    dot.fill.solid(); dot.fill.fore_color.rgb = NAVY
    dot.line.color.rgb = WHITE; dot.line.width = Pt(1.5); dot.shadow.inherit = False
    # name below
    nt = tbox(s, Emu(int(cx) - int(Inches(1.1))), Emu(int(tl_y) + int(Inches(0.16))),
              Inches(2.2), Inches(0.3))
    para(nt, name, 12, SLATE, first=True, align=PP_ALIGN.CENTER, after=0)
checklist(s, [
    ("40+ companies managed,", "with a proven ability to tell a “strong "
     "core” from a “prestige trap”."),
    ("30+ public board seats", "— real experience mandating pivots and "
     "protecting shareholder interests, not writing engagement letters."),
    ("A dual lens:", "a forensic investing lens to find the value and a C-suite "
     "operating lens that understands the “how” of execution."),
    ("Skin in the game.", "The team uses Athanase as its own investment vehicle "
     "and shares in the success economics."),
    ("Institutional memory.", "Multiple market cycles together eliminate "
     "key-person risk and prevent repeating historical mistakes."),
], top + Inches(1.25), size=14.5, gap=10)

# ---- II.3 Market opportunity ----
_f = fig()
s, top = content("Market Opportunity", "The inefficiency Athanase harvests",
                 "Corporate decay is structural — it continuously recreates "
                 "low-risk, high-return entry points.", ref=_f)
checklist(s, [
    ("Duration mismatch.", "CEOs run 7–8 year defensive cycles while moats "
     "erode at hyper-speed; an 80%-correct strategy decays to ~50% efficient."),
    ("The prestige trap.", "Empire-building and the sunk-cost fallacy lead "
     "management to double down on unprofitable ventures."),
    ("Shareholder exhaustion.", "As patience runs out, price hits a long-term low "
     "and liquidity dries up — Athanase becomes the “buyer of last "
     "resort” at a discount."),
    ("Low competition.", "Passive buys the most expensive names; active managers "
     "wait for proof; PE needs 10x the capital — leaving these bargains open."),
], top, left=Inches(0.7), width=Inches(6.3), size=13.5, gap=14)
# probability-decay curve (right)
dd = CategoryChartData()
dd.categories = ["Yr 1", "Yr 2", "Yr 3", "Yr 4", "Yr 5", "Yr 6", "Yr 7", "Yr 8"]
dd.add_series("Strategy efficiency", (80, 76, 71, 65, 60, 55, 52, 50))
gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(7.15), top + Inches(0.2),
                        Inches(5.7), Inches(3.6), dd)
ch = gf.chart
ch.has_title = True
ch.chart_title.text_frame.text = f"{_f}.  “Probability decay”: strategy efficiency over a CEO cycle (%)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(11)
ch.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb = SUBTLE
ch.has_legend = False
sr = ch.plots[0].series[0]
sr.format.line.color.rgb = NAVY
sr.format.line.width = Pt(2.5)
sr.smooth = True
dlabels(sr, fmt='0"%"', size=9, color=NAVY_TX, pos=XL_LABEL_POSITION.ABOVE)
ch.category_axis.tick_labels.font.size = Pt(9)
ch.value_axis.tick_labels.font.size = Pt(9)
ch.value_axis.has_major_gridlines = False
ch.value_axis.minimum_scale = 0
ch.value_axis.maximum_scale = 100
para(tbox(s, Inches(7.15), top + Inches(3.85), Inches(5.7), Inches(0.4)),
     "Even an excellent team’s edge erodes as the external environment shifts.",
     10, FOOT, first=True, italic=True, after=0)

# ---- II.3b Negative-selection rebuttal: hidden high-quality core ----
_f2 = fig()
s, top = content("Investment Strategy",
                 "We don’t buy broken companies — we buy hidden cores",
                 "The classic objection to activism is adverse selection. Our "
                 "model inverts it.", ref=_f2)
checklist(s, [
    ("The critique.", "Buy “cheap and troubled” and you select for lemons — "
     "worse than they look. The market discounted them for a reason."),
    ("Our precondition removes it.", "We invest only where the core is a market "
     "leader earning high returns on capital — attractive even if we change "
     "nothing. No structurally-weak turnarounds."),
    ("The real setup.", "Profitable leaders whose management pivoted into weak "
     "adjacencies — chasing growth as the core matured. A capital-allocation "
     "overlay on a good asset, not a bad asset."),
    ("The inversion.", "Blended results make it look mediocre, so it is priced "
     "mediocre. The appearance understates reality."),
], top, left=Inches(0.7), width=Inches(6.4), size=13.5, gap=13)

# --- sum-of-the-parts bar visual (right) ---
ptf = tbox(s, Inches(7.3), top - Inches(0.05), Inches(5.4), Inches(0.3))
para(ptf, f"{_f2}  ·  SUM-OF-THE-PARTS (ILLUSTRATIVE, INDEXED)", 11, SLATE,
     first=True, bold=True, align=PP_ALIGN.CENTER, after=0)
base = Inches(5.5)
unit = 0.017

def _sotp_bar(xc, val, fill, label):
    h = Emu(int(Inches(unit * val)))
    rect(s, Emu(int(xc) - int(Inches(0.8))), Emu(int(base) - int(h)),
         Inches(1.6), h, fill=fill)
    vt = tbox(s, Emu(int(xc) - int(Inches(0.8))),
              Emu(int(base) - int(h) - int(Inches(0.36))), Inches(1.6), Inches(0.32))
    para(vt, str(val), 14, NAVY_TX, first=True, bold=True,
         align=PP_ALIGN.CENTER, after=0)
    lt = tbox(s, Emu(int(xc) - int(Inches(1.2))),
              Emu(int(base) + int(Inches(0.1))), Inches(2.4), Inches(0.6))
    for k, ln in enumerate(label.split("\n")):
        para(lt, ln, 11.5 if k == 0 else 10, BODY, bold=(k == 0),
             first=(k == 0), align=PP_ALIGN.CENTER, after=0, lead=1.0)

_sotp_bar(Inches(8.8), 100, SLATE, "Market price\n(blended optics)")
_sotp_bar(Inches(11.4), 140, NAVY, "Core alone\n(high ROIIC)")
top100 = Emu(int(base) - int(Inches(unit * 100)))
top140 = Emu(int(base) - int(Inches(unit * 140)))
rect(s, Inches(8.0), top100, Inches(3.4), Pt(1.2), fill=SLATE_LT)  # market level
rect(s, Inches(11.95), top140, Pt(2), Emu(int(top100) - int(top140)), fill=HILITE)
gtf = tbox(s, Inches(8.15), Emu(int(top140) + int(Inches(0.04))),
           Inches(3.6), Inches(0.5))
para(gtf, "+40% hidden value", 12, HILITE, first=True, bold=True,
     align=PP_ALIGN.LEFT, after=0)
para(tbox(s, Inches(0.7), Inches(6.45), Inches(12.0), Inches(0.5)),
     "We remove the drag and refocus capital on the core — positive selection on "
     "a concealed asset, not negative selection on a lemon.", 13, SLATE,
     first=True, italic=True, after=0)

# ---- II.3c How Athanase differs: quality-core constructivism vs Elliott ----
s, top = content("Investment Strategy",
                 "Quality-core constructivism — how we differ",
                 "Even Hohn (TCI) and Loeb (Third Point) pivoted to quality — we "
                 "capture quality at a discount, with a catalyst in our control.",
                 ref=tbl())
rows = [
    ("", "Large generalist activists (e.g. Elliott)",
     "Athanase — quality-core constructivism"),
    ("Entry precondition",
     "Wide range of situations — breakups, credit, event-driven; business "
     "quality often secondary",
     "Only a market-leading, high-ROIIC core — “attractive even if we change "
     "nothing”"),
    ("Source of return",
     "Financial / transactional pressure: force a sale, return cash, exit",
     "Operational refocus of a genuinely good business, from a board seat, over "
     "years"),
    ("What gets fixed",
     "The price, the structure, the outcome of a contest",
     "The capital-allocation overlay — the core already works"),
    ("Margin of safety",
     "Winning the campaign",
     "The core’s proven high-ROIIC economics — the hard floor"),
    ("Cadence & style",
     "Many simultaneous campaigns, leverage, often confrontational",
     "~one deal a year, constructive, board-led, mid-cap surgical"),
]
tl = Inches(0.55); cw3 = [Inches(2.35), Inches(4.85), Inches(4.95)]
rh0 = Inches(0.42); rhd = Inches(0.74); y = top
for ri, row in enumerate(rows):
    head = ri == 0
    rh = rh0 if head else rhd
    x = tl
    for ci, cell in enumerate(row):
        if head:
            fill = SLATE if ci else WHITE
        else:
            fill = HEADERBG if ri % 2 else WHITE
        rect(s, x, y, cw3[ci], rh, fill=fill)
        ctf = tbox(s, Emu(int(x) + int(Inches(0.12))), y,
                   Emu(int(cw3[ci]) - int(Inches(0.2))), rh,
                   anchor=MSO_ANCHOR.MIDDLE)
        if head:
            col = WHITE
        elif ci == 0:
            col = SLATE
        elif ci == 2:
            col = NAVY_TX
        else:
            col = BODY
        para(ctf, cell, 12 if head else 11.5, col,
             bold=(head or ci == 0 or ci == 2), first=True, after=0, lead=1.04)
        x = Emu(int(x) + int(cw3[ci]))
    y = Emu(int(y) + int(rh))
para(tbox(s, Inches(0.55), Emu(int(y) + int(Inches(0.14))), Inches(12.2), Inches(0.5)),
     "The megacap activists left mid-cap surgical engagements behind as they "
     "scaled. That hidden-core niche is our lane.", 13, SLATE, first=True,
     italic=True, after=0)

# ---- II.4 Margin for error vs PE (table) ----
s, top = content("Investment Strategy",
                 "We don’t need to be right more than ~50%",
                 "Because we buy at a discount, not at a 40% control premium.",
                 ref=tbl())
# table
rows = [
    ("Metric", "Private Equity (take-private)", "Athanase (engaged owner)"),
    ("Entry basis", "$140  (market + 40% premium)", "$100  (exhaustion price)"),
    ("Holding period", "5 years", "3 years (rerating cycle)"),
    ("Value growth required", "+148% total growth", "+73% total growth"),
    ("Operational success needed", "~85–100% of thesis", "~45–55% of thesis"),
]
tbl_l = Inches(0.75); tbl_t = top; col_w = [Inches(4.0), Inches(4.0), Inches(3.85)]
row_h = Inches(0.62)
y = tbl_t
for ri, row in enumerate(rows):
    x = tbl_l
    head = (ri == 0)
    for ci, cell in enumerate(row):
        fill = SLATE if head else (HEADERBG if ri % 2 == 0 else WHITE)
        cellrect = rect(s, x, y, col_w[ci], row_h, fill=fill)
        ctf = tbox(s, Emu(int(x) + int(Inches(0.15))), y,
                   Emu(int(col_w[ci]) - int(Inches(0.3))), row_h,
                   anchor=MSO_ANCHOR.MIDDLE)
        emph = head or ci == 2
        para(ctf, cell, 13 if not head else 13.5,
             WHITE if head else (NAVY_TX if ci == 2 else BODY),
             bold=emph, first=True, after=0)
        x = Emu(int(x) + int(col_w[ci]))
    y = Emu(int(y) + int(row_h))
para(tbox(s, Inches(0.75), Emu(int(y) + int(Inches(0.2))), Inches(11.8), Inches(0.5)),
     "A margin-for-error strategy: we win by being directionally right where PE "
     "must be precisely perfect.", 14, SLATE, first=True, italic=True, after=0)

# ---- II.5 Risk system ----
s, top = content("Risk Management", "A risk system built to avoid permanent loss",
                 "Predictability, price and influence combined to lower risk "
                 "systematically.")
# two columns
colL = tbox(s, Inches(0.75), top, Inches(5.85), Inches(4.5))
para(colL, "THE VALUATION FLOOR", 13, SLATE, first=True, bold=True, after=6)
for t in ["Enter only at “shareholder surrender” — dried-up "
          "liquidity and long-term price lows",
          "Value only the rectifiable core; assign zero to “growth trap” "
          "divisions",
          "Creates a structural 30–40% margin of safety vs a PE bidder"]:
    para(colL, t, 14, BODY, after=9, lead=1.12)
colR = tbox(s, Inches(7.0), top, Inches(5.8), Inches(4.5))
para(colR, "THE BOARD-LEVEL KILL SWITCH", 13, SLATE, first=True, bold=True, after=6)
for t in ["Secure board seats and chairmanships as a prerequisite — not "
          "proxy threats",
          "Mandated rectification: freeze prestige projects the moment the macro "
          "shifts",
          "Automatic risk metrics (10/20/30% triggers) remove portfolio-manager "
          "bias"]:
    para(colR, t, 14, BODY, after=9, lead=1.12)
# core & satellite line
para(tbox(s, Inches(0.75), Inches(6.35), Inches(11.8), Inches(0.6)),
     "Core & satellite: a 30-stock equal-weighted core index compounds capital; "
     "8–12 concentrated ideas must clear a 12% IRR hurdle and a ≤20% "
     "downside gate.", 13, SUBTLE, first=True, italic=True, after=0)

# ---- II.5b Strategic reframing: the discomfort is the moat ----
s, top = content("Strategic Case", "The visible risks are the engine — not the cost",
                 "From “corporate raider” to “industrial constructivist”: "
                 "reputational friction and the public J-curve source the alpha "
                 "and protect it from arbitrage.")
checklist(s, [
    ("Friction is the catalyst.", "Quiet engagement does not reprice a stock; "
     "public engagement forces the market to re-rate — and a 20-year reputation "
     "pulls sophisticated capital in behind us (the coattail effect)."),
    ("The discomfort is the moat.", "Few managers have the industrial background "
     "and stamina to wage a multi-year campaign — so headline-shy capital leaves "
     "a persistent pricing inefficiency for those who will."),
    ("A board seat de-risks the asset.", "We trade optical, mark-to-market "
     "volatility for direct control of cash flow, capex and pay. Fundamental risk "
     "falls the moment control is secured — whatever the share price does."),
    ("The J-curve is tradable, not painful.", "Unlike PE’s accounting J-curve, "
     "ours is live and observable — an engineered entry point (and a co-invest "
     "“free look”) at de-risked governance at distressed prices."),
    ("Public scoreboard, daily.", "There is nowhere to hide poor capital "
     "allocation — the discipline behind a 92% deal-level hit rate over 20 years."),
], top, size=14.5, gap=11)

# ---- II.5c The non-linear path, honestly priced ----
_f3 = fig()
s, top = content("Strategic Case", "A non-linear path — honestly priced",
                 "Public engaged ownership marks to market every day. The "
                 "destination is the multiple; the route is rarely a straight "
                 "line.", ref=_f3)
# explanatory bullets (left)
checklist(s, [
    ("Value moves over the holding period.", "A 5× outcome is realised over "
     "years — the interim mark moves with the market, the cycle and the "
     "campaign. It is a journey, not a coupon."),
    ("PE looks smooth because it is appraised, not traded.", "Infrequent, "
     "model-based marks understate true economic volatility — “volatility "
     "laundering.” Marked daily, a levered buyout would swing as much or more."),
    ("Same economics, honest optics.", "We hold the same kind of assets PE "
     "does — we simply report them at live prices and accept the J-curve as the "
     "toll paid for liquidity."),
], top, left=Inches(0.7), width=Inches(6.2), size=13.5, gap=11)
# illustrative line chart (right)
cd = CategoryChartData()
cd.categories = ["0", "", "", "Yr 1", "", "", "Yr 2", "", "", "Yr 3", "", "", "Exit"]
cd.add_series("Public mark-to-market", (100, 92, 80, 72, 85, 95, 88, 112, 142,
                                        175, 212, 255, 300))
cd.add_series("PE appraisal marks", (100, 104, 110, 118, 128, 140, 155, 172,
                                     192, 215, 245, 272, 300))
gf = s.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(7.05), top - Inches(0.05),
                        Inches(5.9), Inches(3.9), cd)
ch = gf.chart
ch.has_title = True
ch.chart_title.text_frame.text = f"{_f3}.  Same destination, different reported path (indexed to 100)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(11)
ch.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb = SUBTLE
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(11)
ser = ch.plots[0].series
ser[0].format.line.color.rgb = NAVY
ser[0].format.line.width = Pt(2.5)
ser[1].format.line.color.rgb = SLATE_LT
ser[1].format.line.width = Pt(2.0)
ser[1].format.line.dash_style = 2  # dashed
for sline in ser:
    sline.smooth = True
dlabels(ser[0], fmt='0', size=8, color=NAVY_TX, pos=XL_LABEL_POSITION.BELOW)
ch.category_axis.tick_labels.font.size = Pt(9)
ch.value_axis.tick_labels.font.size = Pt(9)
ch.value_axis.has_major_gridlines = False
para(tbox(s, Inches(7.05), top + Inches(3.95), Inches(5.9), Inches(0.6)),
     "Illustrative. Both reach the same value; only the public line shows the "
     "honest interim drawdown.", 10, FOOT, first=True, italic=True, after=0)

# ---- II.5d Risk reframing summary table (memo Section 8.3) ----
s, top = content("Strategic Case", "Risk reframing for the investment committee",
                 "What the committee sees on the surface — versus what it is "
                 "actually underwriting.", ref=tbl())
rows = [
    ("What the IC sees", "What it is actually underwriting"),
    ("Reputational / headline risk",
     "The catalyst that forces repricing — and a moat that keeps competitors out"),
    ("Public J-curve & mark-to-market drawdown",
     "A live, tradable entry into an asset already de-risked by board control"),
    ("~50% peak-to-trough swings",
     "Optical volatility; two such drawdowns fully recovered (2010, 2020) · Sortino 2.40"),
    ("Concentration in 8–12 names",
     "Deliberate conviction sizing — the skill (85.7% hit on ≥10% names; +0.73)"),
    ("“Catching a falling knife”",
     "Hard-floor valuation gates; entry below the value of the rectifiable core"),
    ("Single-name dependence",
     "One ≥5× MOIC win every ~2 years; competitive even excluding the largest"),
]
tl = Inches(0.6); cw2 = [Inches(4.55), Inches(7.55)]; rh = Inches(0.56); y = top
for ri, (a, b) in enumerate(rows):
    head = ri == 0
    x = tl
    for ci, cell in enumerate((a, b)):
        fill = SLATE if head else (HEADERBG if ri % 2 else WHITE)
        rect(s, x, y, cw2[ci], rh, fill=fill)
        ctf = tbox(s, Emu(int(x) + int(Inches(0.14))), y,
                   Emu(int(cw2[ci]) - int(Inches(0.24))), rh,
                   anchor=MSO_ANCHOR.MIDDLE)
        para(ctf, cell, 12.5 if head else 12,
             WHITE if head else (SUBTLE if ci == 0 else NAVY_TX),
             bold=(head or ci == 1), first=True, after=0, lead=1.05)
        x = Emu(int(x) + int(cw2[ci]))
    y = Emu(int(y) + int(rh))
para(tbox(s, Inches(0.6), Emu(int(y) + int(Inches(0.15))), Inches(12.1), Inches(0.6)),
     "The IC is not paying for safety. It is paying for proven willingness to "
     "take compensated discomfort — +6 percentage points per year, net of every "
     "fee, for 20 years.", 13, SLATE, first=True, italic=True, after=0)

# ---- II.6 Track record (chart) ----
_f4 = fig()
s, top = content("Track Record", "Proof: outperformance with downside protection",
                 ref=_f4)
chart_data = CategoryChartData()
chart_data.categories = ["All markets", "Up markets", "Down markets"]
chart_data.add_series("Athanase", (1.5, 1.2, 0.2))
chart_data.add_series("MSCI IMI", (0.6, 1.3, -1.7))
cx, cy, cw, ch = Inches(0.7), top, Inches(7.1), Inches(4.2)
gframe = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, cx, cy, cw, ch, chart_data)
chart = gframe.chart
chart.has_title = True
chart.chart_title.text_frame.text = f"{_f4}.  Average monthly return (%)"
chart.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
chart.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb = SUBTLE
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
chart.legend.font.size = Pt(11)
plot = chart.plots[0]
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = NAVY
plot.series[1].format.fill.solid()
plot.series[1].format.fill.fore_color.rgb = SLATE_LT
plot.gap_width = 80
dlabels(plot, fmt='0.0"%"', size=9, color=BODY, pos=XL_LABEL_POSITION.OUTSIDE_END)
for ax in (chart.value_axis, chart.category_axis):
    ax.tick_labels.font.size = Pt(10)
    ax.tick_labels.font.color.rgb = BODY
# right-hand proof points
pr = tbox(s, Inches(8.1), top + Inches(0.1), Inches(4.6), Inches(4.5))
para(pr, "Haldex AB", 16, NAVY_TX, first=True, bold=True, after=2, font=SERIF)
para(pr, "81% IRR · 3.9x money multiple. Spun off Concentric, refocused the "
     "core, re-invested in 2020 for a further 2.0x in under two years.", 13,
     BODY, after=14, lead=1.15)
para(pr, "Compounding", 16, NAVY_TX, bold=True, after=2, font=SERIF)
para(pr, "~18x gross vs ~14x for MSCI IMI since 2006, with EBITA growth of "
     "+14.5% across portfolio companies.", 13, BODY, after=14, lead=1.15)
para(pr, "Hit rate", 16, NAVY_TX, bold=True, after=2, font=SERIF)
para(pr, "39 larger investments since 2006 — profitable in 36, lost money in "
     "only 3.", 13, BODY, after=0, lead=1.15)

# ---- II.6b Transactions: AIP Fund II ----
FUND2, HIST = load_transactions()
s, top = content("Track Record", "AIP Fund II — transactions (2015–present)",
                 "Each position shown against its benchmark; invested capital "
                 "weighted MOIC of 2.2x.", ref=tbl())
cw = [Inches(3.3), Inches(2.0), Inches(1.7), Inches(1.7), Inches(2.0)]
endy = deal_table(s, Inches(0.75), top, FUND2, cw, font=10, rh=Inches(0.212))
# summary strip
sy = Emu(int(endy) + int(Inches(0.10)))
rect(s, Inches(0.75), sy, Inches(10.7), Inches(0.34), fill=NAVY)
stf = tbox(s, Inches(0.9), sy, Inches(10.4), Inches(0.34), anchor=MSO_ANCHOR.MIDDLE)
para(stf, "17 deals   ·   weighted MOIC 2.2x   ·   ex best/worst IRR 41% vs "
     "index ~9%   ·   36 of 39 deals profitable since 2006", 10.5, WHITE,
     first=True, after=0)
ny = Emu(int(sy) + int(Inches(0.34)) + int(Inches(0.08)))
nt = tbox(s, Inches(0.75), ny, Inches(11.8), Inches(0.35))
para(nt, "IRR/MOIC net of the position; “n.m.” where short holding periods make "
     "annualised figures not meaningful. Capital in SEK.", 8.5, FOOT, first=True,
     after=0)

# ---- II.6c Transactions: prior period 2006-2014 ----
s, top = content("Track Record", "Prior period — transactions (2006–2014)",
                 "The same strategy, the prior fund — repeatable across two "
                 "decades and multiple cycles.", ref=tbl())
half = 11
cwh = [Inches(2.7), Inches(1.45), Inches(1.45)]
deal_table(s, Inches(0.6), top, HIST[:half], cwh, font=10,
           rh=Inches(0.285), cols=("company", "irr", "moic"),
           headers=("Company", "IRR", "MOIC"))
deal_table(s, Inches(6.95), top, HIST[half:], cwh, font=10,
           rh=Inches(0.285), cols=("company", "irr", "moic"),
           headers=("Company", "IRR", "MOIC"))
sy = top + Inches(3.55)
rect(s, Inches(0.6), sy, Inches(11.05), Inches(0.34), fill=NAVY)
stf = tbox(s, Inches(0.75), sy, Inches(10.7), Inches(0.34), anchor=MSO_ANCHOR.MIDDLE)
para(stf, "22 deals   ·   weighted MOIC 2.7x   ·   average IRR 68%   ·   "
     "index outperformance +25 pts (invested-weighted)", 10.5, WHITE, first=True,
     after=0)
nt = tbox(s, Inches(0.6), Inches(6.95), Inches(11.8), Inches(0.4))
para(nt, "Selected larger deals of the investment team. Losses shown in italic. "
     "Capital in SEK; ~$400M invested over the period.", 8.5, FOOT, first=True,
     after=0)

# ---- II.6d Independent validation: 20-year NET record ----
s, top = content("Independent Review",
                 "Independently validated: 20 years, net of every fee",
                 "A prospective LP’s investment committee reconciled the "
                 "security-level (StatPro) data and the 236-month LP-level NET "
                 "return series.")
stats = [("+1,402%", "cumulative NET return\nover 19.7 years"),
         ("+14.8%", "annualised NET\n(time-weighted, all fees)"),
         ("+6.0 pts", "annualised alpha,\nnet, for 20 years"),
         ("3.11×", "the relevant equity\nbenchmark")]
bw = Inches(2.95); gapx = Inches(0.18); x = Inches(0.75); y = top
for big, lab in stats:
    rect(s, x, y, bw, Inches(1.8), fill=HEADERBG)
    ctf = tbox(s, x, y + Inches(0.26), bw, Inches(1.4))
    para(ctf, big, 34, NAVY_TX, first=True, align=PP_ALIGN.CENTER, after=4,
         font=SERIF)
    for line in lab.split("\n"):
        para(ctf, line, 11.5, SUBTLE, align=PP_ALIGN.CENTER, after=0)
    x = Emu(int(x) + int(bw) + int(gapx))
checklist(s, [
    ("Two full cycles, survived.", "Drawdowns of −52.8% (2008, recovered by "
     "2010) and −45.9% (2017–2020, recovered by Sep 2020) — navigated live with "
     "the same philosophy, not back-tested."),
    ("Top-tier downside-adjusted return.", "20-year Sortino 1.54; AIPFII Sortino "
     "2.40, with upside volatility 3.9× the downside."),
    ("Realised cash, not marks.", "The largest contributors are realised LP cash "
     "— removing the mark-to-model objection."),
], top + Inches(2.05), size=14, gap=10)
nt = tbox(s, Inches(0.75), Inches(6.95), Inches(11.8), Inches(0.4))
para(nt, "Source: independent institutional due-diligence review, May 2026, on "
     "StatPro security-level data and LP-level NET monthly returns (2006–2026).",
     8.5, FOOT, first=True, after=0)

# ---- II.6e Repeatability & conviction sizing ----
s, top = content("Independent Review",
                 "Repeatable and conviction-sized — not lucky",
                 "The review tested the “you can’t rely on the big winners” "
                 "objection directly against 39 deals over 20 years.")
colL = tbox(s, Inches(0.75), top, Inches(5.85), Inches(4.6))
para(colL, "REPEATABLE", 13, SLATE, first=True, bold=True, after=7)
for t in ["92% deal-level hit rate (36 of 39 profitable); losses ~8% of "
          "cumulative invested capital",
          "10 of 39 deals returned ≥5× MOIC — one such win every ~2 years",
          "Two decade-definers, not one: Klarna 28× (2007) and Athanase Tech "
          "5.4× (2021)",
          "Robust: even excluding Athanase Tech entirely, AIPFII still ~10% "
          "net annualised"]:
    para(colL, t, 14, BODY, after=9, lead=1.13)
colR = tbox(s, Inches(7.0), top, Inches(5.85), Inches(4.6))
para(colR, "CONVICTION-SIZED (THE SKILL SIGNAL)", 13, SLATE, first=True,
     bold=True, after=7)
for t in ["The 7 positions ever sized ≥10% of NAV hit at 85.7% and contributed "
          "+209 pts — ~77% of the positive book",
          "+0.73 rank correlation between position size and outcome within "
          "winners — they size their best ideas biggest",
          "Win/loss asymmetry: average winner 2.74× the average loser; "
          "+333 pts of winners vs −61 pts of losers"]:
    para(colR, t, 14, BODY, after=9, lead=1.13)
para(tbox(s, Inches(0.75), Inches(6.35), Inches(11.8), Inches(0.55)),
     "“The team has done this twice in 20 years, not once — the ‘every 10–20 "
     "years’ objection is empirically unsupportable.”", 13, SLATE, first=True,
     italic=True, after=0)

# ---- II.6f Governing the allocation: why the path won't get you fired ----
s, top = content("Allocation", "Why a non-linear path won’t get you fired",
                 "Career risk is managed by mandate design — not by hoping the "
                 "line goes straight.")
colL = tbox(s, Inches(0.75), top, Inches(5.85), Inches(4.6))
para(colL, "SIZE AND STRUCTURE IT RIGHT", 13, SLATE, first=True, bold=True, after=7)
for t in ["Size at 1–3% of total assets in a special-situations sleeve — no "
          "single mark can dominate the portfolio or force action",
          "Pre-agree a ~50% mark-to-market drawdown tolerance with the IC at "
          "commitment; align reporting to the multi-year holding cycle, not "
          "monthly marks",
          "Define reallocation triggers in advance — manager turnover, style "
          "drift, weaker sizing discipline — never the share price"]:
    para(colL, t, 14, BODY, after=10, lead=1.13)
colR = tbox(s, Inches(7.0), top, Inches(5.85), Inches(4.6))
para(colR, "JUDGE THE RIGHT THING", 13, SLATE, first=True, bold=True, after=7)
for t in ["Underwrite whether business risk is falling — board seat secured, "
          "cash-flow control — not whether the news flow looks calm",
          "The big winners are realised LP cash, not marks: the value is "
          "bankable, not modelled",
          "Two ~50% drawdowns have already been recovered with the same "
          "philosophy — the path has been bumpy before and still paid +6 pts/yr "
          "net"]:
    para(colR, t, 14, BODY, after=10, lead=1.13)
sy = Inches(6.25)
rect(s, Inches(0.75), sy, Inches(11.85), Inches(0.62), fill=NAVY)
para(tbox(s, Inches(0.95), sy, Inches(11.45), Inches(0.62), anchor=MSO_ANCHOR.MIDDLE),
     "A pre-underwritten, correctly-sized, fundamentally-monitored allocation "
     "turns a bad quarter into “as expected” — not a surprise that invites "
     "blame.", 13.5, WHITE, first=True, italic=True, after=0)

# ---- II.7 ESG / responsible ownership ----
s, top = content("Ownership Model", "Responsible ownership, by construction")
checklist(s, [
    ("Governance is the lever.", "Value is created through board representation "
     "and active promotion of strong corporate governance across all holdings."),
    ("Five improvement dimensions.", "Operational, financial, structural, "
     "environmental and social improvements — each tied to enterprise value."),
    ("PRI signatory since 2019,", "adhering to the Six Principles for Responsible "
     "Investment."),
    ("Credible oversight.", "SEB custodian · MUFG administrator · KPMG auditor · "
     "Swedish FI-registered AIFM."),
], top, size=16, gap=14)

# ===========================================================================
# CLOSING
# ===========================================================================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, fill=NAVY)
wordmark(s, Inches(0.6), Inches(0.55), WHITE, scale=1.0)
rect(s, 0, Inches(2.7), Inches(0.5), Pt(2.2), fill=SLATE_LT)
tf = tbox(s, Inches(0.6), Inches(2.95), Inches(11.0), Inches(1.6))
para(tf, "The asset class rewards patient, governance-minded capital.", 26,
     WHITE, first=True, after=8, font=SERIF)
para(tf, "Among engaged owners, Athanase brings the team, the discipline and the "
     "20-year proof to deliver it.", 18, DIVIDER, italic=True, after=0, font=SERIF)
# contact
ct = tbox(s, Inches(0.6), Inches(5.4), Inches(6), Inches(1.4))
para(ct, "Stefan Charette", 15, WHITE, first=True, bold=True, after=2)
para(ct, "Athanase Industrial Partner", 12, DIVIDER, after=8)
para(ct, "charette@athanase.se   ·   +46 73 994 7079", 12, SLATE_LT, after=0)
ot = tbox(s, Inches(9.0), Inches(5.4), Inches(3.8), Inches(1.5))
para(ot, "Offices", 12, WHITE, first=True, bold=True, after=4)
para(ot, "Birger Jarlsgatan 6, 114 34 Stockholm, Sweden", 11, SLATE_LT, after=4)
para(ot, "Landmark Square, West Bay Road, Grand Cayman", 11, SLATE_LT, after=0)

# ===========================================================================
# Contents (TOC) page
# ===========================================================================
def _toc_col(col_x, col_w, header, entries):
    htf = tbox(agenda_s, col_x, agenda_top, col_w, Inches(0.4))
    para(htf, header, 14, SLATE, first=True, bold=True, after=0, font=SERIF)
    ytf = tbox(agenda_s, col_x, Emu(int(agenda_top) + int(Inches(0.5))),
               col_w, Inches(4.7))
    for i, (num, title, ref) in enumerate(entries):
        p = ytf.paragraphs[0] if i == 0 else ytf.add_paragraph()
        p.space_after = Pt(7); p.line_spacing = 1.04
        r0 = p.add_run(); r0.text = num + "   "
        r0.font.size = Pt(11.5); r0.font.bold = True
        r0.font.color.rgb = SLATE; r0.font.name = SANS
        r1 = p.add_run(); r1.text = title
        r1.font.size = Pt(11.5); r1.font.color.rgb = NAVY_TX; r1.font.name = SANS
        if ref:
            r2 = p.add_run(); r2.text = "   (" + ref + ")"
            r2.font.size = Pt(9.5); r2.font.italic = True
            r2.font.color.rgb = SUBTLE; r2.font.name = SANS

_p1 = [e for e in TOC if e[0].startswith("1.")]
_p2 = [e for e in TOC if e[0].startswith("2.")]
_cw = Inches(5.85); _xL = Inches(0.7)
_xR = Emu(int(_xL) + int(_cw) + int(Inches(0.5)))
_toc_col(_xL, _cw, "Part I  ·  Why allocate to engaged ownership", _p1)
_toc_col(_xR, _cw, "Part II  ·  Why choose Athanase", _p2)

# ===========================================================================
# Convert the whole deck from 16:9 to the brand's 4:3 page format
# (10.667" x 8.0" = the official AIP-deck.pptx page size).  Layout was authored
# in 16:9 coordinates, so we rescale every shape's geometry and font.
#   - geometry: x by sx, y by sy (full-bleed rects keep filling the page)
#   - fonts:    by sx (horizontal is the binding constraint -> wrapping preserved)
#   - pictures: uniform sx (preserve logo aspect ratio)
# ===========================================================================
from pptx.oxml.ns import qn as _qn
from pptx.enum.shapes import MSO_SHAPE_TYPE as _MST

TARGET_W = 9753600   # EMU, 10.667"
TARGET_H = 7315200   # EMU, 8.0"


def _rescale_chart_fonts(gframe, f):
    el = gframe.chart.part._element
    for rpr in el.iter():
        tag = rpr.tag.split("}")[-1]
        if tag in ("rPr", "defRPr") and rpr.get("sz") is not None:
            rpr.set("sz", str(max(100, int(round(int(rpr.get("sz")) * f)))))


def _rescale_shape(sh, sx, sy):
    # geometry
    try:
        L, T = sh.left, sh.top
        W, H = sh.width, sh.height
    except Exception:
        L = T = W = H = None
    is_pic = sh.shape_type == _MST.PICTURE
    if None not in (L, T, W, H):
        sh.left = int(L * sx)
        sh.top = int(T * sy)
        if is_pic:                      # uniform scale -> keep aspect ratio
            sh.width = int(W * sx)
            sh.height = int(H * sx)
        else:
            sh.width = int(W * sx)
            sh.height = int(H * sy)
    # fonts (scale by sx so text/box ratio is preserved horizontally)
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size is not None:
                    r.font.size = Pt(round(r.font.size.pt * sx, 1))
    if sh.has_chart:
        _rescale_chart_fonts(sh, sx)


_sx = TARGET_W / int(prs.slide_width)
_sy = TARGET_H / int(prs.slide_height)
for _slide in prs.slides:
    for _sh in _slide.shapes:
        _rescale_shape(_sh, _sx, _sy)
prs.slide_width = TARGET_W
prs.slide_height = TARGET_H

# ---------------------------------------------------------------------------
# Brand the theme itself: replace the default Office colour scheme with the AIP
# Blue group and set theme fonts to Times New Roman / Arial, so no off-brand
# swatch survives anywhere in the file (guideline: stay within the palette).
# ---------------------------------------------------------------------------
def _brand_theme(prs):
    from lxml import etree
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    THEME_REL = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                 "relationships/theme")
    scheme = {
        "dk1": "152130", "lt1": "FFFFFF", "dk2": "0E1620", "lt2": "F6F7F9",
        "accent1": "152130", "accent2": "314359", "accent3": "556A83",
        "accent4": "0E1620", "accent5": "E2E5E9", "accent6": "F6F7F9",
        "hlink": "314359", "folHlink": "556A83",
    }
    seen = set()
    for master in prs.slide_masters:
        theme = master.part.part_related_by(THEME_REL)
        if theme.partname in seen:
            continue
        seen.add(theme.partname)
        root = etree.fromstring(theme.blob)
        clr = root.find(f"{a}themeElements/{a}clrScheme")
        for name, hexv in scheme.items():
            el = clr.find(f"{a}{name}")
            if el is None:
                continue
            srgb = el.find(f"{a}srgbClr")
            sysc = el.find(f"{a}sysClr")
            if srgb is not None:
                srgb.set("val", hexv)
            elif sysc is not None:           # convert sysClr -> srgbClr
                el.remove(sysc)
                etree.SubElement(el, f"{a}srgbClr", {"val": hexv})
        fonts = root.find(f"{a}themeElements/{a}fontScheme")
        for grp, face in (("majorFont", "Times New Roman"), ("minorFont", "Arial")):
            latin = fonts.find(f"{a}{grp}/{a}latin")
            if latin is not None:
                latin.set("typeface", face)
        theme._blob = etree.tostring(root, xml_declaration=True,
                                     encoding="UTF-8", standalone=True)
    return prs


_brand_theme(prs)

out = "Athanase_Engaged_Ownership_Allocator_Deck.pptx"
prs.save(out)
print("Saved", out, "with", len(prs.slides._sldIdLst),
      "slides at 4:3 (%.2f x %.2f in)" % (TARGET_W / 914400, TARGET_H / 914400))
