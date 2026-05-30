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
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

# ---- Brand palette ---------------------------------------------------------
NAVY      = RGBColor(0x14, 0x25, 0x3B)   # cover / divider background
NAVY_TX   = RGBColor(0x1E, 0x33, 0x4E)   # serif title on white
SLATE     = RGBColor(0x4A, 0x5D, 0x75)   # table headers / accent shapes
SLATE_LT  = RGBColor(0x8A, 0x99, 0xAB)   # secondary accent
HEADERBG  = RGBColor(0xF1, 0xF2, 0xF4)   # light grey header band
BODY      = RGBColor(0x33, 0x38, 0x40)
SUBTLE    = RGBColor(0x5A, 0x63, 0x6E)
DIVIDER   = RGBColor(0xC9, 0xD3, 0xDD)   # light blue-grey divider text
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
FOOT      = RGBColor(0x9A, 0xA4, 0xB0)
GOLD      = RGBColor(0x6E, 0x82, 0x9B)   # subtle rule accent

SERIF = "Georgia"
SANS  = "Arial"

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
         align=PP_ALIGN.LEFT, after=8, lead=1.1, font=SANS):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(after)
    p.line_spacing = lead
    r = p.add_run(); r.text = text
    f = r.font
    f.size = Pt(size); f.bold = bold; f.italic = italic
    f.color.rgb = color; f.name = font
    return p


def logo_mark(slide, left, top, color, h=Inches(0.30)):
    """Stylised Athanase bar-chart mark."""
    unit = int(h)
    heights = [0.34, 0.62, 1.0, 0.55, 0.28]
    bw = Emu(int(unit * 0.16))
    gap = Emu(int(unit * 0.12))
    x = left
    base = top + h
    for i, hf in enumerate(heights):
        bh = Emu(int(unit * hf))
        b = rect(slide, x, base - bh, bw, bh, fill=color)
        # round the bars a touch
        x = Emu(int(x) + int(bw) + int(gap))
    return x  # right edge


def wordmark(slide, left, top, color, scale=1.0):
    right = logo_mark(slide, left, top, color, h=Inches(0.30 * scale))
    tf = tbox(slide, Emu(int(right) + Inches(0.10)), top - Inches(0.04),
              Inches(3.2 * scale), Inches(0.5 * scale))
    p = tf.paragraphs[0]; p.space_after = Pt(0); p.line_spacing = 0.95
    r = p.add_run(); r.text = "Athanase"
    r.font.size = Pt(18 * scale); r.font.bold = True
    r.font.color.rgb = color; r.font.name = SANS
    p2 = tf.add_paragraph(); p2.space_after = Pt(0); p2.line_spacing = 1.0
    r2 = p2.add_run(); r2.text = "I N D U S T R I A L   P A R T N E R"
    r2.font.size = Pt(6.5 * scale); r2.font.color.rgb = color; r2.font.name = SANS


def footer(slide):
    _state["n"] += 1
    tf = tbox(slide, Inches(9.0), Inches(7.05), Inches(4.0), Inches(0.3))
    para(tf, f"Strictly confidential          {_state['n']}", 8.5, FOOT,
         first=True, align=PP_ALIGN.RIGHT, after=0)


def divider(title, kicker=None):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=NAVY)
    wordmark(s, Inches(0.6), Inches(0.55), WHITE, scale=0.9)
    if kicker:
        ktf = tbox(s, Inches(1.0), Inches(3.05), Inches(11.3), Inches(0.5),
                   anchor=MSO_ANCHOR.MIDDLE)
        para(ktf, kicker.upper(), 13, SLATE_LT, first=True,
             align=PP_ALIGN.CENTER, after=0)
    tf = tbox(s, Inches(1.0), Inches(3.4), Inches(11.3), Inches(1.2),
              anchor=MSO_ANCHOR.MIDDLE)
    para(tf, title, 40, DIVIDER, italic=True, first=True,
         align=PP_ALIGN.CENTER, after=0, font=SERIF)
    return s


def content(section, title, subtitle=None):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=WHITE)
    # top section label + mark
    logo_mark(s, Inches(0.55), Inches(0.26), SLATE, h=Inches(0.22))
    lt = tbox(s, Inches(1.15), Inches(0.24), Inches(7.0), Inches(0.3))
    para(lt, section, 11, SLATE, first=True, after=0)
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
s, top = content("Overview", "What this presentation covers")
tf = tbox(s, Inches(0.75), top, Inches(11.8), Inches(4.6))
rows = [
    ("Part I", "Why allocate to engaged ownership", "The asset-class case: a "
     "catalyst-driven, diversifying source of equity return that fits a long-"
     "horizon, fiduciary mandate."),
    ("Part II", "Why choose Athanase", "Among engaged owners, the team, "
     "opportunity set, margin-for-error and 20-year track record that "
     "distinguish Athanase Industrial Partner."),
]
for i, (tag, head, body) in enumerate(rows):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.space_after = Pt(10); p.line_spacing = 1.12
    r0 = p.add_run(); r0.text = tag + "   "
    r0.font.size = Pt(20); r0.font.bold = True; r0.font.color.rgb = SLATE
    r0.font.name = SERIF
    r1 = p.add_run(); r1.text = head
    r1.font.size = Pt(20); r1.font.bold = True; r1.font.color.rgb = NAVY_TX
    r1.font.name = SERIF
    p2 = tf.add_paragraph(); p2.space_after = Pt(20); p2.line_spacing = 1.12
    r2 = p2.add_run(); r2.text = body
    r2.font.size = Pt(15); r2.font.color.rgb = BODY; r2.font.name = SANS

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
# heritage line
ht = tbox(s, Inches(0.75), top, Inches(11.8), Inches(0.6))
para(ht, "Heritage:  Öresund (1992)  →  Custos (1996)  →  Creades  "
     "→  Athanase Industrial Partner (2015)", 14, SLATE, first=True,
     bold=True, after=0, font=SERIF)
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
], top + Inches(0.7), size=15, gap=11)

# ---- II.3 Market opportunity ----
s, top = content("Market Opportunity", "The inefficiency Athanase harvests",
                 "Corporate decay is structural — it continuously recreates "
                 "low-risk, high-return entry points.")
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
], top, size=15.5, gap=13)

# ---- II.4 Margin for error vs PE (table) ----
s, top = content("Investment Strategy",
                 "We don’t need to be right more than ~50%",
                 "Because we buy at a discount, not at a 40% control premium.")
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

# ---- II.6 Track record (chart) ----
s, top = content("Track Record", "Proof: outperformance with downside protection")
chart_data = CategoryChartData()
chart_data.categories = ["All markets", "Up markets", "Down markets"]
chart_data.add_series("Athanase", (1.5, 1.2, 0.2))
chart_data.add_series("MSCI IMI", (0.6, 1.3, -1.7))
cx, cy, cw, ch = Inches(0.7), top, Inches(7.1), Inches(4.2)
gframe = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, cx, cy, cw, ch, chart_data)
chart = gframe.chart
chart.has_title = True
chart.chart_title.text_frame.text = "Average monthly return (%)"
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
para(pr, "Down-market years", 16, NAVY_TX, bold=True, after=2, font=SERIF)
para(pr, "Outperformed in 2008 (+2.2%), 2011 (+7.9%), 2015 (+11.8%) and 2022 "
     "(+20.1%).", 13, BODY, after=0, lead=1.15)

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

out = "Athanase_Engaged_Ownership_Allocator_Deck.pptx"
prs.save(out)
print("Saved", out, "with", len(prs.slides._sldIdLst), "slides")
