"""Generate a PowerPoint deck for the activist-investing institutional memo."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Palette
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
ACCENT = RGBColor(0xC8, 0x9B, 0x3C)  # muted gold
DARK = RGBColor(0x22, 0x2A, 0x33)
GREY = RGBColor(0x5A, 0x63, 0x6E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xF4, 0xF1, 0xEA)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def add_box(slide, l, t, w, h, fill=None, line=None):
    shp = slide.shapes.add_shape(1, l, t, w, h)  # rectangle
    shp.fill.solid() if fill else shp.fill.background()
    if fill:
        shp.fill.fore_color.rgb = fill
    if line:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def textbox(slide, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    return tf


def para(tf, text, size, color, bold=False, first=False, align=PP_ALIGN.LEFT,
         space_after=8, italic=False, bullet=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    p.line_spacing = 1.08
    run = p.add_run()
    run.text = ("•  " + text) if bullet else text
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    f.name = "Calibri"
    return p


# ---------- Slide 1: Title ----------
s = prs.slides.add_slide(BLANK)
add_box(s, 0, 0, SW, SH, fill=NAVY)
add_box(s, 0, Inches(4.55), SW, Inches(0.06), fill=ACCENT)
tf = textbox(s, Inches(0.9), Inches(2.2), Inches(11.5), Inches(2.2))
para(tf, "The Case for Activist Investing", 44, WHITE, bold=True, first=True,
     space_after=10)
para(tf, "Why Endowments & Pension Funds Should Allocate to the Activist Asset Class",
     22, ACCENT, space_after=0)
tf2 = textbox(s, Inches(0.9), Inches(4.75), Inches(11.5), Inches(1.2))
para(tf2, "An Institutional Investment Memorandum", 16, LIGHT, first=True,
     italic=True, space_after=4)
para(tf2, "Prepared May 2026", 13, GREY, space_after=0)


def content_slide(kicker, title):
    s = prs.slides.add_slide(BLANK)
    add_box(s, 0, 0, SW, SH, fill=WHITE)
    add_box(s, 0, 0, Inches(0.28), SH, fill=NAVY)
    # header
    htf = textbox(s, Inches(0.7), Inches(0.45), Inches(12), Inches(1.3))
    para(htf, kicker.upper(), 13, ACCENT, bold=True, first=True, space_after=2)
    para(htf, title, 30, NAVY, bold=True, space_after=0)
    add_box(s, Inches(0.72), Inches(1.72), Inches(2.0), Inches(0.05), fill=ACCENT)
    return s


def bullets(s, items, top=Inches(2.05), left=Inches(0.75), width=Inches(11.8),
            size=17, lead_size=None):
    tf = textbox(s, left, top, width, Inches(4.9))
    for i, (lead, body) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(13)
        p.line_spacing = 1.12
        r1 = p.add_run()
        r1.text = "▪  " + lead
        r1.font.size = Pt(lead_size or size)
        r1.font.bold = True
        r1.font.color.rgb = NAVY
        r1.font.name = "Calibri"
        if body:
            r2 = p.add_run()
            r2.text = " — " + body
            r2.font.size = Pt(size)
            r2.font.color.rgb = DARK
            r2.font.name = "Calibri"
    return tf


# ---------- Slide 2: Executive Summary ----------
s = content_slide("Executive Summary", "A Durable, Institutional-Grade Asset Class")
tf = textbox(s, Inches(0.75), Inches(2.0), Inches(11.9), Inches(1.4))
para(tf, "Activism — taking meaningful equity stakes and working to change how "
     "companies are governed, capitalized, and run — offers asset owners a rare "
     "combination of equity-like returns, differentiated alpha, and a natural "
     "fit with fiduciary stewardship.", 16, DARK, first=True, space_after=0)
bullets(s, [
    ("Return & alpha", "an idiosyncratic, catalyst-driven source of excess return"),
    ("Horizon & duty", "matched to long liabilities; stewardship made consequential"),
    ("Diversification", "event-driven returns that lift total-portfolio Sharpe"),
    ("Governance synergy", "applied ESG that improves the broader portfolio's beta"),
    ("Disciplined sizing", "a measured 3%–8% allocation within the equity sleeve"),
], top=Inches(3.45), size=16)

# ---------- Slide 3: Page 1 Returns ----------
s = content_slide("1 · The Proposition", "The Return & Alpha Engine")
bullets(s, [
    ("Catalyst, not absorption", "activists change the outcome at undervalued, "
     "correctable companies rather than passively absorbing underperformance"),
    ("~7% announcement return", "with no reversal over the following year; success "
     "or partial success in ~two-thirds of campaigns (Brav et al., 2008)"),
    ("~10.2% around the 13D", "plus a further ~11.4% over the next year for "
     "hedge-fund activists (Klein & Zur, 2009)"),
    ("Idiosyncratic alpha", "returns tied to a specific change at a specific firm — "
     "scarce and valuable in a beta-dominated portfolio"),
    ("Skill-dependent, hard to arbitrage", "concentrated capital and public "
     "positioning keep the strategy from being commoditized"),
], size=16)

# ---------- Slide 4: Page 2 Horizon / Fiduciary ----------
s = content_slide("2 · Fit", "Aligned With Horizon & Fiduciary Duty")
bullets(s, [
    ("Patient capital wins", "board seats, strategic reviews and turnarounds play "
     "out over years — perpetual endowments and long-dated pensions can hold through"),
    ("Stewardship is the duty", "engaged ownership expresses fiduciary obligation "
     "(UK Stewardship Code; UN PRI) rather than conflicting with it"),
    ("Mission alignment", "demanding accountability and good governance reflects "
     "institutional values"),
    ("Engagement as risk management", "turns hidden governance risk in the book "
     "into a remediable, return-generating exposure"),
], size=17)

# ---------- Slide 5: Page 3 Diversification ----------
s = content_slide("3 · Portfolio", "Diversification & the Total-Portfolio View")
bullets(s, [
    ("A distinct return engine", "alpha driven by proxy contests, spin-offs and "
     "capital returns — largely independent of rates, growth and inflation"),
    ("Diversifies the equity book", "small/mid-cap value and special situations, "
     "underweight in most institutional portfolios"),
    ("Counter-cyclical dry powder", "dislocations create the next targets; locked-up "
     "capital deploys into weakness"),
    ("PE-style upside, public-market terms", "operational engagement with daily "
     "pricing, transparency, lower fees and better liquidity"),
], size=17)

# ---------- Slide 6: Page 4 Governance / ESG ----------
s = content_slide("4 · Stewardship", "Governance & the ESG Synergy")
bullets(s, [
    ("Activism is applied governance", "board independence, pay-for-performance and "
     "capital discipline are the 'G' of ESG put into practice"),
    ("Constructive & ESG-oriented activism", "express return and stewardship goals "
     "through one allocation, disciplined by a value thesis"),
    ("Positive externality", "governance precedents improve the quality of the beta "
     "the whole portfolio owns"),
    ("Defensible to stakeholders", "\"we own companies and make them better-run and "
     "more valuable\" is an easy story to tell beneficiaries"),
], size=17)

# ---------- Slide 7: Page 5 Risks & Implementation ----------
s = content_slide("5 · Execution", "Risks, Selection & Implementation")
# two columns
left = textbox(s, Inches(0.75), Inches(2.0), Inches(5.9), Inches(4.8))
para(left, "KEY RISKS", 15, ACCENT, bold=True, first=True, space_after=8)
for t in ["Concentration & single-name volatility",
          "Headline and litigation exposure",
          "Liquidity during campaigns / lock-ups",
          "Wide manager dispersion",
          "Crowding and capacity limits"]:
    para(left, t, 15, DARK, bullet=True, space_after=7)
right = textbox(s, Inches(7.0), Inches(2.0), Inches(5.8), Inches(4.8))
para(right, "IMPLEMENTATION", 15, ACCENT, bold=True, first=True, space_after=8)
for t in ["Size prudently: 3%–8% of the equity sleeve",
          "Diversify across managers & campaign styles",
          "Match liquidity terms to the liability profile",
          "Govern explicitly at board level",
          "Benchmark over multi-year, event-driven horizons"]:
    para(right, t, 15, DARK, bullet=True, space_after=7)

# ---------- Slide 8: Conclusion ----------
s = prs.slides.add_slide(BLANK)
add_box(s, 0, 0, SW, SH, fill=NAVY)
add_box(s, Inches(0.9), Inches(2.0), Inches(2.0), Inches(0.05), fill=ACCENT)
tf = textbox(s, Inches(0.9), Inches(2.25), Inches(11.5), Inches(3.5))
para(tf, "Conclusion", 34, WHITE, bold=True, first=True, space_after=14)
para(tf, "Activism is uniquely matched to institutional strengths: a long horizon "
     "that rewards patience, a fiduciary duty expressed through engaged ownership, "
     "and an idiosyncratic, diversifying return stream.", 19, LIGHT, space_after=10)
para(tf, "The risks — concentration, headline exposure and manager dispersion — are "
     "real but manageable within a deliberate framework. A measured allocation is a "
     "prudent, mission-aligned investment in better-run companies and a stronger "
     "total portfolio.", 19, LIGHT, space_after=0)

# ---------- Slide 9: References ----------
s = content_slide("Appendix", "References")
tf = textbox(s, Inches(0.75), Inches(2.05), Inches(11.9), Inches(4.8))
refs = [
    "Brav, A., Jiang, W., Partnoy, F., & Thomas, R. (2008). Hedge Fund Activism, "
    "Corporate Governance, and Firm Performance. The Journal of Finance, 63(4), 1729–1775.",
    "Klein, A., & Zur, E. (2009). Entrepreneurial Shareholder Activism: Hedge Funds "
    "and Other Private Investors. The Journal of Finance, 64(1), 187–229.",
    "Bebchuk, L. A., Brav, A., & Jiang, W. (2015). The Long-Term Effects of Hedge "
    "Fund Activism. Columbia Law Review, 115(5), 1085–1156.",
    "Brav, A., Jiang, W., & Kim, H. (2015). The Real Effects of Hedge Fund Activism: "
    "Productivity, Asset Allocation, and Labor Outcomes. The Review of Financial "
    "Studies, 28(10), 2723–2769.",
    "Financial Reporting Council (2020). The UK Stewardship Code 2020; "
    "Principles for Responsible Investment (PRI).",
]
for i, r in enumerate(refs):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.space_after = Pt(12)
    p.line_spacing = 1.1
    r1 = p.add_run()
    r1.text = f"{i+1}.  "
    r1.font.size = Pt(14)
    r1.font.bold = True
    r1.font.color.rgb = ACCENT
    r1.font.name = "Calibri"
    r2 = p.add_run()
    r2.text = r
    r2.font.size = Pt(14)
    r2.font.color.rgb = DARK
    r2.font.name = "Calibri"

out = "Activist_Investing_Case_for_Institutions.pptx"
prs.save(out)
print("Saved", out, "with", len(prs.slides._sldIdLst), "slides")
