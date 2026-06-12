"""Builds TBK_rd_gap_slides.pptx (10 slides)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

DARK = RGBColor(0x1F, 0x4E, 0x78)
RED = RGBColor(0x7F, 0x1D, 0x1D)
GREEN = RGBColor(0x54, 0x82, 0x35)
GREY = RGBColor(0x59, 0x59, 0x59)
LGREY = RGBColor(0xF2, 0xF2, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def slide():
    return prs.slides.add_slide(BLANK)


def txt(s, l, t, w, h, text, size=14, bold=False, color=RGBColor(0, 0, 0), align=PP_ALIGN.LEFT):
    box = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = align
        for r in p.runs:
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
    return box


def header(s, title, sub=None):
    bar = s.shapes.add_textbox(Inches(0), Inches(0), prs.slide_width, Inches(0.95))
    bar.fill.solid()
    bar.fill.fore_color.rgb = DARK
    tf = bar.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.5)
    tf.margin_top = Inches(0.18)
    p = tf.paragraphs[0]
    p.text = title
    for r in p.runs:
        r.font.size = Pt(26)
        r.font.bold = True
        r.font.color.rgb = WHITE
    if sub:
        txt(s, 0.5, 1.0, 12.3, 0.4, sub, size=12, color=GREY)


def panel(s, l, t, w, h, title, body, title_color=DARK, body_size=12):
    box = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = LGREY
    box.line.color.rgb = title_color
    box.line.width = Pt(1.25)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)
    tf.margin_top = Inches(0.12)
    p = tf.paragraphs[0]
    p.text = title
    for r in p.runs:
        r.font.size = Pt(15)
        r.font.bold = True
        r.font.color.rgb = title_color
    for line in body.split("\n"):
        p = tf.add_paragraph()
        p.text = line
        p.space_before = Pt(4)
        for r in p.runs:
            r.font.size = Pt(body_size)
            r.font.color.rgb = RGBColor(0x20, 0x20, 0x20)


# ---------- 1 Title ----------
s = slide()
bg = s.shapes.add_textbox(Inches(0), Inches(0), prs.slide_width, prs.slide_height)
bg.fill.solid()
bg.fill.fore_color.rgb = DARK
txt(s, 1.0, 2.3, 11.3, 1.2, "Closing the R&D Gap", size=44, bold=True, color=WHITE)
txt(s, 1.0, 3.4, 11.3, 0.9, "TBK Co., Ltd. (7277.T) — brakes and pumps: market evolution, competitor innovation, and the reinvestment requirement", size=20, color=RGBColor(0xD0, 0xDC, 0xE8))
txt(s, 1.0, 6.5, 11.3, 0.5, "Based on peer-reviewed and adversarially verified research, competitor filings (incl. Concentric AR2023), and TBK's audited FY3/2026 results  |  June 2026", size=12, color=RGBColor(0xA8, 0xB8, 0xC8))

# ---------- 2 Brakes market ----------
s = slide()
header(s, "Brakes: drum brakes are a shrinking island",
       "Verified findings from peer-reviewed literature, regulatory records, and industry data")
panel(s, 0.5, 1.55, 4.0, 3.1, "Europe — converted ~2000",
      "1996 Actros: first heavy truck with\nall-axle EBS disc brakes\n~70% disc/EBS take rate by 2000\nNear-universal on-highway since\nDrums: ~18% of EU demand (off-road)")
panel(s, 4.7, 1.55, 4.0, 3.1, "North America — tipped 2023-24",
      "FMVSS 121 (2009): −30% stopping\ndistance, drum-compatible by design\nTractors: <5% disc (2011) → ~25%\n(2018) → ~50%+ (2023-24)\nTrailers still lag at ~15%")
panel(s, 8.9, 1.55, 3.95, 3.1, "Japan — the last drum market",
      "Drums still majority on MD/HD trucks\nUD Quon all-disc since 2017\nNew Giga tractor: all-wheel discs\nIsuzu export heavies all-disc (MY25)\nDomestic flip is when, not if", title_color=RED)
panel(s, 0.5, 4.85, 6.1, 2.3, "Why discs won (peer-reviewed)",
      "Fade resistance under thermal load (SAE 902206, 1990; physics in Day et al. 1984)\n~30% shorter stops vs standard drums (NHTSA simulator, 108 drivers)\n10-20% braking advantage, 12-80% lifecycle savings in demanding duty (Ampadu 2023)\nEBS/AEB/ADAS mandates favor disc actuation precision")
panel(s, 6.8, 4.85, 6.05, 2.3, "Electrification re-shuffles the deck",
      "Regen braking cuts friction wear 64-95% (multiple peer-reviewed studies)\nFriction aftermarket growth ~2% CAGR with EV headwind\nEU brake-dust limits for heavy vehicles expected ~2030\nValue shifts: wear parts → corrosion/NVH/low-dust + retarders & blending", title_color=GREEN)

# ---------- 3 Pumps market ----------
s = slide()
header(s, "Pumps: electrification grows the category — and reallocates it",
       "Verified findings: peer-reviewed thermal studies, regulatory texts, and adversarial company documents")
panel(s, 0.5, 1.55, 4.0, 3.1, "Diesel era — refined commodity",
      "Pumps = only ~1-2% of fuel energy\nE-coolant pump on 13L truck engine\nsaves just 0.71-0.94% (TU Graz 2021)\nVariable pumps: ~1-1.5% verified vs\n3-6% in industry marketing\nSmall lever → slow innovation")
panel(s, 4.7, 1.55, 4.0, 3.1, "Regulation — Japan pushed least",
      "US GEM: codified credits for\nelectric pumps (0.5-1 g/ton-mi)\nEU VECTO: pump tech classes;\n62.5% of fleet still fixed-displ.\nJapan: NO accessory-level signal —\nprotected short-term, atrophied long-term", title_color=RED)
panel(s, 8.9, 1.55, 3.95, 3.1, "ZEV trucks — content multiplies",
      "BEV: 2-4 coolant loops vs 1 pump\n(battery <40°C, motor ~85°C)\nFCEV: 40-45% of fuel energy out\nvia coolant; 1.5-2x heat exchanger\nHV e-pump ≈ 4-8x unit value\n→ content/vehicle ~8-30x diesel", title_color=GREEN)
panel(s, 0.5, 4.85, 6.1, 2.3, "The adversarial record on ICE pumps",
      "Rheinmetall sold Pierburg (€2bn rev) for €350M ≈ 0.18x sales, after ~€550M impairments\nDayco carved out OE engine components; SHW cites Asian entrants + price-downs\nHanon forced to unwind flattering R&D capitalization\nEC antitrust (M.8102): pump niches at 90-100% concentration — sockets are sticky", title_color=RED)
panel(s, 6.8, 4.85, 6.05, 2.3, "Japan: slow melt, early sockets",
      "TBK's own filing: EV shift in buses/small-medium trucks may cut pump demand\nBEV truck volumes tiny today (Dutro Z EV: ~400 units yr 1) — diesel persists past 2030\nBUT: Mikuni already produces e-oil pumps for a domestic small BEV truck (2022)\nSockets are awarded years before volume (Pierburg order book runs 2028-35)", title_color=GREEN)

# ---------- 4 Same customers ----------
s = slide()
header(s, "Same customers — two different purchase decisions",
       "Isuzu, Fuso, Hino, Komatsu buy both product lines, but through different people, cycles, and triggers")
panel(s, 0.5, 1.55, 6.1, 4.0, "BRAKES — safety & chassis domain",
      "Decision owners: chassis/safety engineering + homologation, OEM-wide platform decisions\n"
      "Trigger: regulation (stopping distance, EBS/AEB, brake dust) and platform renewals — spec flips are abrupt and fleet-wide (Europe converted in ~5 yrs)\n"
      "Cycle: long platform homologation; switching mid-cycle is rare → incumbent until the flip, then winner-takes-platform\n"
      "Revenue model: aftermarket-heavy (wear parts annuity) — but discs + regen shrink friction events\n"
      "TBK sales motion: defend drum platforms while qualifying ADB for the domestic disc flip; access EBS/systems layer via Brakes India alliance", body_size=12)
panel(s, 6.8, 1.55, 6.05, 4.0, "PUMPS — powertrain & thermal domain",
      "Decision owners: engine/e-powertrain and thermal-architecture teams, decided at program kick-off per powertrain program\n"
      "Trigger: efficiency/cost in ICE era (weak in Japan — no regulatory credit); now EV thermal architecture awards 2-4 sockets at once\n"
      "Cycle: sockets fixed years before volume (orders running 2028-35 signed now); concentrated niches — share lost is rarely regained\n"
      "Revenue model: OE-heavy, little aftermarket (pumps last engine life) — value is in content growth per vehicle, not wear parts\n"
      "TBK sales motion: convert dominant mechanical-pump incumbency into e-pump/TCU sockets on Isuzu/Fuso/Hino EV & FC programs before Mikuni/Yamada lock them", body_size=12, title_color=GREEN)
txt(s, 0.5, 5.75, 12.3, 1.5, "Implication: one customer list, two sales motions and two R&D programs. Brake timing is regulatory/platform-driven (defend + be ready for the flip); pump timing is program-driven and running NOW (sockets being awarded while diesel volume still pays the bills). TBK must be present in both rooms at the same OEMs — chassis engineering for ADB, thermal/e-powertrain teams for e-pumps — which argues for ring-fenced teams and budgets per domain, not one shared engineering pool.", size=13, bold=True, color=DARK)

# ---------- 5 Innovation ----------
s = slide()
header(s, "Competitors kept innovating — in brakes and in pumps",
       "Every relevant peer has repositioned its product line for discs, electronics, and electrification")
panel(s, 0.5, 1.55, 6.1, 5.6, "BRAKES — from castings to systems",
      "Knorr-Bremse/Bendix: ADB standard on Peterbilt/Kenworth steer axles (2013); ~20M discs in field; Global Scalable Brake Control platform; driverless-truck braking (ATLAS-L4)\n"
      "ZF/Wabco: EBS + ADAS integration; $7bn acquisition built the #2 systems house\n"
      "Brembo: Sensify brake-by-wire; 6.1% of sales into R&D\n"
      "Akebono: electro-mechanical brakes (EMB) + copper-free friction — spending up despite losses\n"
      "SAF-Holland/Haldex: trailer ADB + EBS at low cost\n"
      "Consolidation priced the lesson: Haldex sold at HALF Knorr-Bremse's blocked 2017 offer after staying sub-scale", body_size=11.5)
panel(s, 6.8, 1.55, 6.05, 5.6, "PUMPS — mechanical → electric thermal management",
      "Aisin: record ¥250bn R&D plan; e-water/oil pumps + coolant valves inside the eAxle strategy\n"
      "Mahle: thermal-management modules; €1.2bn BEV module order (2024) — largest in company history\n"
      "Hanon: BEV heat pumps, CO2 refrigerant, e-coolant pumps (~4-6% of sales)\n"
      "Mikuni (3x TBK's size): e-oil pump ALREADY in production for a Japanese small BEV truck\n"
      "Yamada: e-water pumps shipping in Honda EVs/FCEVs (2024)\n"
      "Concentric: e-products 20% of 2023 sales (AR2023), 25% gross margin on focused niches; Pierburg: 400-800V e-coolant pump families\n"
      "The pattern: every pump peer redirected R&D to electric pumps & BEV thermal — the exact adjacency open to TBK", body_size=11.5, title_color=GREEN)

# ---------- 6 Spend ----------
s = slide()
header(s, "What competitors spend on R&D",
       "Latest fiscal year, group level (no peer discloses brake-only or pump-only R&D); USD indicative")
cd = CategoryChartData()
cd.categories = ["SAF-Holland", "TBK", "Concentric*", "Cummins", "Akebono", "Aisin", "Hanon*", "Mahle", "Mikuni", "Brembo", "Knorr-Bremse", "ZF"]
cd.add_series("R&D % of sales", (1.7, 2.2, 2.3, 4.1, 4.3, 4.8, 5.0, 5.4, 5.9, 6.1, 6.8, 8.6))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.6), Inches(7.6), Inches(4.4), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = True
ch.chart_title.text_frame.text = "R&D intensity (% of sales)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
plot = ch.plots[0]
plot.has_data_labels = True
plot.data_labels.number_format = '0.0"%"'
plot.data_labels.number_format_is_linked = False
plot.data_labels.font.size = Pt(10)
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = DARK
ch.category_axis.tick_labels.font.size = Pt(10)
panel(s, 8.4, 1.6, 4.45, 4.4, "Absolute spend (≈USD)",
      "ZF €3.6bn (~$3.9bn)\nAisin ¥237bn (~$1.6bn)\nCummins $1.46bn\nMahle €607M (~$655M)\nKnorr-Bremse €544M (~$590M)\nHanon ~$360-440M\nBrembo ~$250M\nSAF-Holland €39M (~$42M)\nMikuni ¥5.5bn (~$37M)\nConcentric MSEK 95 (~$9M)*\nTBK ¥1.2bn (~$8M)", body_size=12)
txt(s, 0.5, 6.25, 12.3, 1.1, "Brake systems players spend 6-9% of sales; large pump/thermal players cluster at 4-5.5%; Akebono now at 4.3% (EMB + copper-free friction ramp, per consistent peer dataset). TBK sits at 2.2%. *Concentric (AR2023, verified): expensed product development only 2.3% — yet 25% gross margin: niche dominance, not raw R&D scale, drives pump profitability. *Hanon mid-range of conflicting filings.", size=11.5, color=GREY)

# ---------- 7 Capitalized R&D: the knowledge-capital stock ----------
s = slide()
header(s, "The stock, not the flow: competitors' capitalized R&D",
       "R&D Capital Base = cumulative R&D capitalized and amortized on a consistent basis (peer dataset, 2025, USD) — the knowledge asset each company has built")
cd = CategoryChartData()
cd.categories = ["TBK", "Akebono", "SAF-Holland", "Mikuni", "Knorr-Bremse", "Cummins"]
cd.add_series("R&D Capital Base ($M)", (43, 123, 131, 194, 2569, 5594))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.6), Inches(7.3), Inches(4.6), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = True
ch.chart_title.text_frame.text = "Capitalized R&D stock, US$M (2025)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
plot = ch.plots[0]
plot.has_data_labels = True
plot.data_labels.number_format = '#,##0'
plot.data_labels.number_format_is_linked = False
plot.data_labels.font.size = Pt(11)
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = DARK
ch.category_axis.tick_labels.font.size = Pt(11)
panel(s, 8.1, 1.6, 4.75, 2.3, "The stock gap",
      "Cummins $5.6bn — 130x TBK\nKnorr-Bremse $2.6bn — 60x TBK\nMikuni $194M — 4.5x TBK\nSAF-Holland $131M | Akebono $123M\nTBK: $43M — the smallest knowledge\nbase in the peer set", body_size=12.5)
panel(s, 8.1, 4.05, 4.75, 2.3, "Knowledge intensity (RDCB / sales)",
      "Knorr-Bremse 29% | Mikuni 29%\nCummins 17%\nTBK 12% | Akebono 12%\nSAF-Holland 7%\nThe systems winners run structurally\nknowledge-intensive business models", body_size=12.5, title_color=GREEN)
txt(s, 0.5, 6.35, 12.3, 1.0, "Why it matters: the R&D gap is cumulative, not annual. A decade of under-spending leaves TBK rebuilding a depleted asset while competitors amortize billions of accumulated know-how into every bid (brake control software, ADB iterations, e-pump motor/inverter design). This is why the catch-up phase must run YEARS at elevated spend — one good budget year cannot rebuild a stock — and why partnering (Brakes India) to borrow an existing knowledge base is rational.", size=11.5, color=GREY)

# ---------- 8 The argument: gross margin is too low ----------
s = slide()
header(s, "The case that TBK's gross margin is structurally too low",
       "TBK FY3/2026: 12.5% gross margin, 2.7% operating margin — four independent lines of evidence say this is a fixable defect, not the nature of the business")
panel(s, 0.5, 1.55, 6.1, 2.35, "1. Peers prove the products can carry more",
      "Every comparable brake peer except Akebono earns 22-27% (Haldex 27.3%, Cummins 24.7%, SAF-Holland ~22%)\n"
      "Pump peers: Concentric 25.2% (verified AR2023), TPR 21.7% — the high-share niche players, i.e. exactly TBK's structural position in Japanese truck pumps\n"
      "Same product categories, same customers' industry — 5-15 points more margin", body_size=11.5)
panel(s, 6.8, 1.55, 6.05, 2.35, "2. TBK earns LESS than its own customers",
      "Isuzu gross margin ~19%, Hino ~17% vs TBK 12.5%\n"
      "Healthy specialist suppliers out-margin their OEMs (Knorr-Bremse 13% EBIT vs truck OEMs' mid-single digits)\n"
      "Inverted economics = pricing power sits with the customer: keiretsu concentration (Isuzu alone ¥12bn = 22% of sales) and annual price-downs, against commodity drum/mechanical products with no systems, electronics or branded-aftermarket content to defend price", body_size=11.5, title_color=RED)
panel(s, 0.5, 4.05, 6.1, 2.35, "3. The COGS line carries an underloaded footprint",
      "PP&E at 39% of sales — 3x Concentric's ~12%; depreciation 5.3% of sales\n"
      "China: depreciation 12.8% of segment sales, losses both years; Japan segment margin just 1.6%\n"
      ">¥2.2bn impairments + restructuring losses in 2 years = the cost of underutilized plants crystallizing (booked below the operating line, so the true factory burden is even larger than the margin gap shows)", body_size=11.5)
panel(s, 6.8, 4.05, 6.05, 2.35, "4. The repair has already started — and works",
      "FY3/2026: gross margin 10.6% → 12.5% in one year (cost of sales cut ¥0.7bn on flat sales); North America exited; ¥712M Japan impairment = footprint action\n"
      "Concentric's playbook validates the destination: focused niches, 25% gross margin, opex GROWING into the e-transition at 18.4% average operating margin\n"
      "Nothing about brakes or pumps caps TBK at 12.5% — peers with the same products price and load their factories better", body_size=11.5, title_color=GREEN)
txt(s, 0.5, 6.5, 12.3, 0.9, "Conclusion: at peer margins TBK's existing sales generate $22-24M more gross profit per year (next pages) — enough to fund the entire R&D catch-up. The margin is the constraint, and it is self-inflicted: underloaded factories in COGS, commodity mix, and customer-held pricing power — all addressable.", size=13, bold=True, color=DARK)

# ---------- 8 Margin gap per business (USD) ----------
s = slide()
header(s, "TBK underperforms peer gross margins in BOTH businesses",
       "TBK audited FY3/2026 (blended 12.5%) vs comparable peers per panel  |  all figures USD at ¥150/USD")
cd = CategoryChartData()
cd.categories = ["BRAKES\n(TBK sales ~$127M)", "PUMPS & ENGINE COMP.\n(TBK sales ~$238M)"]
cd.add_series("TBK gross margin", (12.5, 12.5))
cd.add_series("Peer average", (21.0, 17.0))
cd.add_series("Peer median", (23.4, 16.9))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.6), Inches(7.3), Inches(4.6), cd)
ch = gf.chart
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(12)
ch.has_title = True
ch.chart_title.text_frame.text = "Gross margin: TBK vs peers, by business (%)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
plot = ch.plots[0]
plot.has_data_labels = True
plot.data_labels.number_format = '0.0"%"'
plot.data_labels.number_format_is_linked = False
plot.data_labels.font.size = Pt(11)
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = RED
plot.series[1].format.fill.solid()
plot.series[1].format.fill.fore_color.rgb = DARK
plot.series[2].format.fill.solid()
plot.series[2].format.fill.fore_color.rgb = GREEN
ch.category_axis.tick_labels.font.size = Pt(11)
panel(s, 8.1, 1.6, 4.75, 2.3, "Brake peers (comparable basis)",
      "Haldex 27.3% | Cummins 24.7%\nSAF-Holland ~22% | Akebono 10.0%\nTBK trails every brake peer except\nAkebono — by 9-11 points", body_size=12.5)
panel(s, 8.1, 4.05, 4.75, 2.3, "Pump peers (comparable basis)",
      "Concentric 25.2% (verified AR2023)\nTPR 21.7% | Aisin 12.1% | Hanon 9.1%\nTBK trails the panel average by\n~4.5 points — and the pure-play\nCV pump comp by ~13 points", body_size=12.5, title_color=GREEN)
txt(s, 0.5, 6.35, 12.3, 0.9, "TBK does not disclose gross profit per division, so its blended 12.5% is shown against both panels. Even on the most charitable reading — assuming all underperformance sits in one business — TBK is below peer average in the other. Peer set: IFRS cost-of-sales / JGAAP reporters only.", size=11.5, color=GREY)

# ---------- 8 Value of closing the gap (USD) ----------
s = slide()
header(s, "Closing the margin gap is worth ~$22-24M more gross profit per year",
       "TBK FY3/2026 sales ($365M) at peer margins  |  USD at ¥150/USD  |  division split estimated from FY3/2025 (35% / 65%)")
cd = CategoryChartData()
cd.categories = ["Actual\nFY3/2026", "At peer\naverage", "At peer\nmedian"]
cd.add_series("Brakes ($M)", (None, 26.6, 29.5))
cd.add_series("Pumps & engine components ($M)", (None, 40.5, 40.3))
cd.add_series("Actual blended ($M)", (45.5, None, None))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(0.5), Inches(1.6), Inches(6.8), Inches(4.6), cd)
ch = gf.chart
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(11)
ch.has_title = True
ch.chart_title.text_frame.text = "TBK annual gross profit (US$M)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
p0 = ch.plots[0]
p0.has_data_labels = True
p0.data_labels.number_format = '$0.0"M"'
p0.data_labels.number_format_is_linked = False
p0.data_labels.font.size = Pt(10)
p0.series[0].format.fill.solid()
p0.series[0].format.fill.fore_color.rgb = DARK
p0.series[1].format.fill.solid()
p0.series[1].format.fill.fore_color.rgb = GREEN
p0.series[2].format.fill.solid()
p0.series[2].format.fill.fore_color.rgb = RED
ch.category_axis.tick_labels.font.size = Pt(11)
panel(s, 7.6, 1.6, 5.25, 2.0, "The annual prize",
      "+$21.6M/yr at peer AVERAGE margins\n+$24.4M/yr at peer MEDIAN margins\n≈ +50% on today's $45.5M gross profit\nSplit roughly half brakes, half pumps (~$11-14M each)", body_size=13)
panel(s, 7.6, 3.75, 5.25, 2.6, "What it funds",
      "The entire R&D catch-up program costs ~$27M/yr\n(parity run-rate $17M + deficit repayment $10M)\n→ closing ~90% of the margin gap pays for ALL of it\n\nFY3/2026 already delivered the first ~$7M/yr\n(margin 10.6% → 12.5%) — the path is proven, not theoretical", body_size=13, title_color=GREEN)
txt(s, 0.5, 6.35, 12.3, 0.9, "Method: brakes ~$127M sales x peer avg 21.0% / median 23.4%; pumps & engine components ~$238M x 17.0% / 16.9%; vs actual blended 12.5%. Margin levers per the benchmark: aftermarket/branded content, niche pricing power (Concentric/TPR model), price-down discipline, and mix shift to disc brakes and e-pumps.", size=11.5, color=GREY)

# ---------- 9 OPEX comparison (USD) ----------
s = slide()
header(s, "Why TBK looks lean on OPEX: its factory costs hide in COGS",
       "OPEX here = gross margin − operating margin (SG&A + R&D). Under JGAAP, factory overhead, depreciation and idle capacity sit in COST OF SALES  |  USD at ¥150/USD")
cd = CategoryChartData()
cd.categories = ["BRAKES panel", "PUMPS panel"]
cd.add_series("TBK (blended)", (9.7, 9.7))
cd.add_series("Peer average", (13.8, 10.4))
cd.add_series("Peer median", (13.6, 9.3))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.6), Inches(7.3), Inches(4.6), cd)
ch = gf.chart
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(12)
ch.has_title = True
ch.chart_title.text_frame.text = "OPEX as % of sales: TBK vs peers, by business"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
plot = ch.plots[0]
plot.has_data_labels = True
plot.data_labels.number_format = '0.0"%"'
plot.data_labels.number_format_is_linked = False
plot.data_labels.font.size = Pt(11)
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = RED
plot.series[1].format.fill.solid()
plot.series[1].format.fill.fore_color.rgb = DARK
plot.series[2].format.fill.solid()
plot.series[2].format.fill.fore_color.rgb = GREEN
ch.category_axis.tick_labels.font.size = Pt(11)
panel(s, 8.1, 1.6, 4.75, 2.3, "The lean look is an artifact",
      "TBK's 9.7% covers only SG&A + its\nthin ~2.2% R&D — the future-building costs\nit has starved. The factory burden is in COGS:\nPP&E 39% of sales vs Concentric ~12% (3x)\nDepreciation 5.3% of sales; China: 12.8% of\nsegment sales + losses both years", body_size=11.5)
panel(s, 8.1, 4.05, 4.75, 2.3, "Restructuring bypasses BOTH metrics",
      "Impairments ¥459M + ¥712M and\nrestructuring losses ¥775M + ¥297M\n(>¥2.2bn in 2 years) booked as JGAAP\nextraordinary items — below operating\nprofit, invisible to GM and OPEX alike.\nThat IS the underutilized-footprint cost", body_size=11.5, title_color=RED)
txt(s, 0.5, 6.35, 12.3, 1.0, "Two-front program, mapped to the P&L: (1) footprint/utilization repair attacks COGS — the $22-24M/yr gross-margin prize (North America exit and the ¥712M Japan impairment show it has started); (2) R&D catch-up deliberately GROWS opex toward peer levels (peers spend 10-14% incl. 4%+ R&D and earn 2.5x TBK's operating margin — Concentric grew opex 10.3%→11.7% funding its e-pump transition at 18.4% avg OPM). Cut above the line, invest below it.", size=11.5, color=GREY)

# ---------- 10 Deficit ----------
s = slide()
header(s, "TBK's accumulated R&D deficit: ~US$100M over a decade",
       "TBK dataset FY2016-2025; benchmark = blended peer floor for a brake + pump supplier")
cd = CategoryChartData()
cd.categories = [str(y) for y in range(2017, 2026)]
cd.add_series("TBK actual R&D ($M)", (11.1, 11.4, 10.6, 11.4, 11.6, 11.8, 8.9, 8.1, 7.8))
cd.add_series("At 4.5% of sales ($M)", (19.4, 20.9, 21.8, 21.3, 18.7, 20.5, 17.8, 17.7, 16.1))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.6), Inches(7.6), Inches(4.4), cd)
ch = gf.chart
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(11)
ch.has_title = True
ch.chart_title.text_frame.text = "Actual R&D vs peer-floor benchmark (US$M)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
p0 = ch.plots[0]
p0.series[0].format.fill.solid()
p0.series[0].format.fill.fore_color.rgb = RED
p0.series[1].format.fill.solid()
p0.series[1].format.fill.fore_color.rgb = DARK
ch.category_axis.tick_labels.font.size = Pt(10)
panel(s, 8.4, 1.6, 4.45, 4.4, "The deficit, quantified",
      "Cumulative sales FY16-25: $4.26bn\nCumulative R&D: $93.5M (2.2%)\n\nGap vs 4.0% benchmark: ~$77M\nGap vs 4.5% benchmark: ~$98M\nGap vs 5.0% benchmark: ~$120M\n\nCentral estimate: ~US$100M\n\nAnd intensity is FALLING:\n2.6% (FY22) → 2.2% (FY25)", body_size=12.5, title_color=RED)
txt(s, 0.5, 6.25, 12.3, 1.0, "Compounding the problem: TBK's marginal return on invested capital has been negative on every horizon (-2% to -10%) — the deficit is in direction as well as amount. Capex fell to $16M by FY25, though FY3/26 shows a turn: capex back up to ¥3.3bn and ¥1.1bn equity raised from Brakes India.", size=12, color=GREY)

# ---------- 11 Plan ----------
s = slide()
header(s, "The reinvestment requirement: two phases",
       "Catch-up to repay the technological deficit, then a permanently higher run-rate to match peers")
cd = CategoryChartData()
cd.categories = ["FY26\nactual est.", "FY27", "FY28", "FY29", "FY30", "FY31", "Steady state\nFY32+"]
cd.add_series("Parity run-rate (~4.5-5% of sales)", (7.8, 17, 17, 17, 17, 17, 17))
cd.add_series("Deficit repayment (catch-up)", (0, 10, 10, 10, 10, 10, 0))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(0.5), Inches(1.6), Inches(7.6), Inches(4.4), cd)
ch = gf.chart
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(11)
ch.has_title = True
ch.chart_title.text_frame.text = "R&D spend plan (US$M per year)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
p0 = ch.plots[0]
p0.series[0].format.fill.solid()
p0.series[0].format.fill.fore_color.rgb = DARK
p0.series[1].format.fill.solid()
p0.series[1].format.fill.fore_color.rgb = GREEN
ch.category_axis.tick_labels.font.size = Pt(9)
panel(s, 8.4, 1.6, 4.45, 4.4, "The numbers",
      "PHASE 1 — Catch-up (5 yrs):\n~$27M/yr ≈ 7.5% of sales\n= parity run-rate ($16-18M)\n+ deficit repayment (~$10M/yr,\n~$50M split between ADB\nindustrialization and e-pump/TCU\nprograms)\nTotal: ~$135M over 5 years\n\nPHASE 2 — Steady state:\n$16-18M/yr ≈ 4.5-5% of sales\n= 2.2x today's spend, forever\n\nPlus one-off ADB line capex\n(~$30-50M, outside R&D line)", body_size=12, title_color=GREEN)
txt(s, 0.5, 6.25, 12.3, 1.1, "Allocation discipline, per domain: (1) brakes — ADB industrialization for the Japanese disc flip, systems/ADAS access via Brakes India, timed to regulation and platform renewals; (2) pumps — e-pump/TCU sockets on Isuzu/Fuso/Hino electrified programs, timed to program kick-offs happening NOW. Funding: the ~$22-24M/yr peer-margin prize covers the ~$27M/yr program almost fully. The Concentric lesson: focused niche engineering at 2.3% intensity earns 25% gross margins — spend must buy sockets, not just scale.", size=11.5, color=GREY)

# ---------- 14 Research quality notes ----------
s = slide()
header(s, "Research quality & methodology notes",
       "How this material was produced, verified, and where its limits are")
panel(s, 0.5, 1.55, 6.1, 2.6, "Method & confidence tiers",
      "Two deep-research studies (brake market; pump market): 5 parallel search angles each, claims extracted as falsifiable statements, then ADVERSARIAL VERIFICATION — 3 independent voters per claim set, actively attempting refutation\n"
      "Tier 1: peer-reviewed papers, primary regulation, audited filings (TBK tanshin, Concentric AR2023, SEC)\n"
      "Tier 2: government data and multi-source-verified company/industry figures\n"
      "Tier 3: single-source estimates and market research — always labeled in the underlying documents", body_size=11)
panel(s, 6.8, 1.55, 6.05, 2.6, "Corrections made by verification (transparency)",
      "TBK's famous '~70% share' belongs to its WATER PUMPS, not brakes (brakes: 'top share', unaudited)\n"
      "Euro 7 heavy-duty brake-dust limits apply ~2030, not 2028/29 (those are exhaust dates)\n"
      "Concentric R&D = 2.3% of sales per AR2023 — the circulating '~5%' was wrong\n"
      "'85-90% of Asia on drums' could not be corroborated — stated as 'majority' only\n"
      "A claimed TBK-Mikuni alliance is erroneous (conflation with Mikuni-Suzuki)\n"
      "Aggregator's 'TBK FY26 ¥138M profit' contradicted by the audited tanshin (-¥131M)", body_size=11, title_color=RED)
panel(s, 0.5, 4.3, 6.1, 2.6, "Known limitations",
      "Full-text retrieval was blocked in the research environment: paywalled SAE/journal numbers rest on abstracts; flagged where unverified\n"
      "No peer discloses brake-only or pump-only R&D, gross profit, or segment margins — group figures used; ratios are the valid comparison\n"
      "Accounting bases differ: German/Italian nature-of-expense 'gross margins' (Knorr ~50%, Brembo ~50%) are materials-only and EXCLUDED from margin comparisons; JGAAP fully expenses R&D while IFRS peers capitalize part\n"
      "TBK FY3/26 division split estimated from FY3/25 (yuho due June 2026); USD at ¥150/USD", body_size=11)
panel(s, 6.8, 4.3, 6.05, 2.6, "Data hierarchy & framework constructs",
      "Where sources conflict, the audited primary document wins (e.g., TBK tanshin over aggregators)\n"
      "Consistent-basis peer metrics (margins, R&D, ROIC) come from the user-provided 10-year/LTM peer dataset — same methodology across all 7 companies\n"
      "ROIIC (ROICm) = return on incremental invested capital over the stated horizon; RR = reinvestment rate; R&D Capital Base = cumulative R&D capitalized/amortized per the dataset framework — a modeling construct, not a balance-sheet line\n"
      "Scenario pages state assumptions on-slide; peer-margin scenarios use comparable reporters only", body_size=11, title_color=GREEN)

# ---------- 15 Sources ----------
s = slide()
header(s, "Sources")
panel(s, 0.5, 1.3, 6.1, 2.85, "Primary & audited",
      "TBK Co., Ltd. — Consolidated Financial Results FY ended 31 Mar 2026 (tanshin, 14 May 2026); securities-report risk factors (有価証券報告書 via kitaishihon/irbank)\n"
      "Concentric AB — Annual Report 2023 (income statement, balance sheet, 5-year key figures)\n"
      "User-provided datasets: TBK 10-year financials; Peers 10-year; Peers LTM (7 companies, consistent methodology)\n"
      "SEC filings: Cummins 10-K/10-Q/8-K (FY2024-25); WABCO & Meritor merger 8-Ks\n"
      "Company releases: Knorr-Bremse FY2023-25; ZF AR2024-25; SAF-Holland AR2024; Mahle AR2024-25; Haldex year-end 2021; Rheinmetall FY2025 + AEQUITA divestment (3 Jun 2026); Circle BidCo/A.P. Møller offer documents (2024); Brakes India-TBK alliance (Dec 2025)", body_size=10)
panel(s, 6.8, 1.3, 6.05, 2.85, "Peer-reviewed & academic",
      "Göhring & von Glasner, SAE 902206 (1990); Kajiyama et al. (Isuzu/TBK), SAE 902200 (1990); Day, Harding & Newcomb, Proc IMechE (1984); Hoover et al., SAE 2005-01-3614; Subel & Kienhöfer, Proc IMechE D (2020)\n"
      "Ampadu, Alrejjal & Ksaibati, JSDTL 8(1) 2023 — disc vs drum cost-effectiveness\n"
      "Granitz et al., SN Applied Sciences (2021) — e-coolant pump, 13L HD diesel; Bitsis & Miwa, SAE 2018-01-0980; Cummins SuperTruck 2, SAE 2019-01-0247; MIT S.M. thesis (2012) — HD parasitic losses\n"
      "Zhang et al., J. Hazardous Materials 465 (2024); Hagino, Atmosphere 15:75 (2024, JARI); Environmental Pollution (2023) — regen braking & brake wear\n"
      "Energies 16:4024 (2023), 19:1748 (2026); Sustainability 15:3132 (2023); Energies 16:168 (2023) — BEV/FCEV truck & bus thermal management", body_size=10)
panel(s, 0.5, 4.35, 6.1, 2.85, "Regulatory & legal",
      "NHTSA: FMVSS 121 final rule 74 FR 37122 (2009) + Regulatory Impact Analysis; DOT HS 811 367 (2010, NADS simulator); ESV 07-0242 (2007); heavy-vehicle AEB NPRM (2023)\n"
      "UNECE Regulation No. 13 (braking, Type-II fade tests)\n"
      "EU: Reg 2019/1242 & 2024/1610 (HDV CO2); Reg 2024/1257 (Euro 7); Reg 2017/2400 Annex IX (VECTO auxiliaries) + JRC112015; EPA Phase 2 (40 CFR 1037.520 accessory credits) & Phase 3 (2024)\n"
      "EC merger decisions: M.8102 Valeo/FTE (2017, 90-100% concentration, Raicam remedy); Knorr-Bremse/Haldex Phase II (IP-17-2126, 2017)\n"
      "Japan: MY2025 HDV fuel-economy standards (via ICCT/TransportPolicy); METI supplier-transition studies (2023, 2025); FMCSA Crash Cost Methodology (2025); AAA Foundation (2017)", body_size=10)
panel(s, 6.8, 4.35, 6.05, 2.85, "Industry & financial data (verified, labeled where single-source)",
      "Adoption data: Haldex ADB history; Bendix/Knorr production milestones; FleetOwner, TruckingInfo, Fleet Maintenance, Transport Topics, Fleet Equipment; NACFE confidence reports; NAS 25542; DOE SuperTruck (OSTI)\n"
      "Deal records: Knorr-Bremse-Bendix (2002); ZF-Wabco $7bn (2020); Cummins-Meritor $3.7bn (2022); SAF-Holland-Haldex (2022); Concentric take-private SEK 230 (2024); Pierburg-AEQUITA €350M (2026)\n"
      "Japanese filings via irbank, kitaishihon, buffett-code, kabutan, MarkLines (incl. TBK JSAE 2022 exhibit; Mikuni 2022 BEV-truck e-oil pump disclosure; Yamada e-water pump adoptions)\n"
      "Aggregators for computed ratios: macrotrends, stockanalysis, Simply Wall St — cross-checked against filings; market research (MarketsandMarkets, GVR, IDTechEx) used directionally only", body_size=10)

import warnings
with warnings.catch_warnings():
    warnings.simplefilter("error")
    prs.save("/home/user/Claude/TBK_rd_gap_slides.pptx")
print("saved", len(prs.slides._sldIdLst), "slides cleanly")
