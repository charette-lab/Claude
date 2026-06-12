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
cd.categories = ["SAF-Holland", "TBK", "Concentric*", "Cummins", "Aisin", "Hanon*", "Mahle", "Mikuni", "Brembo", "Knorr-Bremse", "ZF"]
cd.add_series("R&D % of sales", (2.1, 2.2, 2.3, 4.3, 4.8, 5.0, 5.4, 5.4, 6.1, 6.9, 8.6))
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
txt(s, 0.5, 6.25, 12.3, 1.1, "Brake systems players spend 6-9% of sales; large pump/thermal players cluster at 4-5.5%. TBK sits at 2.2%. *Concentric (AR2023, verified): expensed product development only 2.3% — yet 25% gross margin: niche dominance, not raw R&D scale, drives pump profitability. The catch-up case rests on the larger peers and on winning specific e-pump sockets. *Hanon mid-range of conflicting filings.", size=11.5, color=GREY)

# ---------- 7 Margin gap per business (USD) ----------
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
header(s, "OPEX: TBK's problem is not cost discipline — it under-spends",
       "OPEX = gross margin − operating margin (captures SG&A + R&D + other operating items), comparable peer set  |  USD at ¥150/USD")
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
panel(s, 8.1, 1.6, 4.75, 2.3, "Peers spend MORE below the line…",
      "Brakes: Haldex 19.9%, Cummins 13.7%\n(SG&A 9.6% + R&D 4.3%), SAF-H 13.4%,\nAkebono 8.1%\nPumps: TPR 16.3%, Concentric 11.0%,\nAisin 7.6%, Hanon 6.6%\nTBK SG&A (incl. its ~2.2% R&D): 9.7%", body_size=11.5)
panel(s, 8.1, 4.05, 4.75, 2.3, "…and still earn 2.5x the margin",
      "Operating margin: TBK 2.7% vs\nbrake peers avg 7.2% | pump peers avg 6.7%\nAt peer GM AND peer opex, TBK would\nspend ~$7M MORE opex ($42M vs $35M)\nyet earn ~$25M operating profit vs\n$10M actual → +$15M/yr", body_size=11.5, title_color=GREEN)
txt(s, 0.5, 6.35, 12.3, 1.0, "Concentric's history makes the point: opex rose 10.3% (2022) → 11.7% (2023) as it funded the e-pump transition, yet 5-yr operating margin averaged 18.4% (23.5% pre-EMP in 2019) — high gross margins pay for development. Implication for TBK: cutting opex is the wrong lever; the gap is in COGS/gross margin, and peer-level opex means spending MORE below the line — exactly the R&D catch-up program.", size=11.5, color=GREY)

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

import warnings
with warnings.catch_warnings():
    warnings.simplefilter("error")
    prs.save("/home/user/Claude/TBK_rd_gap_slides.pptx")
print("saved", len(prs.slides._sldIdLst), "slides cleanly")
