"""Staff-facing Performance Compensation Framework -> PDF (reportlab). v2 with Roles & Responsibilities."""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                HRFlowable, KeepTogether)

NAVY=colors.HexColor("#1F3864"); BLUE=colors.HexColor("#2E5496"); GREY=colors.HexColor("#595959")
OUT="/home/user/Claude/Performance_Compensation_Framework.pdf"

ss=getSampleStyleSheet()
H1=ParagraphStyle("H1",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=17,textColor=NAVY,spaceAfter=0,leading=20)
SUB=ParagraphStyle("SUB",parent=ss["Normal"],fontName="Helvetica",fontSize=13,textColor=BLUE,spaceAfter=2,leading=15)
MUT=ParagraphStyle("MUT",parent=ss["Normal"],fontName="Helvetica-Oblique",fontSize=10,textColor=GREY,spaceAfter=2)
HD=ParagraphStyle("HD",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=12.5,textColor=NAVY,spaceBefore=9,spaceAfter=3,leading=14)
SH=ParagraphStyle("SH",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=10.6,textColor=BLUE,spaceBefore=5,spaceAfter=2,leading=13)
BODY=ParagraphStyle("BODY",parent=ss["Normal"],fontName="Helvetica",fontSize=10.2,leading=13.2,spaceAfter=5,alignment=TA_LEFT)
RB=ParagraphStyle("RB",parent=BODY,spaceAfter=4,leftIndent=8)
BUL=ParagraphStyle("BUL",parent=BODY,leftIndent=14,bulletIndent=4,spaceAfter=2.5)
BOX=ParagraphStyle("BOX",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=10.5,textColor=NAVY,alignment=TA_CENTER,leading=15)
CELL=ParagraphStyle("CELL",parent=ss["Normal"],fontName="Helvetica",fontSize=8.6,leading=10.4)
CELLB=ParagraphStyle("CELLB",parent=CELL,fontName="Helvetica-Bold")
CELLH=ParagraphStyle("CELLH",parent=CELL,fontName="Helvetica-Bold",textColor=colors.white)
CELLT=ParagraphStyle("CELLT",parent=CELL,fontName="Helvetica-Bold",textColor=NAVY)
CELLW=ParagraphStyle("CELLW",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=10,textColor=BLUE,alignment=TA_CENTER)
FOOT=ParagraphStyle("FOOT",parent=ss["Normal"],fontName="Helvetica-Oblique",fontSize=8,textColor=GREY)

def B(t): return Paragraph(t,BODY)
def R(t): return Paragraph(t,RB)
def bullet(t,style=BUL): return Paragraph(("&bull;&nbsp;&nbsp;"+t),style)

story=[]
story+=[Paragraph("Athanase Industrial Partner",H1),
        Paragraph("Roles, Responsibilities &amp; Performance Compensation Framework",SUB),
        Paragraph("Staff guide to how we work and how the performance pool is created and shared",MUT),
        HRFlowable(width="100%",thickness=1.4,color=NAVY,spaceBefore=4,spaceAfter=7)]

# 1. Purpose
story+=[Paragraph("1.&nbsp; Purpose",HD),
 B("This document explains how we work at Athanase Industrial Partner &mdash; the roles in the firm and what each is "
   "responsible for &mdash; and how performance-based compensation is created and shared. Our philosophy is simple: "
   "when the fund performs, the people who create that performance share directly in it. The framework is designed to "
   "be transparent, predictable, and to reward both contribution and progression.")]

# 2. Roles and responsibilities
story+=[Paragraph("2.&nbsp; Roles and responsibilities",HD)]

story+=[Paragraph("2.1&nbsp; How we invest",SH),
 B("Athanase is an <b>engaged owner</b> of public companies. We run a concentrated portfolio and take meaningful "
   "ownership positions, then work actively to increase the value of each holding &mdash; through constructive "
   "engagement with boards and management on strategy, operations, capital allocation and governance. This makes our "
   "model closer to private equity than to a traditional trading desk: we create returns by improving businesses we "
   "own, not by trading in and out of them."),
 B("Every prospective investment is owned by a <b>deal team</b> made up of a Portfolio Manager and at least one "
   "Analyst. The deal team develops the thesis and the value-creation plan and presents it to the wider <b>investment "
   "group</b>. The group's role is to <b>pressure-test the investment premise</b> &mdash; challenging assumptions, "
   "surfacing risks and stress-testing the thesis from every angle. The <b>CIO</b> chairs these meetings, makes the "
   "final selection decision, and owns portfolio construction &mdash; position sizing and the pace of accumulation "
   "and divestment &mdash; working closely with the responsible Portfolio Manager and the Head of Research.")]

story+=[Paragraph("2.2&nbsp; The investment team",SH),
 R("<b>Chief Investment Officer (CIO).</b> Holds ultimate responsibility for the firm's investment performance. The "
   "CIO chairs the investment meetings and makes the final decision on which investments enter and leave the "
   "portfolio. The CIO owns <b>portfolio construction</b> &mdash; determining position sizes and the pace of "
   "accumulation and divestment &mdash; in close collaboration with the responsible Portfolio Manager and the Head of "
   "Research. The CIO sets the investment strategy and safeguards the discipline and consistency of our process."),
 R("<b>Portfolio Manager (PM).</b> Leads assigned investments. Unlike a traditional hedge fund, a PM at Athanase is "
   "<b>not allocated a discrete pool of capital to trade</b>. As an engaged owner, the PM acts much like a "
   "<b>private equity deal partner</b>: they take ownership of each assigned investment and are responsible for "
   "driving the improvement of the underlying portfolio company. This includes building and maintaining the thesis "
   "and value-creation plan, engaging with the company's board and management, pursuing operational, strategic, "
   "capital-allocation and governance change, and representing the firm in its ownership role. The PM leads the deal "
   "team and presents and defends the thesis to the investment group."),
 R("<b>Junior PM.</b> Supports and co-leads investments alongside a Portfolio Manager, taking ownership of defined "
   "workstreams within the engagement and value-creation agenda. The role bridges deep analytical work and full "
   "investment ownership, building the judgment and company-engagement experience needed to progress to PM."),
 R("<b>Head of Research.</b> A distinct senior role that leads the firm's research function and partners closely with "
   "the CIO on portfolio construction &mdash; helping shape conviction, position sizing and risk assessment across "
   "the portfolio &mdash; while overseeing the rigour and quality of our due diligence. This role is currently held "
   "on a voluntary basis and does not participate in the performance pool."),
 R("<b>Senior Analyst.</b> A lead analyst with deep sector or thematic coverage who generates original investment "
   "ideas, leads the due-diligence effort on assigned investments and mentors junior analysts. The senior analyst is "
   "the most experienced member of the research bench and a primary feeder to the Junior PM and PM roles."),
 R("<b>Analyst.</b> Responsible for the <b>confirmatory due diligence</b> that underpins each investment decision and "
   "for helping to <b>source new ideas</b>. Working within a deal team, the analyst builds the financial models, "
   "gathers and verifies evidence, tests the key premises of the thesis and monitors investments once owned. Strong "
   "analysts are active idea generators and form the bench from which Senior Analysts and PMs are developed.")]

story+=[Paragraph("2.3&nbsp; Leadership, control and support",SH),
 bullet("<b>Chief Executive Officer (CEO)</b> &mdash; leads the firm as a business: strategy, commercial development, "
        "governance and overall management, so the investment team can focus on investing."),
 bullet("<b>Chief Financial Officer (CFO)</b> &mdash; owns the financial management of the firm and the funds: fund "
        "accounting, treasury, financial control, audit and tax."),
 bullet("<b>Chief Operating Officer (COO)</b> &mdash; responsible for the operating platform: operations, "
        "infrastructure, technology, vendors and the day-to-day running of the business."),
 bullet("<b>Execution</b> &mdash; implements the CIO's accumulation and divestment decisions efficiently and "
        "discreetly, given our concentrated ownership positions."),
 bullet("<b>Risk</b> &mdash; independent monitoring and management of portfolio and firm-wide risk, including "
        "exposure, liquidity and concentration."),
 bullet("<b>Compliance</b> &mdash; ensures the firm meets its regulatory obligations and upholds the highest "
        "standards of conduct."),
 bullet("<b>Investor Relations (IR)</b> &mdash; owns the investor relationship: reporting, communication and capital "
        "raising."),
 bullet("<b>Support</b> &mdash; provides the operational and administrative backbone that keeps the firm running.")]

# 3. pool created
story+=[Paragraph("3.&nbsp; How the performance pool is created",HD),
 B("When the fund delivers gains for our investors, it earns a performance fee. A fixed share of that performance "
   "fee &mdash; the <b>team take of 60%</b> &mdash; is set aside to form the <b>performance pool</b> for staff. The "
   "remaining 40% is retained by the firm. The pool is therefore directly linked to the returns we generate: a "
   "stronger year for the fund means a larger pool, and as assets under management (AUM) grow, the pool grows with "
   "them."),
 B("The performance fee earned on the firm's <b>S class of shares</b> is excluded from the pool: performance units "
   "do not receive any share of the performance fee generated by the S class of shares.")]

# 4. shared
story+=[Paragraph("4.&nbsp; How the pool is shared &mdash; performance weights and units",HD),
 B("Each role carries a <b>performance weight</b> that reflects its seniority and its contribution to investment "
   "performance. Weights convert the pool into individual awards through a simple, fully transparent calculation:"),
 bullet("Each person's weight represents a number of <b>performance units</b>."),
 bullet("We add up the weights of everyone in the firm to get the <b>total units</b>."),
 bullet("The pool is divided by the total units to give the <b>value of one unit</b>."),
 bullet("Your award = <b>your weight &times; the value of one unit</b>.")]
box=Table([[Paragraph("Value of one unit&nbsp; =&nbsp; Performance pool&nbsp; &divide;&nbsp; Total performance units",BOX)],
           [Paragraph("Your performance award&nbsp; =&nbsp; Your weight&nbsp; &times;&nbsp; Value of one unit",BOX)]],
          colWidths=[16.4*cm])
box.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#EAF0FA")),
    ("BOX",(0,0),(-1,-1),0.8,BLUE),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
    ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6)]))
story+=[Spacer(1,2),box,Spacer(1,6)]

# 5. roles & weights table
story+=[Paragraph("5.&nbsp; Roles and performance weights",HD),
 B("The table below sets out the roles in the firm, grouped by track, and the performance weight attached to each. "
   "Weights are reviewed periodically and apply per person in the role.")]
track_fill={"Investment":"#DDE9F7","Leadership":"#E7EFDC","Control":"#FBE9D8","Client &amp; Support":"#F2F2F2"}
rows=[
 ("Investment","Chief Investment Officer (CIO)","Selection, chairs investment meetings, owns portfolio construction.","16"),
 ("Investment","Portfolio Manager (PM)","Leads investments; drives value creation in portfolio companies.","8"),
 ("Investment","Junior PM","Co-leads investments; owns workstreams of the engagement agenda.","6"),
 ("Investment","Head of Research","Leads research function; partners with CIO on portfolio construction.","—"),
 ("Investment","Senior Analyst","Original idea generation, leads due diligence, mentors analysts.","4"),
 ("Investment","Analyst","Confirmatory due diligence and idea sourcing within deal teams.","2.5"),
 ("Leadership","Chief Executive Officer (CEO)","Firm leadership and overall business management.","8"),
 ("Leadership","Chief Financial Officer (CFO)","Finance, fund accounting and treasury.","3"),
 ("Leadership","Chief Operating Officer (COO)","Operations and firm infrastructure.","3"),
 ("Control","Execution","Trade execution of accumulation and divestment.","2"),
 ("Control","Risk","Risk management and portfolio monitoring.","2"),
 ("Control","Compliance","Regulatory compliance.","1"),
 ("Client &amp; Support","Investor Relations (IR)","Investor relations and capital raising.","1"),
 ("Client &amp; Support","Support","Operational and administrative support.","1"),
]
data=[[Paragraph("Track",CELLH),Paragraph("Role",CELLH),Paragraph("Primary responsibility",CELLH),Paragraph("Weight",CELLH)]]
for tr,role,desc,wt in rows:
    data.append([Paragraph(tr,CELLT),Paragraph(role,CELLB),Paragraph(desc,CELL),Paragraph(wt,CELLW)])
tbl=Table(data,colWidths=[2.5*cm,4.7*cm,7.2*cm,2.0*cm],repeatRows=1)
tstyle=[("BACKGROUND",(0,0),(-1,0),NAVY),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ("ALIGN",(3,0),(3,-1),"CENTER"),("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#BFBFBF")),
    ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
    ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5)]
for i,(tr,*_ ) in enumerate(rows,start=1):
    tstyle.append(("BACKGROUND",(0,i),(0,i),colors.HexColor(track_fill[tr])))
tbl.setStyle(TableStyle(tstyle))
story+=[tbl,Spacer(1,2),
 Paragraph("&ldquo;&mdash;&rdquo;&nbsp; The Head of Research is currently held on a voluntary basis and does not "
           "participate in the performance pool.",FOOT),
 Spacer(1,6)]

# 6. growth
story+=[Paragraph("6.&nbsp; How awards behave as the firm grows",HD),
 B("Two things happen as the firm adds people and assets, and it is important to understand both:"),
 bullet("Your <b>percentage share</b> of the pool may gradually reduce as the team grows and total units increase."),
 bullet("Your <b>award in absolute terms is expected to rise</b>, because the pool grows with AUM and fund performance."),
 B("In other words, a smaller slice of a much larger pie. This keeps the framework fair as we scale: building a "
   "stronger, larger team is what drives the bigger pool that benefits everyone.")]

# 7. progression
story+=[Paragraph("7.&nbsp; Career progression",HD),
 B("Weights increase with responsibility. The investment track in particular offers a clear ladder &mdash; "
   "<b>Analyst &rarr; Senior Analyst &rarr; Junior PM &rarr; Portfolio Manager &rarr; CIO</b> &mdash; with each step "
   "carrying a meaningfully higher weight. Progression is how your share of the pool grows alongside the firm.")]

# 8. units and leaving
story+=[Paragraph("8.&nbsp; Performance units and leaving the firm",HD),
 B("Performance units are the firm's <b>performance-related bonus mechanism</b>: they are funded entirely from the "
   "firm's performance fees and represent a share of that bonus pool. They are <b>not</b> equity, ownership, or a "
   "vested or transferable asset."),
 B("Performance units are valid only for as long as you are working at the firm. When you leave &mdash; for any "
   "reason, and regardless of whether you are treated as a <b>good leaver or a bad leaver</b> &mdash; your "
   "performance units are <b>relinquished in full</b> and lapse with immediate effect, and no further award is "
   "payable in respect of them. Units cannot be transferred or assigned.")]

# 9. governance
story+=[Paragraph("9.&nbsp; Governance",HD),
 B("The framework, the list of roles and the weights are reviewed periodically by the firm's leadership. Awards are "
   "determined by this framework; the firm retains reasonable discretion in exceptional circumstances and in line "
   "with regulatory requirements. This document is provided for information and does not form part of any "
   "individual's contract of employment."),
 Spacer(1,4),
 HRFlowable(width="100%",thickness=0.6,color=colors.HexColor("#BFBFBF"),spaceBefore=2,spaceAfter=4),
 Paragraph("Athanase Industrial Partner&nbsp; &middot;&nbsp; Roles, Responsibilities &amp; Performance Compensation "
           "Framework&nbsp; &middot;&nbsp; Internal &mdash; for staff information",FOOT)]

doc=SimpleDocTemplate(OUT,pagesize=A4,topMargin=1.6*cm,bottomMargin=1.4*cm,leftMargin=2.0*cm,rightMargin=2.0*cm,
                      title="Roles, Responsibilities & Performance Compensation Framework",
                      author="Athanase Industrial Partner")
doc.build(story)
print("Saved",OUT)
