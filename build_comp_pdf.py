"""Staff-facing Performance Compensation Framework -> PDF (reportlab)."""
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
BODY=ParagraphStyle("BODY",parent=ss["Normal"],fontName="Helvetica",fontSize=10.3,leading=13.4,spaceAfter=5,alignment=TA_LEFT)
BUL=ParagraphStyle("BUL",parent=BODY,leftIndent=14,bulletIndent=4,spaceAfter=2)
BOX=ParagraphStyle("BOX",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=10.5,textColor=NAVY,alignment=TA_CENTER,leading=15)
CELL=ParagraphStyle("CELL",parent=ss["Normal"],fontName="Helvetica",fontSize=8.6,leading=10.4)
CELLB=ParagraphStyle("CELLB",parent=CELL,fontName="Helvetica-Bold")
CELLH=ParagraphStyle("CELLH",parent=CELL,fontName="Helvetica-Bold",textColor=colors.white)
CELLT=ParagraphStyle("CELLT",parent=CELL,fontName="Helvetica-Bold",textColor=NAVY)
CELLW=ParagraphStyle("CELLW",parent=ss["Normal"],fontName="Helvetica-Bold",fontSize=10,textColor=BLUE,alignment=TA_CENTER)
FOOT=ParagraphStyle("FOOT",parent=ss["Normal"],fontName="Helvetica-Oblique",fontSize=8,textColor=GREY)

def B(t): return Paragraph(t,BODY)
def bullet(t): return Paragraph(("&bull;&nbsp;&nbsp;"+t),BUL)

story=[]
story+=[Paragraph("Athanase Industrial Partner",H1),
        Paragraph("Performance Compensation Framework",SUB),
        Paragraph("Staff guide to how the performance pool is created and shared",MUT),
        HRFlowable(width="100%",thickness=1.4,color=NAVY,spaceBefore=4,spaceAfter=7)]

story+=[Paragraph("1.&nbsp; Purpose",HD),
 B("This document explains how performance-based compensation works at Athanase Industrial Partner. "
   "Our philosophy is simple: when the fund performs, the people who create that performance share directly in it. "
   "The framework is designed to be transparent, predictable, and to reward both contribution and progression &mdash; "
   "so that everyone understands how their reward is determined and how it grows over time.")]

story+=[Paragraph("2.&nbsp; How the performance pool is created",HD),
 B("When the fund delivers gains for our investors, it earns a performance fee. A fixed share of that performance "
   "fee &mdash; the <b>team take of 60%</b> &mdash; is set aside to form the <b>performance pool</b> for staff. The "
   "remaining 40% is retained by the firm. The pool is therefore directly linked to the returns we generate: a "
   "stronger year for the fund means a larger pool for the team, and as assets under management (AUM) grow, the "
   "pool grows with them.")]

story+=[Paragraph("3.&nbsp; How the pool is shared &mdash; performance weights and units",HD),
 B("Each role at the firm carries a <b>performance weight</b>. Your weight reflects the role's seniority and its "
   "contribution to investment performance. Weights convert the pool into individual awards through a simple, "
   "fully transparent calculation:"),
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

story+=[Paragraph("4.&nbsp; Roles and performance weights",HD),
 B("The table below sets out the roles in the firm, grouped by track, and the performance weight attached to each. "
   "Weights are reviewed periodically and apply per person in the role.")]

track_fill={"Investment":"#DDE9F7","Leadership":"#E7EFDC","Control":"#FBE9D8","Client &amp; Support":"#F2F2F2"}
rows=[
 ("Investment","Chief Investment Officer (CIO)","Sets investment strategy; ultimate responsibility for fund performance.","16"),
 ("Investment","Portfolio Manager (PM)","Manages capital and runs an investment book.","8"),
 ("Investment","Junior PM","Runs a sleeve or co-manages a book; developing toward full PM.","6"),
 ("Investment","Senior Analyst","Lead analyst with deep sector coverage; mentors analysts.","4"),
 ("Investment","Analyst","Fundamental research and idea generation.","2.5"),
 ("Leadership","Chief Executive Officer (CEO)","Firm leadership and overall business management.","8"),
 ("Leadership","Chief Financial Officer (CFO)","Finance, fund accounting and treasury.","3"),
 ("Leadership","Chief Operating Officer (COO)","Operations and firm infrastructure.","3"),
 ("Control","Execution","Trade execution.","2"),
 ("Control","Risk","Risk management and portfolio monitoring.","2"),
 ("Control","Compliance","Regulatory compliance.","1"),
 ("Client &amp; Support","Investor Relations (IR)","Investor relations and capital raising.","1"),
 ("Client &amp; Support","Support","Operational and administrative support.","1"),
]
data=[[Paragraph("Track",CELLH),Paragraph("Role",CELLH),Paragraph("What the role does",CELLH),Paragraph("Weight",CELLH)]]
for tr,role,desc,wt in rows:
    data.append([Paragraph(tr,CELLT),Paragraph(role,CELLB),Paragraph(desc,CELL),Paragraph(wt,CELLW)])
tbl=Table(data,colWidths=[2.5*cm,4.3*cm,7.6*cm,2.0*cm],repeatRows=1)
tstyle=[("BACKGROUND",(0,0),(-1,0),NAVY),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ("ALIGN",(3,0),(3,-1),"CENTER"),("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#BFBFBF")),
    ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
    ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5)]
for i,(tr,*_ ) in enumerate(rows,start=1):
    tstyle.append(("BACKGROUND",(0,i),(0,i),colors.HexColor(track_fill[tr])))
tbl.setStyle(TableStyle(tstyle))
story+=[tbl,Spacer(1,7)]

story+=[Paragraph("5.&nbsp; How awards behave as the firm grows",HD),
 B("Two things happen as the firm adds people and assets, and it is important to understand both:"),
 bullet("Your <b>percentage share</b> of the pool may gradually reduce as the team grows and total units increase."),
 bullet("Your <b>award in absolute terms is expected to rise</b>, because the pool grows with AUM and fund performance."),
 B("In other words, a smaller slice of a much larger pie. This keeps the framework fair as we scale: building a "
   "stronger, larger team is what drives the bigger pool that benefits everyone.")]

story+=[Paragraph("6.&nbsp; Career progression",HD),
 B("Weights increase with responsibility. The investment track in particular offers a clear ladder &mdash; "
   "<b>Analyst &rarr; Senior Analyst &rarr; Junior PM &rarr; Portfolio Manager &rarr; CIO</b> &mdash; with each step "
   "carrying a meaningfully higher weight. Progression is how your share of the pool grows alongside the firm.")]

story+=[Paragraph("7.&nbsp; Governance",HD),
 B("The framework, the list of roles and the weights are reviewed periodically by the firm's leadership. Awards are "
   "determined by this framework; the firm retains reasonable discretion in exceptional circumstances and in line "
   "with regulatory requirements. This document is provided for information and does not form part of any "
   "individual's contract of employment."),
 Spacer(1,4),
 HRFlowable(width="100%",thickness=0.6,color=colors.HexColor("#BFBFBF"),spaceBefore=2,spaceAfter=4),
 Paragraph("Athanase Industrial Partner&nbsp; &middot;&nbsp; Performance Compensation Framework&nbsp; &middot;&nbsp; "
           "Internal &mdash; for staff information",FOOT)]

doc=SimpleDocTemplate(OUT,pagesize=A4,topMargin=1.6*cm,bottomMargin=1.4*cm,leftMargin=2.0*cm,rightMargin=2.0*cm,
                      title="Performance Compensation Framework",author="Athanase Industrial Partner")
doc.build(story)
print("Saved",OUT)
