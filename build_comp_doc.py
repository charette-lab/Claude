"""Generate the staff-facing Performance Compensation Framework document (.docx)."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY=RGBColor(0x1F,0x38,0x64); BLUE=RGBColor(0x2E,0x54,0x96); GREY=RGBColor(0x59,0x59,0x59)
OUT="/home/user/Claude/Performance_Compensation_Framework.docx"

doc=Document()
# page + base style
sec=doc.sections[0]
sec.top_margin=Cm(1.8); sec.bottom_margin=Cm(1.6); sec.left_margin=Cm(2.0); sec.right_margin=Cm(2.0)
st=doc.styles["Normal"]; st.font.name="Calibri"; st.font.size=Pt(10.5)
st.paragraph_format.space_after=Pt(6); st.paragraph_format.line_spacing=1.06

def shade(cell,hexcolor):
    tcPr=cell._tc.get_or_add_tcPr(); sh=OxmlElement("w:shd")
    sh.set(qn("w:val"),"clear"); sh.set(qn("w:fill"),hexcolor); tcPr.append(sh)

def setcellfont(cell,size=9.5,bold=False,color=None,align=None):
    for p in cell.paragraphs:
        if align is not None: p.alignment=align
        p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
        for r in p.runs:
            r.font.size=Pt(size); r.font.bold=bold
            if color is not None: r.font.color.rgb=color

def heading(text,size=12.5,color=NAVY,space_before=8):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(space_before); p.paragraph_format.space_after=Pt(3)
    r=p.add_run(text); r.font.bold=True; r.font.size=Pt(size); r.font.color.rgb=color
    return p

def body(text,size=10.5,after=6):
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(after)
    r=p.add_run(text); r.font.size=Pt(size)
    return p

def bullet(text,size=10.5):
    p=doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_after=Pt(2)
    r=p.add_run(text); r.font.size=Pt(size); return p

# ---------- title ----------
t=doc.add_paragraph(); t.alignment=WD_ALIGN_PARAGRAPH.LEFT
r=t.add_run("Athanase Industrial Partner"); r.font.bold=True; r.font.size=Pt(17); r.font.color.rgb=NAVY
t.paragraph_format.space_after=Pt(0)
s=doc.add_paragraph(); r=s.add_run("Performance Compensation Framework"); r.font.size=Pt(13); r.font.color.rgb=BLUE
s.paragraph_format.space_after=Pt(2)
m=doc.add_paragraph(); r=m.add_run("Staff guide to how the performance pool is created and shared"); r.italic=True; r.font.size=Pt(10); r.font.color.rgb=GREY
m.paragraph_format.space_after=Pt(2)
# divider
pb=doc.add_paragraph(); pPr=pb._p.get_or_add_pPr(); bdr=OxmlElement("w:pBdr"); bot=OxmlElement("w:bottom")
bot.set(qn("w:val"),"single"); bot.set(qn("w:sz"),"12"); bot.set(qn("w:space"),"1"); bot.set(qn("w:color"),"1F3864")
bdr.append(bot); pPr.append(bdr); pb.paragraph_format.space_after=Pt(6)

# ---------- 1. purpose ----------
heading("1.  Purpose",space_before=2)
body("This document explains how performance-based compensation works at Athanase Industrial Partner. "
     "Our philosophy is simple: when the fund performs, the people who create that performance share directly in it. "
     "The framework is designed to be transparent, predictable, and to reward both contribution and progression — "
     "so that everyone understands how their reward is determined and how it grows over time.")

# ---------- 2. how the pool is created ----------
heading("2.  How the performance pool is created")
body("When the fund delivers gains for our investors, it earns a performance fee. A fixed share of that performance "
     "fee — the team take of 60% — is set aside to form the performance pool for staff. The remaining 40% is "
     "retained by the firm. The pool is therefore directly linked to the returns we generate: a stronger year for the "
     "fund means a larger pool for the team, and as assets under management (AUM) grow, the pool grows with them.")

# ---------- 3. weights & units ----------
heading("3.  How the pool is shared — performance weights and units")
body("Each role at the firm carries a performance weight. Your weight reflects the role's seniority and its "
     "contribution to investment performance. Weights convert the pool into individual awards through a simple, "
     "fully transparent calculation:")
bullet("Each person's weight represents a number of performance units.")
bullet("We add up the weights of everyone in the firm to get the total number of units.")
bullet("The pool is divided by the total units to give the value of one unit.")
bullet("Your award = your weight × the value of one unit.")

# formula box (single-cell shaded table)
fb=doc.add_table(rows=1,cols=1); fb.alignment=WD_TABLE_ALIGNMENT.CENTER
cell=fb.cell(0,0); shade(cell,"EAF0FA")
cell.width=Inches(6.2)
p=cell.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run("Value of one unit  =  Performance pool  ÷  Total performance units")
r.font.bold=True; r.font.size=Pt(10.5); r.font.color.rgb=NAVY
p2=cell.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p2.add_run("Your performance award  =  Your weight  ×  Value of one unit")
r.font.bold=True; r.font.size=Pt(10.5); r.font.color.rgb=NAVY
for bd in ("top","bottom","left","right"):
    e=OxmlElement(f"w:{bd}"); e.set(qn("w:val"),"single"); e.set(qn("w:sz"),"6"); e.set(qn("w:color"),"2E5496")
    cell._tc.get_or_add_tcPr().append(e)
doc.add_paragraph().paragraph_format.space_after=Pt(2)

# ---------- 4. roles & weights table ----------
heading("4.  Roles and performance weights")
body("The table below sets out the roles in the firm, grouped by track, and the performance weight attached to each. "
     "Weights are reviewed periodically and apply per person in the role.",after=4)

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
 ("Client & Support","Investor Relations (IR)","Investor relations and capital raising.","1"),
 ("Client & Support","Support","Operational and administrative support.","1"),
]
tbl=doc.add_table(rows=1,cols=4); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; tbl.style="Table Grid"
hdr=tbl.rows[0].cells
for i,(txt,w) in enumerate([("Track",1.1),("Role",1.9),("What the role does",3.0),("Weight",0.7)]):
    hdr[i].text=txt; shade(hdr[i],"1F3864")
    setcellfont(hdr[i],size=9.5,bold=True,color=RGBColor(0xFF,0xFF,0xFF),
                align=WD_ALIGN_PARAGRAPH.CENTER if i==3 else WD_ALIGN_PARAGRAPH.LEFT)
track_fill={"Investment":"DDE9F7","Leadership":"E7EFDC","Control":"FBE9D8","Client & Support":"F2F2F2"}
for tr,role,desc,wt in rows:
    c=tbl.add_row().cells
    c[0].text=tr; c[1].text=role; c[2].text=desc; c[3].text=wt
    shade(c[0],track_fill[tr])
    setcellfont(c[0],size=9,bold=True,color=NAVY)
    setcellfont(c[1],size=9,bold=True)
    setcellfont(c[2],size=9)
    setcellfont(c[3],size=10,bold=True,color=BLUE,align=WD_ALIGN_PARAGRAPH.CENTER)
# column widths
for row in tbl.rows:
    row.cells[0].width=Inches(1.1); row.cells[1].width=Inches(1.9); row.cells[2].width=Inches(3.0); row.cells[3].width=Inches(0.7)

# ---------- 5. as the firm grows ----------
heading("5.  How awards behave as the firm grows")
body("Two things happen as the firm adds people and assets, and it is important to understand both:")
bullet("Your percentage share of the pool may gradually reduce as the team grows and total units increase.")
bullet("Your award in absolute terms is expected to rise, because the pool grows with AUM and fund performance.")
body("In other words, a smaller slice of a much larger pie. This keeps the framework fair as we scale: building a "
     "stronger, larger team is what drives the bigger pool that benefits everyone.",after=4)

# ---------- 6. progression ----------
heading("6.  Career progression")
body("Weights increase with responsibility. The investment track in particular offers a clear ladder — Analyst → "
     "Senior Analyst → Junior PM → Portfolio Manager → CIO — with each step carrying a meaningfully higher weight. "
     "Progression is how your share of the pool grows alongside the firm.")

# ---------- 7. governance ----------
heading("7.  Governance")
body("The framework, the list of roles and the weights are reviewed periodically by the firm's leadership. Awards are "
     "determined by this framework; the firm retains reasonable discretion in exceptional circumstances and in line "
     "with regulatory requirements. This document is provided for information and does not form part of any "
     "individual's contract of employment.",after=2)

# footer
foot=doc.add_paragraph(); foot.alignment=WD_ALIGN_PARAGRAPH.LEFT
r=foot.add_run("Athanase Industrial Partner  ·  Performance Compensation Framework  ·  Internal — for staff information")
r.italic=True; r.font.size=Pt(8); r.font.color.rgb=GREY

doc.save(OUT)
print("Saved",OUT)
