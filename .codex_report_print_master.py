from pathlib import Path
import re, subprocess, shutil

from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image
from reportlab.pdfgen import canvas
from docx import Document
from docx.shared import Inches

ROOT=Path(__file__).resolve().parent
REPORT=ROOT/"Report"
BUILD=ROOT/".codex_full_report_build"
SRC=BUILD/"source_full_rebuild.pdf"
CUR=REPORT/"EcoSort_AI_Project_Report_Professional_Full.pdf"
FRONT=BUILD/"verified_frontmatter.pdf"
SUPP=BUILD/"supplementary_figures.pdf"
OUT=REPORT/"EcoSort_AI_Project_Report_Professional_Complete.pdf"
DOCX=REPORT/"EcoSort_AI_Project_Report_Professional_Complete.docx"
RENDER=BUILD/"final_pages"
RENDER.mkdir(exist_ok=True)
for old_page in RENDER.glob("page-*.jpg"):
    old_page.unlink()

source=PdfReader(str(SRC))
source_text=[p.extract_text() or "" for p in source.pages]
chapter_rows=[]
fig_rows=[]
for i,text in enumerate(source_text):
    m=re.search(r"(?m)^\s*(\d{2})\s*\|\s*([^\n]+)",text)
    if m:
        n=int(m.group(1)); name=m.group(2).strip()
        if 1<=n<=38:
            chapter_rows.append((n,name,str(n+7)))
chapter_rows.append((39,"Appendices","47"))
for i,text in enumerate(source_text):
    for m in re.finditer(r"(?m)^\s*(Figure|FlowDiagram|DFD)\s*([0-9.]+)[.: -]*([^\n]*)",text,re.I):
        caption=(m.group(3) or "").strip()
        fig_rows.append((f"{m.group(1)} {m.group(2)}",caption[:86],str(i+1)))
fig_rows += [("Figure A","Level-0 context data-flow diagram","50"),("Figure B","Level-1 detailed data-flow diagram","51"),("Figure C","Tools and technologies visual overview","52")]

styles=getSampleStyleSheet()
styles.add(ParagraphStyle(name="CoverUni",fontName="Helvetica-Bold",fontSize=12,leading=17,alignment=TA_CENTER,textColor=colors.HexColor("#164A36"),spaceAfter=4))
styles.add(ParagraphStyle(name="CoverProject",fontName="Helvetica-Bold",fontSize=26,leading=32,alignment=TA_CENTER,textColor=colors.HexColor("#103E2E"),spaceAfter=8))
styles.add(ParagraphStyle(name="FrontHead",fontName="Helvetica-Bold",fontSize=18,leading=22,textColor=colors.HexColor("#103E2E"),spaceAfter=10))
styles.add(ParagraphStyle(name="FrontBody",fontName="Times-Roman",fontSize=10.5,leading=15,spaceAfter=8))
styles.add(ParagraphStyle(name="FrontCell",fontName="Helvetica",fontSize=7.6,leading=9,spaceAfter=0))
styles.add(ParagraphStyle(name="Tiny",fontName="Helvetica",fontSize=7.3,leading=8.8,spaceAfter=0))

def para(text,style="FrontBody"):
    return Paragraph(text.replace("&","&amp;"),styles[style])
def table(rows,widths,fontstyle="FrontCell"):
    dat=[[Paragraph(str(c).replace("&","&amp;"),styles[fontstyle]) for c in r] for r in rows]
    t=Table(dat,colWidths=widths,repeatRows=1,hAlign="LEFT")
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#176547")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.35,colors.HexColor("#D7E2DA")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F2F7F3")]),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
    return t
def front_footer(c,d):
    c.saveState()
    if d.page>1:
        c.setFont("Helvetica",8);c.setFillColor(colors.HexColor("#63756A"));c.drawRightString(A4[0]-18*mm,11*mm,str(d.page))
        c.setStrokeColor(colors.HexColor("#D7E2DA"));c.line(18*mm,15*mm,A4[0]-18*mm,15*mm)
    c.restoreState()

story=[]
logo=ROOT/"static"/"garbage-ai-logo.jpg"
story += [Spacer(1,10*mm),Image(str(logo),width=32*mm,height=32*mm),Spacer(1,5*mm),para("IK GUJRAL PUNJAB TECHNICAL UNIVERSITY, HOSHIARPUR CAMPUS","CoverUni"),para("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING","CoverUni"),Spacer(1,12*mm),para("PROJECT REPORT","CoverUni"),para("EcoSort AI","CoverProject"),para("AI-Powered Waste Classification","CoverUni"),para("Garbage Classification with Deep Learning","FrontBody"),Spacer(1,9*mm),para("Submitted in partial fulfilment of the requirements for the award of the degree of Bachelor of Technology in Computer Science Engineering.","FrontBody"),Spacer(1,4*mm)]
story += [table([["SUBMITTED BY","SUBMITTED TO"],["Arshpreet Singh\nSarabjit Singh\nPooja\nUniversity Roll Nos.: ____________________","Dr. Brijesh Bakariya\nProject Mentor and Head, CSE\nEr. Satyam Sharma\nProject Co-Mentor"]],[78*mm,78*mm]),Spacer(1,7*mm),para("Semester: 7th\nAcademic session: ____________________\nSubmission date: ____________________","FrontBody"),para("Hard-bound project report","FrontBody"),PageBreak()]
story += [para("DECLARATION","FrontHead"),para("We, the undersigned project team, declare that this report describes the EcoSort AI work carried out by our team. Project implementation details are based on the supplied repository, and reported model measures are transcribed from the saved model metadata. External technical references are listed at the end of the report. The declaration must be reviewed and signed by each student before submission."),table([["Student name","Signature","Date"],["Arshpreet Singh","____________________________","____________"],["Sarabjit Singh","____________________________","____________"],["Pooja","____________________________","____________"]],[46*mm,75*mm,32*mm]),Spacer(1,6*mm),para("Department: Computer Science and Engineering\nInstitution: IK Gujral Punjab Technical University, Hoshiarpur Campus\nProgramme: Bachelor of Technology in Computer Science Engineering\nSemester: 7th","FrontBody"),PageBreak()]
story += [para("CERTIFICATE","FrontHead"),para("This is to certify that the project entitled EcoSort AI: AI-Powered Waste Classification has been carried out by Arshpreet Singh, Sarabjit Singh, and Pooja in the Department of Computer Science and Engineering, IK Gujral Punjab Technical University, Hoshiarpur Campus, in partial fulfilment of the requirements for the Bachelor of Technology in Computer Science Engineering. The project was completed under the guidance of Dr. Brijesh Bakariya, Project Mentor and Head of Department, and Er. Satyam Sharma, Project Co-Mentor. This certificate is subject to institutional review and authorized signatures."),Spacer(1,18*mm),table([["Project Mentor / HOD","Project Co-Mentor","External Examiner"],["Dr. Brijesh Bakariya","Er. Satyam Sharma","Name: __________________"],["Signature: ____________","Signature: ____________","Signature: ____________"],["Date: _________________","Date: _________________","Date: _________________"]],[51*mm,51*mm,48*mm]),Spacer(1,12*mm),para("Institution seal / remarks\n\n________________________________________________________________________________\n\n________________________________________________________________________________"),PageBreak()]
story += [para("ABSTRACT","FrontHead"),para("EcoSort AI classifies a waste-item image into one of twelve dataset categories and presents the result through a responsive web application. The system combines a MobileNetV2 transfer-learning model, a Flask inference service, and browser workflows for photo upload and camera capture. Camera users see a live preview, capture and review a still, and explicitly request prediction."),para("The dataset contains 15,515 images. Model input is 224 by 224 pixels. Training uses an 80/20 training-validation split, class weighting, augmentation, early stopping, and fine-tuning. Saved metadata reports 94.68 percent validation accuracy and 0.172 validation loss. These values describe the saved validation split, not an independent test set or a guarantee for new user photographs."),para("Keywords: waste classification, computer vision, MobileNetV2, TensorFlow, Flask, image recognition."),PageBreak()]
story += [para("ACKNOWLEDGEMENTS","FrontHead"),para("The project team expresses sincere gratitude to Dr. Brijesh Bakariya, Project Mentor and Head of the Department of Computer Science and Engineering, for his academic guidance. The team also thanks Er. Satyam Sharma, Project Co-Mentor, for his guidance and feedback. We acknowledge the University and Department for the academic environment in which this project was developed, and the authors of the open-source frameworks cited in the report. Dataset source and licence details should be verified before redistribution."),Spacer(1,8*mm),para("Project team\nArshpreet Singh\nSarabjit Singh\nPooja"),PageBreak()]

story += [para("TABLE OF CONTENTS","FrontHead"),table([["CHAPTER NO.","TOPIC","PAGE NO."]]+[[f"CHAPTER {n}",name,page] for n,name,page in chapter_rows]+[["REFERENCES","Technical references","46"],["APPENDICES","Repository and submission material","47-49"],["SUPPLEMENTS","Additional project diagrams and technology overview","50-52"]],[29*mm,108*mm,18*mm],"Tiny"),para("LIST OF FIGURES","FrontHead")]
story += [table([["FIGURE NO.","TOPIC","PAGE NO."]]+[list(r) for r in fig_rows],[32*mm,105*mm,18*mm],"Tiny")]
SimpleDocTemplate(str(FRONT),pagesize=A4,leftMargin=18*mm,rightMargin=18*mm,topMargin=17*mm,bottomMargin=18*mm,title="Project Report Front Matter").build(story,onFirstPage=front_footer,onLaterPages=front_footer)

# Three polished supplementary pages hold the app-specific DFDs and technology graphic.
supp_story=[]
supp_story += [para("SUPPLEMENTARY FIGURE A","FrontHead"),para("Level-0 context data-flow diagram. The user submits an image and receives ranked class suggestions."),Image(str(BUILD/"dfd_level0.png"),width=170*mm,height=94*mm),Spacer(1,8*mm),para("The system boundary includes image validation, preprocessing, model inference, and result mapping. The diagram is a project-specific DFD created from the implemented request flow.") ,PageBreak()]
supp_story += [para("SUPPLEMENTARY FIGURE B","FrontHead"),para("Level-1 data-flow diagram. Internal processes and saved artifacts are shown separately."),Image(str(BUILD/"dfd_level1.png"),width=170*mm,height=94*mm),Spacer(1,8*mm),para("Invalid image data or unavailable model selection returns an error. A valid image is converted to RGB, resized, classified, and mapped back to the configured class names."),PageBreak()]
supp_story += [para("SUPPLEMENTARY FIGURE C","FrontHead"),para("Tools and technologies used in EcoSort AI."),Image(str(BUILD/"technology_stack.png"),width=170*mm,height=90*mm),Spacer(1,8*mm),para("This visual summarizes the runtime, model framework, classifier backbone, server, image utilities, browser interface, and deployment components documented in the project repository.")]
SimpleDocTemplate(str(SUPP),pagesize=A4,leftMargin=18*mm,rightMargin=18*mm,topMargin=18*mm,bottomMargin=18*mm,title="Supplementary Figures").build(supp_story)

# Combine corrected front matter, the full project-authored body, and three labeled supplemental pages.
front_reader=PdfReader(str(FRONT)); supp_reader=PdfReader(str(SUPP)); writer=PdfWriter()
for page in front_reader.pages: writer.add_page(page)
for idx in range(7,len(source.pages)): writer.add_page(source.pages[idx])
for page in supp_reader.pages: writer.add_page(page)

# Replace the source report's footer values with final physical numbering where printed.
from reportlab.pdfgen.canvas import Canvas
from io import BytesIO
for i in range(7,len(front_reader.pages)+len(source.pages)-7):
    buf=BytesIO(); c=Canvas(buf,pagesize=A4); c.setFillColor(colors.white); c.rect(A4[0]-46*mm,7*mm,34*mm,12*mm,fill=1,stroke=0); c.setFillColor(colors.HexColor("#5B6B62")); c.setFont("Helvetica",8); c.drawRightString(A4[0]-18*mm,12*mm,f"{i+1}  |  52"); c.save(); buf.seek(0)
    overlay=PdfReader(buf).pages[0]; writer.pages[i].merge_page(overlay)
for j in range(3):
    page=writer.pages[len(front_reader.pages)+len(source.pages)-7+j]
    buf=BytesIO(); c=Canvas(buf,pagesize=A4); c.saveState(); c.setFont("Helvetica",8); c.setFillColor(colors.HexColor("#5B6B62")); c.drawRightString(A4[0]-18*mm,10*mm,f"Supplement {j+1} of 3"); c.setStrokeColor(colors.HexColor("#D7E2DA")); c.line(18*mm,15*mm,A4[0]-18*mm,15*mm); c.restoreState(); c.save(); buf.seek(0); page.merge_page(PdfReader(buf).pages[0])
with open(OUT,"wb") as f: writer.write(f)
print(f"Created {OUT} ({len(writer.pages)} pages)")

# Create a page-faithful Word print copy, one rendered page per A4 page.
pdftoppm=Path(r"C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin\pdftoppm.exe")
subprocess.run([str(pdftoppm),"-jpeg","-r","110","-jpegopt","quality=82",str(OUT),str(RENDER/"page")],check=True,capture_output=True)
images=sorted(RENDER.glob("page-*.jpg"),key=lambda p:int(re.search(r"-(\d+)\.jpg$",p.name).group(1)))
doc=Document(); s=doc.sections[0]; s.page_width=Inches(8.27);s.page_height=Inches(11.69);s.top_margin=Inches(0);s.bottom_margin=Inches(0);s.left_margin=Inches(0);s.right_margin=Inches(0)
style=doc.styles["Normal"];style.paragraph_format.space_after=0;style.paragraph_format.space_before=0;style.paragraph_format.line_spacing=1
for i,img in enumerate(images):
    p=doc.add_paragraph();p.paragraph_format.space_after=0;p.paragraph_format.space_before=0;p.paragraph_format.keep_together=True;p.add_run().add_picture(str(img),width=Inches(8.27),height=Inches(11.69))
    if i<len(images)-1: doc.add_page_break()
doc.core_properties.title="EcoSort AI Professional Complete Project Report"
doc.core_properties.subject="Page-faithful print copy of the complete 52-page project report"
doc.save(DOCX)
print(f"Created {DOCX} ({len(images)} page images)")
