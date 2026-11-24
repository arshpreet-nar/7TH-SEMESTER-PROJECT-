from pathlib import Path
import json, re, runpy, subprocess, shutil
from io import BytesIO

from pypdf import PdfReader, PdfWriter, Transformation, PageObject
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image, KeepTogether
from reportlab.pdfgen.canvas import Canvas
from docx import Document
from docx.shared import Inches
from PIL import Image as PILImage, ImageDraw, ImageFont

ROOT=Path(__file__).resolve().parent
REPORT=ROOT/"Report"
BUILD=ROOT/".codex_format_alignment"
SOURCE_BUILD=BUILD/"source_build"
SOURCE_BUILD.mkdir(exist_ok=True)
REPORT.mkdir(exist_ok=True)
SOURCE_PDF=SOURCE_BUILD/"eco_sort_letter_source.pdf"
FRONT_PDF=SOURCE_BUILD/"front_matter_letter.pdf"
SUPP_PDF=SOURCE_BUILD/"required_supplements_letter.pdf"
OUT_PDF=REPORT/"Garbage_Classification_Deep_Learning_Project_Report_ImagesFixed.pdf"
OUT_DOCX=REPORT/"Garbage_Classification_Deep_Learning_Project_Report_ImagesFixed.docx"
PAGE_DIR=SOURCE_BUILD/"final_render"
PAGE_DIR.mkdir(exist_ok=True)

# Build a two-panel overview from the user's supplied real interface screenshots.
shot_dir=BUILD/"user_screenshots"
upload_img=PILImage.open(shot_dir/"upload-screen.png").convert("RGB")
camera_img=PILImage.open(shot_dir/"camera-screen.png").convert("RGB")
panel_w=1080; gap=36; label_h=72; pad=28
panel_h=round(panel_w*upload_img.height/upload_img.width)
montage=PILImage.new("RGB",(pad*2+panel_w*2+gap,pad*2+panel_h+label_h),(247,250,248))
draw=ImageDraw.Draw(montage)
try:
    label_font=ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf",38)
except Exception:
    label_font=ImageFont.load_default()
for idx,(img,label) in enumerate([(upload_img,"UPLOAD A PHOTO"),(camera_img,"SCAN LIVE")]):
    x=pad+idx*(panel_w+gap); y=pad+label_h
    scaled=img.resize((panel_w,panel_h),PILImage.Resampling.LANCZOS)
    montage.paste(scaled,(x,y))
    draw.text((x,y-55),label,font=label_font,fill=(17,76,56))
montage.save(shot_dir/"interface-comparison.png")

# Rebuild the repository-authored report with the university paper and text settings.
original=(ROOT/"project_archive"/"presentation-build"/"create_full_report.py").read_text(encoding="utf-8")
original=original.replace("from reportlab.lib.units import mm","from reportlab.lib.units import mm, inch")
old_split='p("The validation split is configured with seed 42 for reproducibility. A fixed seed makes the split repeatable under the same loader and data ordering; it does not by itself establish that the split is stratified or free of near duplicates."), PageBreak()]'
new_split='p("The validation split is configured with seed 42 for reproducibility. A fixed seed makes the split repeatable under the same loader and data ordering; it does not by itself establish that the split is stratified or free of near duplicates."), sec("5.4 Data readiness for evaluation"), p("The category distribution affects how an aggregate result should be read. Clothes forms 34.3% of the recorded set and shoes 12.7%, while each glass class has fewer than 800 images. Class weighting changes each class contribution during training but does not add examples or demonstrate equal class-level performance. The report therefore treats the saved overall validation accuracy as a split-specific summary, not as a balanced accuracy measure."), p("The seed-42 split supports repeatability when the same dataset, class order, loader and software are used. The repository does not document split indices, duplicate detection, label review or a separate held-out test set. Before wider use, preserve the exact data version and split, audit labels and near-duplicates, then evaluate a separate test set with per-class precision, recall, F1 and a confusion matrix."), make_table([["READINESS CHECK","CURRENT EVIDENCE"],["Dataset source and licence","Not recorded in the supplied repository."],["Duplicate and label audit","No audit report is present."],["Independent held-out test","No separate test set is documented."],["Class-level metrics","Per-class precision/recall/F1 are not recorded."]],[51*mm,100*mm]), PageBreak()]'
if old_split not in original: raise RuntimeError("Could not find the data-split narrative insertion point")
original=original.replace(old_split,new_split)
replacements=[
 ('from reportlab.lib.pagesizes import A4','from reportlab.lib.pagesizes import letter'),
 ('PAGE_W, PAGE_H = A4','PAGE_W, PAGE_H = letter'),
 ('pagesize=A4','pagesize=letter'),
 ('ROOT = Path(r"C:\\Users\\HP\\Desktop\\data science\\7TH-SEMESTER-PROJECT-")',f'ROOT = Path({str(ROOT)!r})'),
 ('BUILD = ROOT / ".ppt-team-build"',f'BUILD = Path({str(SOURCE_BUILD)!r})'),
 ('OUT = ROOT / "Report" / "EcoSort_AI_Professional_Project_Report_50_Page_Edition.pdf"',f'OUT = Path({str(SOURCE_PDF)!r})'),
 ('preview_path = BUILD / "slide-08.png"','preview_path = ROOT / "project_archive" / "presentation-build" / "report-pages" / "page-08.png"'),
 ('preview.crop((72, 155, 620, 610)).resize((1096, 910)).save(BUILD / "upload_screen_detail.png")','preview.resize((1096, 910)).save(BUILD / "upload_screen_detail.png")'),
 ('preview.crop((660, 155, 1208, 610)).resize((1096, 910)).save(BUILD / "camera_screen_detail.png")','preview.resize((1096, 910)).save(BUILD / "camera_screen_detail.png")'),
 ('ROOT / "garbage_classification" / class_name','ROOT / "garbage_classification Dataset" / class_name'),
 ('add_image(BUILD/"slide-08.png", 170*mm, 115*mm)','add_image(ROOT/".codex_format_alignment"/"user_screenshots"/"interface-comparison.png", 165*mm, 74*mm)'),
 ('170*mm','165*mm'),
 ('fontName="Helvetica", fontSize=9.35, leading=14.1','fontName="Times-Roman", fontSize=12, leading=18'),
 ('fontName="Helvetica-Bold", fontSize=21, leading=25','fontName="Times-Bold", fontSize=18, leading=22'),
 ('fontName="Helvetica-Bold", fontSize=13.4, leading=17','fontName="Times-Bold", fontSize=16, leading=19'),
 ('leftMargin=20*mm, rightMargin=20*mm, topMargin=21*mm, bottomMargin=20*mm','leftMargin=1.38*inch, rightMargin=.49*inch, topMargin=.98*inch, bottomMargin=.49*inch'),
 ('Frame(20*mm, 20*mm, PAGE_W-40*mm, PAGE_H-40*mm','Frame(1.38*inch, .49*inch, PAGE_W-1.87*inch, PAGE_H-1.47*inch'),
 ('Frame(20*mm, 19*mm, PAGE_W-40*mm, PAGE_H-40*mm','Frame(1.38*inch, .49*inch, PAGE_W-1.87*inch, PAGE_H-1.47*inch'),
 ('20*mm, PAGE_H-14*mm, PAGE_W-20*mm, PAGE_H-14*mm','1.38*inch, PAGE_H-.72*inch, PAGE_W-.49*inch, PAGE_H-.72*inch'),
 ('20*mm, PAGE_H-11*mm','1.38*inch, PAGE_H-.61*inch'),
 ('PAGE_W-20*mm, PAGE_H-11*mm','PAGE_W-.49*inch, PAGE_H-.61*inch'),
 ('20*mm, 16*mm, PAGE_W-20*mm, 16*mm','1.38*inch, .49*inch, PAGE_W-.49*inch, .49*inch'),
 ('PAGE_W - 20*mm, 12*mm','PAGE_W - .49*inch, .35*inch'),
 ('PAGE_W - 20*mm, 12*mm','PAGE_W - .49*inch, .31*inch'),
 ('PAGE_W-1.38*inch, PAGE_H-.61*inch','PAGE_W-.49*inch, PAGE_H-.61*inch'),
 ('add_image(BUILD/"upload_screen_detail.png", 162*mm, 145*mm)','add_image(ROOT/".codex_format_alignment"/"user_screenshots"/"upload-screen.png", 162*mm, 141*mm)'),
 ('add_image(BUILD/"camera_screen_detail.png", 162*mm, 145*mm)','add_image(ROOT/".codex_format_alignment"/"user_screenshots"/"camera-screen.png", 162*mm, 141*mm)'),
 ("The figures below are interface previews reconstructed from the repository's HTML and JavaScript controls. They document the implemented design and workflow but are not screenshots captured from a running Flask/TensorFlow server. The example object image is from the project's dataset.","The following two screenshots, supplied by the project team on 24 November 2026, show the implemented photo-upload and Scan Live interface states. The upload screen is waiting for a photo; the camera screen shows camera ready while the camera is off. Neither image shows a completed prediction. Full-size captures are included in Chapters 30 and 31."),
 ("Figure 12.1  EcoSort AI upload and camera interface previews. Upload flow includes model selection and category action; camera flow shows capture, review, and prediction states.","Figure 12.1. Overview of the team-supplied EcoSort AI upload and Scan Live interface screenshots, captured on 24 November 2026."),
 ("The following project-interface preview is derived from the implemented HTML/CSS interface and presentation capture. It illustrates the upload-oriented screen and is not a screenshot of a verified live hosted service.","The following screenshot was supplied by the project team on 24 November 2026. It shows the implemented upload screen before a photo is selected; it documents the interface state shown and does not demonstrate a completed prediction."),
 ("Figure 30.1. EcoSort AI upload screen preview. The user selects an image and requests server-side classification.","Figure 30.1. Team-supplied EcoSort AI upload-screen screenshot, captured on 24 November 2026."),
 ("This camera-screen preview shows the live preview and capture-oriented controls represented in the project interface. It documents the intended browser workflow, not evidence that a camera was active in the screenshot or that a particular host deployment was tested.","The following screenshot was supplied by the project team on 24 November 2026. It shows the Scan Live tab in its camera-ready state with the camera off; the screenshot does not show a captured item or prediction result."),
 ("Figure 31.1. Camera workflow interface preview. Capture first, review the still image, then explicitly request prediction.","Figure 31.1. Team-supplied EcoSort AI Scan Live screenshot, showing the camera-ready state, captured on 24 November 2026."),
]
for old,new in replacements:
    if old not in original:
        print("NOTE: source pattern not found:",old[:90])
    original=original.replace(old,new)
(SOURCE_BUILD/"generate_letter_source.py").write_text(original,encoding="utf-8")
runpy.run_path(str(SOURCE_BUILD/"generate_letter_source.py"),run_name="__main__")

source=PdfReader(str(SOURCE_PDF))
source_text=[p.extract_text() or "" for p in source.pages]
chapter_rows=[]
chapter_page_indices={}
for i,text in enumerate(source_text):
    match=re.search(r"(?m)^\s*(\d{2})\s*\|\s*([^\n]+)",text)
    if not match: continue
    number=int(match.group(1)); title=match.group(2).strip()
    if 1<=number<=38 and number not in chapter_page_indices:
        chapter_page_indices[number]=i
        chapter_rows.append((number,title))
start_candidates=[i for i,t in enumerate(source_text) if "01  |  Introduction" in t and "1.1 Background" in t]
if not start_candidates:
    start_candidates=[i for i,t in enumerate(source_text) if re.search(r"01\s*\|\s*Introduction",t) and "1.1 Background" in t]
if not start_candidates:
    raise RuntimeError("Could not find the first substantive chapter page in generated source PDF")
body_start=start_candidates[0]
if len(chapter_page_indices)<30:
    raise RuntimeError(f"Expected the full detailed report body; found {len(chapter_page_indices)} numbered sections")
chapter_rows.sort()
body_page=lambda source_index: source_index-body_start+1
canonical_titles=["Introduction","Objectives and Scope","Technical Background and Related Work","System Overview","Dataset Description","Tools and Technologies","Data Preprocessing","Methodology","Model Architecture and Training","Web Application Implementation","System Architecture","User Interface Screenshots","Results and Discussion","Verification and Test Plan","Deployment and Operation","Privacy, Security, and Operational Risks","Advantages and Limitations","Future Scope","Software Stack at a Glance","Python Runtime and Dependency Management","TensorFlow and Keras","MobileNetV2 Architecture","Classifier Head and Transfer-Learning Stages","Augmentation and Class Imbalance","Training Controls and Reproducibility","NumPy, Pillow, and Image Preparation","Flask Backend and Prediction API","Browser Interface: HTML, CSS, and JavaScript","Progressive Web App and Camera APIs","Interface Screenshot: Upload Workflow","Interface Screenshot: Camera Capture","Dataset Sample Gallery","Class Distribution and Results Interpretation","User Guide: Photo Upload","User Guide: Camera and Troubleshooting","Deployment Environment and Compatibility","Integration, Verification, and Traceability","Conclusion"]
def first_source_page(*needles):
    for idx,txt in enumerate(source_text[body_start:],body_start):
        if all(re.search(n,txt,re.I) for n in needles): return idx
    raise RuntimeError(f"Could not locate report content page for {needles}")
def source_body_page(source_index): return body_page(source_index)
references_index=next(i for i in range(chapter_page_indices[max(chapter_page_indices)]+1,len(source_text)) if re.search(r"References",source_text[i],re.I) and "[1]" in source_text[i])
appendix_index=first_source_page(r"39\s*\|\s*Appendices")

meta=json.loads((ROOT/"models"/"metadata.json").read_text(encoding="utf-8"))
counts=meta["class_counts"]
total=sum(counts.values())

# A level-2 DFD for the actual image inference path.
def make_level2(path):
    w,h=1900,1120; im=PILImage.new("RGB",(w,h),(250,252,250)); d=ImageDraw.Draw(im)
    fonts=Path(r"C:\Windows\Fonts")
    def font(name,size):
        for f in [fonts/name,fonts/"arial.ttf"]:
            if f.exists(): return ImageFont.truetype(str(f),size)
        return ImageFont.load_default()
    titlef=font("arialbd.ttf",44); boxf=font("arialbd.ttf",26); smallf=font("arial.ttf",20)
    dark=(16,62,46); green=(23,101,71); pale=(233,244,237); blue=(228,239,246); gray=(244,246,246)
    d.text((70,35),"DATA FLOW DIAGRAM — LEVEL 2: IMAGE PREDICTION",font=titlef,fill=dark)
    def box(x,y,bw,bh,title,sub,fill=pale):
        d.rounded_rectangle((x,y,x+bw,y+bh),radius=22,fill=fill,outline=green,width=4)
        lines=title.split("\n")
        yy=y+18
        for ln in lines:
            b=d.textbbox((0,0),ln,font=boxf); d.text((x+(bw-(b[2]-b[0]))/2,yy),ln,font=boxf,fill=dark); yy+=32
        for ln in sub.split("\n"):
            b=d.textbbox((0,0),ln,font=smallf); d.text((x+(bw-(b[2]-b[0]))/2,yy+2),ln,font=smallf,fill=(65,82,72)); yy+=27
    def arrow(a,b,label=""):
        d.line([a,b],fill=green,width=5)
        import math
        ang=math.atan2(b[1]-a[1],b[0]-a[0]); L=17
        p1=(b[0]-L*math.cos(ang-.5),b[1]-L*math.sin(ang-.5));p2=(b[0]-L*math.cos(ang+.5),b[1]-L*math.sin(ang+.5))
        d.polygon([b,p1,p2],fill=green)
        if label:
            mx=(a[0]+b[0])//2;my=(a[1]+b[1])//2-27
            d.text((mx-45,my),label,font=smallf,fill=dark)
    box(45,390,230,145,"USER","upload or captured still",blue)
    box(350,370,270,190,"1.0 Validate request","file type, size, model name")
    box(700,370,270,190,"2.0 Decode image","Pillow; convert to RGB")
    box(1050,370,270,190,"3.0 Prepare input","resize to 224 × 224; array")
    box(1400,370,250,190,"4.0 Run model","selected Keras model")
    box(1510,760,310,190,"5.0 Rank and map","scores + class_names; top five")
    box(1030,790,320,140,"6.0 Build response","JSON or rendered result")
    box(370,780,340,150,"ERROR RESPONSE","invalid image/model or decode failure",(251,239,234))
    box(1040,90,360,170,"D1 — Model artifacts","*.keras weights / architecture",gray)
    box(1460,90,360,170,"D2 — metadata.json","class_names; image_size",gray)
    arrow((275,462),(350,462),"image + model")
    arrow((620,462),(700,462),"validated")
    arrow((970,462),(1050,462),"RGB image")
    arrow((1320,462),(1400,462),"224×224 tensor")
    arrow((1525,560),(1590,760),"probabilities")
    arrow((1510,845),(1350,845),"ranked classes")
    arrow((1030,860),(620,560),"result / error")
    arrow((1110,260),(1470,370),"model file")
    arrow((1640,260),(1605,370),"class order")
    arrow((485,560),(485,780),"validation error")
    d.text((50,1030),"Camera frames follow the same inference path after capture; the app does not persist user images or predictions in a relational database.",font=smallf,fill=(65,82,72))
    im.save(path)

level2_png=SOURCE_BUILD/"dfd_level2.png"
make_level2(level2_png)

styles=getSampleStyleSheet()
styles.add(ParagraphStyle(name="Body",fontName="Times-Roman",fontSize=12,leading=18,alignment=TA_JUSTIFY,spaceAfter=9,textColor=colors.HexColor("#1B1B1B")))
styles.add(ParagraphStyle(name="H1",fontName="Times-Bold",fontSize=16,leading=20,alignment=TA_CENTER,spaceAfter=14,keepWithNext=True))
styles.add(ParagraphStyle(name="H2",fontName="Times-Bold",fontSize=12,leading=16,spaceBefore=9,spaceAfter=5,keepWithNext=True))
styles.add(ParagraphStyle(name="SignatureLine",fontName="Times-Roman",fontSize=12,leading=18,alignment=TA_LEFT,spaceAfter=9))
styles.add(ParagraphStyle(name="Cell",fontName="Times-Roman",fontSize=9.2,leading=11,spaceAfter=0))
styles.add(ParagraphStyle(name="CellHead",fontName="Times-Bold",fontSize=9.2,leading=11,spaceAfter=0))
styles.add(ParagraphStyle(name="TOC",fontName="Times-Roman",fontSize=8.2,leading=9.7,spaceAfter=0))
styles.add(ParagraphStyle(name="CoverTitle",fontName="Times-Bold",fontSize=24,leading=29,alignment=TA_CENTER,spaceAfter=15))
styles.add(ParagraphStyle(name="CoverProject",fontName="Times-Bold",fontSize=17,leading=22,alignment=TA_CENTER,spaceAfter=8))
styles.add(ParagraphStyle(name="CoverBody",fontName="Times-Roman",fontSize=13,leading=19,alignment=TA_CENTER,spaceAfter=8))
styles["CoverBody"].spaceAfter=4
styles.add(ParagraphStyle(name="CoverBold",fontName="Times-Bold",fontSize=14,leading=20,alignment=TA_CENTER,spaceAfter=4))
styles.add(ParagraphStyle(name="CoverUni",fontName="Times-Bold",fontSize=15,leading=19,alignment=TA_CENTER,spaceAfter=4))

def par(text,style="Body"):
    safe=str(text).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")
    return Paragraph(safe,styles[style])

def tbl(rows,widths,compact=False):
    data=[]
    for i,row in enumerate(rows):
        data.append([Paragraph(str(v).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"),styles["CellHead" if i==0 else ("TOC" if compact else "Cell")]) for v in row])
    t=Table(data,colWidths=widths,repeatRows=1,hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E8EEE9")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.HexColor("#163E2B")),
        ("GRID",(0,0),(-1,-1),.45,colors.HexColor("#9AA89D")),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),4 if compact else 6),
        ("RIGHTPADDING",(0,0),(-1,-1),4 if compact else 6),
        ("TOPPADDING",(0,0),(-1,-1),2 if compact else 5),
        ("BOTTOMPADDING",(0,0),(-1,-1),2 if compact else 5),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F7F9F7")]),
    ]))
    return t

def roman(n):
    vals=[(50,"l"),(40,"xl"),(10,"x"),(9,"ix"),(5,"v"),(4,"iv"),(1,"i")]
    out=""
    for v,sym in vals:
        while n>=v: out+=sym;n-=v
    return out

def prelim_footer(c,d):
    c.saveState()
    if d.page>1:
        c.setFont("Times-Roman",9);c.setFillColor(colors.HexColor("#555555"));c.drawCentredString(letter[0]/2,.28*inch,roman(d.page-1))
    c.restoreState()

logo=BUILD/"image1.jpeg"
if not logo.exists():
    raise FileNotFoundError("University seal extracted from the supplied cover template is missing")

# Cover keeps the supplied title-page hierarchy, with project identity fields replaced.
cover_story=[Spacer(1,.2*inch),Image(str(logo),width=1.42*inch,height=1.05*inch),Spacer(1,.1*inch),
    par("IK GUJRAL PUNJAB TECHNICAL UNIVERSITY CAMPUS, HOSHIARPUR","CoverUni"),
    par("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING","CoverUni"),Spacer(1,.22*inch),
    par("PROJECT REPORT","CoverTitle"),par("GARBAGE CLASSIFICATION USING DEEP LEARNING","CoverProject"),Spacer(1,.15*inch),
    par("PROJECT-2  ·  7TH SEMESTER","CoverBold"),Spacer(1,.12*inch),
    par("SUBMITTED BY","CoverBold"),par("Arshpreet Singh  (2323174)","CoverBody"),
    par("Sarabjit Singh  (2323218)","CoverBody"),par("Pooja  (2323200)","CoverBody"),Spacer(1,.12*inch),
    par("SUBMITTED TO","CoverBold"),par("Dr. Brijesh Bakariya  |  Project Mentor and HOD","CoverBody"),
    par("Er. Satyam Sharma  |  Project Co-Mentor","CoverBody"),Spacer(1,.12*inch),
    par("Submission Date: 26 November 2026","CoverBody")]
SimpleDocTemplate(str(SOURCE_BUILD/"cover.pdf"),pagesize=letter,leftMargin=inch,rightMargin=inch,topMargin=inch,bottomMargin=inch,title="EcoSort AI Project Report Cover").build(cover_story)

toc_rows=[[f"CHAPTER {n}",canonical_titles[n-1],str(n)] for n,title in chapter_rows]
body_count=len(source.pages)-body_start
toc_rows += [["REFERENCES","Technical references",str(source_body_page(references_index))],
             ["CHAPTER 39","Appendices",str(source_body_page(appendix_index))],
             ["SUPPLEMENT 1","Level-2 DFD: image prediction",str(body_count+1)],
             ["SUPPLEMENT 2","Data persistence and schema inventory",str(body_count+2)],
             ["SUPPLEMENT 3","EcoSort AI project datasheet",str(body_count+3)]]
dataset_page=source_body_page(chapter_page_indices[5])
tools_page=source_body_page(chapter_page_indices[6])
results_page=source_body_page(chapter_page_indices[13])
verify_page=source_body_page(chapter_page_indices[14])
deploy_page=source_body_page(chapter_page_indices[36])
trace_page=source_body_page(chapter_page_indices[37])
tables_rows=[
 ["TABLE NO.","TOPIC","PAGE NO."],
 ["Table 5.1","Dataset class counts",str(dataset_page)],
 ["Table 5.2","Dataset readiness checks",str(source_body_page(next(i for i,t in enumerate(source_text) if "READINESS CHECK" in t)))],
 ["Table 6.1","Tools and technologies used",str(tools_page)],
 ["Table 13.1","Saved validation metrics",str(13)],
 ["Table 14.1","Verification and test-plan matrix",str(verify_page)],
 ["Table 36.1","Deployment environment and compatibility",str(deploy_page)],
 ["Table 37.1","Integration and requirements traceability",str(trace_page)],
 ["Table 44.1","Data persistence and artifact inventory",str(body_count+2)],
 ["Table 45.1","Project technical datasheet",str(body_count+3)],
]
figures_rows=[
 ["FIGURE NO.","TOPIC","PAGE NO."],
 ["Figure 4.1","Level-0 context data-flow diagram",str(source_body_page(chapter_page_indices[4]))],
 ["Figure 4.2","Level-1 data-flow diagram",str(source_body_page(chapter_page_indices[4]))],
 ["Figure 6.1","EcoSort AI tools and technology stack",str(tools_page)],
 ["Figure 12.1","Upload and camera interface overview",str(source_body_page(chapter_page_indices[12]))],
 ["Figure 13.1","Training and validation history",str(13)],
 ["Figure 30.1","Upload workflow interface",str(source_body_page(chapter_page_indices[30]))],
 ["Figure 31.1","Camera capture and review workflow",str(source_body_page(chapter_page_indices[31]))],
 ["Figure 32.1","Representative dataset samples",str(source_body_page(chapter_page_indices[32]))],
 ["Figure 43.1","Level-2 inference data-flow diagram",str(body_count+1)],
]
abbr_rows=[["TERM","MEANING IN THIS PROJECT"],["AI","Artificial intelligence"],["API","Application programming interface"],["CNN","Convolutional neural network"],["DFD","Data-flow diagram"],["HTTP","Hypertext transfer protocol"],["JPEG","Joint Photographic Experts Group image format"],["JSON","JavaScript Object Notation"],["PWA","Progressive web application"],["RGB","Red, green and blue image channels"],["UI","User interface"],["URL","Uniform resource locator"],["WSGI","Web Server Gateway Interface"]]

prelim=[]
prelim += [par("DECLARATION AND CERTIFICATION","H1"),
    par("We hereby declare that the project report entitled “Garbage Classification Using Deep Learning” is an authentic record of the work undertaken by our project team for Project-2 in the 7th semester at IK Gujral Punjab Technical University, Hoshiarpur Campus. The project details and reported results are based on the submitted project repository and model metadata. We have cited the technical references used in preparing this report."),
    Spacer(1,8),tbl([["STUDENT NAME AND ROLL NO.","SIGNATURE","DATE"],["Arshpreet Singh (2323174)","________________________","____________"],["Sarabjit Singh (2323218)","________________________","____________"],["Pooja (2323200)","________________________","____________"]],[2.3*inch,2.1*inch,1.45*inch]),Spacer(1,15),
    par("Certified that the above declaration is correct to the best of our knowledge and belief. The project report is submitted for examination and institutional review.") ,Spacer(1,22),
    tbl([["PROJECT MENTOR / HOD","PROJECT CO-MENTOR"],["Dr. Brijesh Bakariya","Er. Satyam Sharma"],["Signature: ______________","Signature: ______________"],["Date: ___________________","Date: ___________________"]],[3.0*inch,3.0*inch]),Spacer(1,22),
    par("Head of Department (Signature and Seal): ______________________________________________","SignatureLine"),PageBreak()]
prelim += [par("ACKNOWLEDGEMENT","H1"),par("We express our sincere gratitude to Dr. Brijesh Bakariya, Project Mentor and Head of the Department of Computer Science and Engineering, for his guidance and academic support throughout the project. We also thank Er. Satyam Sharma, Project Co-Mentor, for his suggestions and assistance. We acknowledge IK Gujral Punjab Technical University, Hoshiarpur Campus, for providing the academic setting in which this project was completed. Finally, we thank our families and colleagues for their encouragement."),Spacer(1,14),par("Project Team","H2"),par("Arshpreet Singh (2323174)\nSarabjit Singh (2323218)\nPooja (2323200)"),PageBreak()]
prelim += [par("ABSTRACT","H1"),par("EcoSort AI is an image-classification prototype that identifies one of twelve waste categories from an uploaded photograph or a still image captured with a browser camera. A MobileNetV2 transfer-learning model is served through a Flask application. The interface returns a ranked set of category suggestions and supports photo upload and camera capture with review before prediction."),par(f"The dataset metadata records {total:,} images across twelve categories. The model input size is 224 × 224 pixels. Training uses a seeded validation split, augmentation, class weighting, early stopping and fine-tuning. Saved metadata reports {meta['validation_accuracy']*100:.2f}% validation accuracy and {meta['validation_loss']:.4f} validation loss. These values describe the recorded validation split and do not establish independent or field performance."),par("Keywords: waste classification, computer vision, MobileNetV2, TensorFlow, Flask, image recognition."),PageBreak()]
prelim += [par("ABOUT THE UNIVERSITY AND DEPARTMENT","H1"),par("This project report is presented through the Department of Computer Science and Engineering at IK Gujral Punjab Technical University, Hoshiarpur Campus. The cover and academic front matter follow the supplied university project-report examples."),par("EcoSort AI is a software project that combines dataset preparation, a trained image classifier, a web inference service and a browser interface. The report documents the project as implemented in its repository and distinguishes measured validation results from future work that still requires independent evaluation."),PageBreak()]
prelim += [par("TABLE OF CONTENTS","H1"),tbl([["CHAPTER NO.","TOPIC","PAGE NO."]]+toc_rows,[1.25*inch,4.55*inch,.65*inch],compact=True),PageBreak()]
prelim += [par("LIST OF TABLES","H1"),tbl(tables_rows,[1.2*inch,4.5*inch,.75*inch]),PageBreak()]
prelim += [par("LIST OF FIGURES","H1"),tbl(figures_rows,[1.25*inch,4.45*inch,.75*inch]),PageBreak()]
prelim += [par("ABBREVIATIONS AND NOMENCLATURE","H1"),par("The terms below are used in the report. Standard image and software terms are listed where they help interpret the system diagrams and implementation."),tbl(abbr_rows,[1.2*inch,5.2*inch]),PageBreak()]
INNER=SOURCE_BUILD/"preliminary.pdf"
SimpleDocTemplate(str(INNER),pagesize=letter,leftMargin=1.38*inch,rightMargin=.49*inch,topMargin=.98*inch,bottomMargin=.49*inch,title="EcoSort AI Preliminary Pages").build(prelim,onFirstPage=prelim_footer,onLaterPages=prelim_footer)

# The final substantive pages complete the specific checklist without inventing a database.
supp_story=[par("SUPPLEMENT 1  DFD LEVEL 2","H1"),
    par("The Level-2 diagram expands the prediction process into validation, image decoding, RGB conversion, resizing, model inference, class mapping and response generation. Both upload and captured-camera images use the same inference path. Invalid files or model choices return an error to the user."),
    Image(str(level2_png),width=6.55*inch,height=3.86*inch),
    par("Figure 43.1. Detailed image-prediction data flow for EcoSort AI.","Cell"),
    par("The model weights and class-order metadata are read from project files. Uploaded or captured user images are processed for the request; the application does not configure a relational database for user records or prediction history."),PageBreak(),
    par("SUPPLEMENT 2  DATA PERSISTENCE AND SCHEMA","H1"),
    par("The supplied EcoSort application does not define a relational database, SQL tables, ORM models or persistent user accounts. A database schema therefore does not exist in this project version. The inventory below documents the actual data artifacts and transient request data instead of presenting a fictional schema."),
    tbl([["ARTIFACT / FIELD","TYPE","ROLE / KEY DETAILS"],
         ["class_names","list[str]","Ordered class-label key; output index maps to the same index here."],
         ["class_counts","dict[str, int]","Dataset count keyed by category name."],
         ["image_size","list[int]","Model input dimensions [224, 224]."],
         ["validation_accuracy","float","Recorded validation accuracy; not a relational key."],
         ["validation_loss","float","Recorded validation loss; not a relational key."],
         ["*.keras","model file","Saved model weights and architecture used for inference."],
         ["Uploaded/captured image","in-memory image","Decoded, converted to RGB, resized and classified for a request; no DB row."],
         ["Prediction response","JSON/object","Top result, ranked candidates and confidence status returned to the client."]],
        [1.55*inch,1.25*inch,3.6*inch]),
    par("Table 44.1. EcoSort AI data and persistence artifact inventory.","Cell"),
    par("If a future version stores users, uploads or prediction history, it will require an explicit privacy policy, retention design, access controls, database schema and migrations. Those features are outside the current implementation."),PageBreak(),
    par("SUPPLEMENT 3  ECOSORT AI PROJECT DATASHEET","H1"),
    par("This datasheet summarizes the configured model and application behavior documented in the project repository."),
    tbl([["ITEM","RECORDED PROJECT DETAIL"],
         ["Project","EcoSort AI — AI-powered waste-image classification"],
         ["Task","Single-image, single-label image classification"],
         ["Categories","12 waste categories (see model metadata and dataset gallery)"],
         ["Dataset size",f"{total:,} images recorded in class-count metadata"],
         ["Image input","RGB, resized to 224 × 224 pixels"],
         ["Model backbone","MobileNetV2 transfer learning with a 12-class output head"],
         ["Training measures",f"Validation accuracy {meta['validation_accuracy']*100:.2f}%; validation loss {meta['validation_loss']:.4f}"],
         ["Interface","Flask web app; photo upload and captured-camera still workflow"],
         ["Accepted formats","JPG, JPEG, PNG, BMP and WEBP"],
         ["Upload limit","8 MB configured request limit"],
         ["Output","Top category and up to five ranked candidate classes"],
         ["Authentication / DB","Not implemented in the supplied application"],
         ["Operational note","Camera access requires browser permission and a secure context such as HTTPS or localhost"]],
        [1.65*inch,4.75*inch]),
    par("Table 45.1. EcoSort AI project technical datasheet.","Cell"),
    par("No login screen is included because the current app has no authentication feature. The report documents the home/upload, camera, prediction and error states that exist in the supplied interface."),
    par("The saved validation metrics should be updated only after the team documents the model version, evaluation split and reproducible run used to produce them.")]
SUPP=SOURCE_BUILD/"supplements.pdf"
SimpleDocTemplate(str(SUPP),pagesize=letter,leftMargin=1.38*inch,rightMargin=.49*inch,topMargin=.98*inch,bottomMargin=.49*inch,title="EcoSort AI Required Supplementary Sections").build(supp_story)

def _footer_page(number,pagesize):
    b=BytesIO();c=Canvas(b,pagesize=pagesize);c.setFont("Times-Roman",9);c.setFillColor(colors.HexColor("#555555"));c.drawRightString(pagesize[0]-.49*inch,.31*inch,str(number));c.save();b.seek(0);return PdfReader(b).pages[0]

cover=PdfReader(str(SOURCE_BUILD/"cover.pdf")); prelim_pdf=PdfReader(str(INNER)); supplement_pdf=PdfReader(str(SUPP))
front_pages=list(cover.pages)+list(prelim_pdf.pages)
front_count=len(front_pages)
if front_count<9:
    raise RuntimeError(f"Unexpectedly short front matter: {front_count} pages")
writer=PdfWriter()
for p in front_pages: writer.add_page(p)
body_pages=list(source.pages[body_start:])
scale=letter[1]/source.pages[body_start].mediabox.height
tx=letter[0]-source.pages[body_start].mediabox.width*scale
if tx<0: tx=0
arabic_total=len(body_pages)+len(supplement_pdf.pages)
for n,source_page in enumerate(body_pages,1):
    blank=PageObject.create_blank_page(width=letter[0],height=letter[1])
    blank.merge_transformed_page(source_page,Transformation().scale(scale).translate(tx,0))
    # Mask the original source pagination and replace it with the final Arabic body number.
    packet=BytesIO(); c=Canvas(packet,pagesize=letter)
    c.setFillColor(colors.white);c.rect(letter[0]-1.05*inch,.20*inch,.80*inch,.32*inch,fill=1,stroke=0)
    c.setFillColor(colors.HexColor("#555555"));c.setFont("Times-Roman",9);c.drawRightString(letter[0]-.49*inch,.31*inch,str(n));c.save();packet.seek(0)
    blank.merge_page(PdfReader(packet).pages[0]);writer.add_page(blank)
for i,p in enumerate(supplement_pdf.pages,len(body_pages)+1):
    p.merge_page(_footer_page(i,letter))
    writer.add_page(p)

with OUT_PDF.open("wb") as f: writer.write(f)
print(f"Created {OUT_PDF} ({len(writer.pages)} pages; front matter {front_count}, body {len(body_pages)}, required supplements {len(supplement_pdf.pages)})")

# Build a page-faithful Letter Word print copy, one rendered PDF page per page.
pdftoppm=Path(r"C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin\pdftoppm.exe")
for old in PAGE_DIR.glob("page-*.jpg"): old.unlink()
subprocess.run([str(pdftoppm),"-jpeg","-r","110","-jpegopt","quality=86",str(OUT_PDF),str(PAGE_DIR/"page")],check=True,capture_output=True)
images=sorted(PAGE_DIR.glob("page-*.jpg"),key=lambda p:int(re.search(r"-(\d+)\.jpg$",p.name).group(1)))
if len(images)!=len(writer.pages): raise RuntimeError("Word image page count does not match the final PDF")
doc=Document();sec=doc.sections[0];sec.page_width=Inches(8.5);sec.page_height=Inches(11)
sec.top_margin=Inches(0);sec.bottom_margin=Inches(0);sec.left_margin=Inches(0);sec.right_margin=Inches(0)
normal=doc.styles["Normal"];normal.paragraph_format.space_before=0;normal.paragraph_format.space_after=0;normal.paragraph_format.line_spacing=1
for i,img in enumerate(images):
    p=doc.add_paragraph();p.paragraph_format.space_before=0;p.paragraph_format.space_after=0;p.paragraph_format.keep_together=True
    p.add_run().add_picture(str(img),width=Inches(8.5),height=Inches(11))
    if i<len(images)-1: doc.add_page_break()
doc.core_properties.title="EcoSort AI University Format Project Report"
doc.core_properties.subject=f"Page-faithful {len(images)}-page Letter print copy; body text in the accompanying PDF is also available as searchable text."
doc.core_properties.author="EcoSort AI project team"
doc.save(OUT_DOCX)
print(f"Created {OUT_DOCX} ({len(images)} page images)")
