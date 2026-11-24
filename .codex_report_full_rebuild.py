from pathlib import Path
import re, json, runpy

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pypdf import PdfReader
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image as RLImage, KeepTogether
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent
REPORT = ROOT / "Report"
BUILD = ROOT / ".codex_full_report_build"
BUILD.mkdir(exist_ok=True)
OLD_SOURCE = ROOT / "project_archive" / "presentation-build" / "create_full_report.py"
OLD_PDF = REPORT / "EcoSort_AI_Professional_Project_Report_50_Page_Edition.pdf"
DOCX_OUT = REPORT / "EcoSort_AI_Project_Report_Professional_Full.docx"
PDF_OUT = REPORT / "EcoSort_AI_Project_Report_Professional_Full.pdf"

# Recreate verified, project-authored source report PDFs from its existing detailed generator.
source = OLD_SOURCE.read_text(encoding="utf-8")
source = source.replace('ROOT = Path(r"C:\\Users\\HP\\Desktop\\data science\\7TH-SEMESTER-PROJECT-")', f'ROOT = Path({str(ROOT)!r})')
source = source.replace('BUILD = ROOT / ".ppt-team-build"', f'BUILD = Path({str(BUILD)!r})')
source = source.replace('OUT = ROOT / "Report" / "EcoSort_AI_Professional_Project_Report_50_Page_Edition.pdf"', f'OUT = Path({str(BUILD / "source_full_rebuild.pdf")!r})')
source = source.replace('preview_path = BUILD / "slide-08.png"', 'preview_path = ROOT / "project_archive" / "presentation-build" / "report-pages" / "page-08.png"')
source = source.replace('preview.crop((72, 155, 620, 610)).resize((1096, 910)).save(BUILD / "upload_screen_detail.png")', 'preview.resize((1096, 910)).save(BUILD / "upload_screen_detail.png")')
source = source.replace('preview.crop((660, 155, 1208, 610)).resize((1096, 910)).save(BUILD / "camera_screen_detail.png")', 'preview.resize((1096, 910)).save(BUILD / "camera_screen_detail.png")')
source = source.replace('ROOT / "garbage_classification" / class_name', 'ROOT / "garbage_classification Dataset" / class_name')
source = source.replace('add_image(BUILD/"slide-08.png", 170*mm, 115*mm)', 'add_image(BUILD/"upload_screen_detail.png", 170*mm, 115*mm)')
(BUILD / "generate_source.py").write_text(source, encoding="utf-8")
runpy.run_path(str(BUILD / "generate_source.py"), run_name="__main__")

# The source authored report has 39 chapters. Create readable project-specific technology cards.
tech_path = BUILD / "technology_stack.png"
W,H = 1600,850
im = Image.new("RGB",(W,H),(247,250,247)); d=ImageDraw.Draw(im)
try:
    fbig=ImageFont.truetype("arialbd.ttf",38); fmed=ImageFont.truetype("arialbd.ttf",27); fsmall=ImageFont.truetype("arial.ttf",22)
except Exception:
    fbig=fmed=fsmall=ImageFont.load_default()
d.text((55,30),"ECOSORT AI — TOOLS AND TECHNOLOGIES",font=fbig,fill=(16,62,46))
items=[("Python 3.11","Training notebook and Flask server",(35,89,117)),("TensorFlow / Keras","Model training, saving and inference",(237,117,44)),("MobileNetV2","Pretrained visual feature extractor",(54,116,82)),("Flask","HTTP routes and prediction API",(49,64,75)),("NumPy + Pillow","Image arrays, decoding and resizing",(97,82,154)),("HTML + CSS + JavaScript","Upload, camera and responsive UI",(204,111,58)),("Gunicorn + Render","WSGI serving and deployment setup",(88,108,93)),("PWA + Camera API","Installable shell and still-image capture",(32,128,138))]
for i,(a,b,c) in enumerate(items):
    col=i%2; row=i//2; x=55+col*760; y=115+row*170
    d.rounded_rectangle((x,y,x+700,y+135),radius=22,fill=(255,255,255),outline=(215,226,218),width=4)
    d.rounded_rectangle((x+20,y+20,x+105,y+105),radius=18,fill=c)
    d.text((x+135,y+22),a,font=fmed,fill=(16,62,46)); d.text((x+135,y+70),b,font=fsmall,fill=(65,82,72))
im.save(tech_path)

meta=json.loads((ROOT/"models"/"metadata.json").read_text(encoding="utf-8"))
source_pdf=PdfReader(str(BUILD/"source_full_rebuild.pdf"))
source_text=[p.extract_text() or "" for p in source_pdf.pages]

# DFD diagrams designed for this actual upload/camera classifier flow.
def draw_dfd(path, level):
    w,h=1800,1000; img=Image.new("RGB",(w,h),(250,252,250)); dr=ImageDraw.Draw(img)
    try: titlef=ImageFont.truetype("arialbd.ttf",42); boxf=ImageFont.truetype("arialbd.ttf",27); textf=ImageFont.truetype("arial.ttf",23)
    except Exception: titlef=boxf=textf=ImageFont.load_default()
    dr.text((70,40),f"DATA FLOW DIAGRAM — LEVEL {level}",font=titlef,fill=(16,62,46))
    def box(x,y,ww,hh,label,sub="",fill=(230,243,236)):
        dr.rounded_rectangle((x,y,x+ww,y+hh),radius=24,fill=fill,outline=(23,101,71),width=4)
        lines=label.split("\n")
        for j,t in enumerate(lines):
            bb=dr.textbbox((0,0),t,font=boxf); dr.text((x+(ww-bb[2])/2,y+22+j*39),t,font=boxf,fill=(16,62,46))
        if sub:
            sb=dr.textbbox((0,0),sub,font=textf); dr.text((x+(ww-sb[2])/2,y+hh-52),sub,font=textf,fill=(74,93,82))
    def arrow(a,b,label=""):
        dr.line((a[0],a[1],b[0],b[1]),fill=(22,134,160),width=7)
        import math
        ang=math.atan2(b[1]-a[1],b[0]-a[0]); L=25
        pts=[b,(b[0]-L*math.cos(ang-.48),b[1]-L*math.sin(ang-.48)),(b[0]-L*math.cos(ang+.48),b[1]-L*math.sin(ang+.48))]
        dr.polygon(pts,fill=(22,134,160))
        if label: dr.text(((a[0]+b[0])//2-75,(a[1]+b[1])//2-34),label,font=textf,fill=(35,73,83))
    if level==0:
        box(80,350,290,190,"USER","selects one item",(255,239,222))
        box(550,330,600,240,"ECOSORT AI","image classification process")
        box(1350,350,360,190,"USER","reviews suggestion",(255,239,222))
        arrow((370,430),(550,430),"image / still")
        arrow((1150,430),(1350,430),"top-5 labels + scores")
        arrow((1350,500),(370,500),"new image / retake",)
    else:
        box(50,370,260,180,"USER","upload or camera",(255,239,222))
        box(390,330,310,260,"1. VALIDATE","file, format, model")
        box(780,330,310,260,"2. PREPARE","decode RGB; resize")
        box(1170,330,260,260,"3. PREDICT","Keras model")
        box(1510,370,240,180,"USER","view result",(255,239,222))
        arrow((310,460),(390,460),"image")
        arrow((700,460),(780,460),"valid input")
        arrow((1090,460),(1170,460),"224 × 224")
        arrow((1430,460),(1510,460),"ranked scores")
        box(800,720,300,130,"MODEL + METADATA",".keras + class names",(237,244,250))
        arrow((930,720),(1280,590),"load / class map")
        dr.text((75,780),"Invalid image or model selection returns a clear error to the user.",font=textf,fill=(87,100,92))
    img.save(path)
draw_dfd(BUILD/"dfd_level0.png",0); draw_dfd(BUILD/"dfd_level1.png",1)

# Build a robust DOCX in the sample report structure and an equivalent print PDF.
doc=Document(); sec=doc.sections[0]; sec.page_width=Inches(8.27); sec.page_height=Inches(11.69)
sec.top_margin=Inches(.75); sec.bottom_margin=Inches(.7); sec.left_margin=Inches(.85); sec.right_margin=Inches(.75)
GREEN="176547"; DEEP="103E2E"; TEAL="1686A0"; PALE="F2F7F3"; INK="1C2922"; MUTED="5B6B62"; LINE="D7E2DA"
normal=doc.styles["Normal"]; normal.font.name="Times New Roman"; normal.font.size=Pt(11); normal.font.color.rgb=RGBColor.from_string(INK); normal.paragraph_format.line_spacing=1.2; normal.paragraph_format.space_after=Pt(6)
for st,sz,col in [("Title",24,DEEP),("Heading 1",18,DEEP),("Heading 2",13,GREEN),("Heading 3",11,TEAL)]:
    s=doc.styles[st]; s.font.name="Times New Roman"; s.font.size=Pt(sz); s.font.bold=True; s.font.color.rgb=RGBColor.from_string(col); s.paragraph_format.keep_with_next=True

def shade(cell, color):
    sh=OxmlElement("w:shd"); sh.set(qn("w:fill"),color); cell._tc.get_or_add_tcPr().append(sh)
def set_repeat(row):
    pr=OxmlElement("w:tblHeader"); pr.set(qn("w:val"),"true"); row._tr.get_or_add_trPr().append(pr)
def field(p, text):
    r=p.add_run(); b=OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"),"begin"); t=OxmlElement("w:instrText"); t.set(qn("xml:space"),"preserve"); t.text=text; e=OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"),"end"); r._r.extend([b,t,e])
hp=sec.header.paragraphs[0]; hp.add_run("ECOSORT AI  |  PROJECT REPORT").bold=True; hp.runs[0].font.size=Pt(8); hp.runs[0].font.color.rgb=RGBColor.from_string(GREEN)
fp=sec.footer.paragraphs[0]; fp.alignment=WD_ALIGN_PARAGRAPH.RIGHT; fp.add_run("Page "); field(fp,"PAGE")
def para(txt="",style=None,align=None,bold=False,italic=False):
    x=doc.add_paragraph(style=style); 
    if align is not None: x.alignment=align
    r=x.add_run(txt); r.bold=bold; r.italic=italic; return x
def tbl(rows,widths=None,font=9):
    t=doc.add_table(rows=0,cols=len(rows[0])); t.style="Table Grid"; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        if ri==0:set_repeat(t.rows[-1])
        for j,v in enumerate(row):
            c=cells[j]; c.text=""; c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
            if widths:c.width=Inches(widths[j])
            if ri==0:shade(c,GREEN)
            elif ri%2==0:shade(c,PALE)
            r=c.paragraphs[0].add_run(str(v));r.font.name="Times New Roman";r.font.size=Pt(font)
            if ri==0:r.bold=True;r.font.color.rgb=RGBColor(255,255,255)
    doc.add_paragraph().paragraph_format.space_after=Pt(1); return t
def pic(path,w=6.2,caption=None):
    q=doc.add_paragraph();q.alignment=WD_ALIGN_PARAGRAPH.CENTER;q.paragraph_format.keep_together=True;q.add_run().add_picture(str(path),width=Inches(w))
    if caption: para(caption,align=WD_ALIGN_PARAGRAPH.CENTER,italic=True)
def pb():doc.add_page_break()

# Cover page based on front-page reference.
para("IK GUJRAL PUNJAB TECHNICAL UNIVERSITY, HOSHIARPUR CAMPUS",align=WD_ALIGN_PARAGRAPH.CENTER,bold=True)
para("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING",align=WD_ALIGN_PARAGRAPH.CENTER,bold=True)
pic(ROOT/"static"/"garbage-ai-logo.jpg",0.75)
para("PROJECT REPORT",align=WD_ALIGN_PARAGRAPH.CENTER,bold=True)
para("EcoSort AI", "Title", WD_ALIGN_PARAGRAPH.CENTER)
para("AI-Powered Waste Classification",align=WD_ALIGN_PARAGRAPH.CENTER,bold=True)
para("Garbage Classification with Deep Learning",align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_paragraph()
para("Submitted in partial fulfilment of the requirements for the award of the degree of",align=WD_ALIGN_PARAGRAPH.CENTER)
para("Bachelor of Technology in Computer Science Engineering",align=WD_ALIGN_PARAGRAPH.CENTER,bold=True)
doc.add_paragraph()
tbl([["SUBMITTED BY","SUBMITTED TO"],["Arshpreet Singh\nSarabjit Singh\nPooja\nRoll Nos.: 2323174, 2323218, 2323200","Dr. Brijesh Bakariya\nProject Mentor and HOD, CSE\nEr. Satyam Sharma\nProject Co-Mentor"]],[3.1,3.1],10)
para("Semester: 7th  |  Academic session: 2027  |  Submission date: ____________________",align=WD_ALIGN_PARAGRAPH.CENTER)
para("Hard-bound project report submission",align=WD_ALIGN_PARAGRAPH.CENTER,italic=True)
pb()

para("DECLARATION","Heading 1")
para("We, the undersigned project team, declare that this report describes EcoSort AI project work carried out by our team. Project implementation details and results are based on the project repository and saved model metadata. External technical sources are cited in the references. This declaration should be signed by all listed students after review.")
tbl([["Name","Signature","Date"],["Arshpreet Singh","____________________________","____________"],["Sarabjit Singh","____________________________","____________"],["Pooja","____________________________","____________"]],[2,2.9,1.1],10)
para("Department: Computer Science and Engineering\nInstitution: IK Gujral Punjab Technical University, Hoshiarpur Campus\nProject title: EcoSort AI — AI-Powered Waste Classification\nSemester: 7th; Academic session: 2027")
pb()
para("CERTIFICATE","Heading 1")
para("This is to certify that the project entitled EcoSort AI: AI-Powered Waste Classification has been carried out by Arshpreet Singh, Sarabjit Singh, and Pooja in the Department of Computer Science and Engineering, IK Gujral Punjab Technical University, Hoshiarpur Campus, in partial fulfilment of the requirements for the Bachelor of Technology in Computer Science Engineering under the guidance of Dr. Brijesh Bakariya, Project Mentor and Head of Department, and Er. Satyam Sharma, Project Co-Mentor. Certification is subject to institutional review and signatures.")
for _ in range(3):doc.add_paragraph()
tbl([["Project Mentor / HOD","Project Co-Mentor","External Examiner"],["Dr. Brijesh Bakariya","Er. Satyam Sharma","Name: __________________"],["Signature: ______________","Signature: ______________","Signature: ______________"],["Date: ___________________","Date: ___________________","Date: ___________________"]],[2.0,2.0,2.0],9)
para("Institution seal / remarks", "Heading 2")
for _ in range(2):para("________________________________________________________________________________")
pb()
para("ABSTRACT","Heading 1")
para("EcoSort AI classifies a waste-item image into one of twelve dataset categories and presents the result through a responsive web application. The system combines a MobileNetV2 transfer-learning model, a Flask inference service, and browser workflows for photo upload and camera capture. Camera users see a live preview, capture and review a still, and explicitly request prediction.")
para("The dataset contains 15,515 images. The model input is 224 × 224 pixels; the notebook uses an 80/20 training-validation split, class weighting, augmentation, early stopping, and fine-tuning. Saved metadata reports 94.68% validation accuracy and 0.172 validation loss. These are validation-split results, not independent test results or guarantees for new photographs. The project demonstrates a complete academic prototype joining data preparation, transfer learning, inference, and a browser interface.")
para("Keywords: waste classification, computer vision, MobileNetV2, TensorFlow, Flask, image recognition",italic=True)
pb()
para("ACKNOWLEDGEMENTS","Heading 1")
para("The project team expresses sincere gratitude to Dr. Brijesh Bakariya, Project Mentor and Head of the Department of Computer Science and Engineering, IK Gujral Punjab Technical University, Hoshiarpur Campus, for his academic guidance. The team also thanks Er. Satyam Sharma, Project Co-Mentor, for guidance and feedback. We acknowledge the Department and the authors of the open-source frameworks cited in this report. Dataset source and licence details should be confirmed before redistribution.")
para("Project team\nArshpreet Singh\nSarabjit Singh\nPooja")
pb()

# Main nine-chapter content using the sample's conventional outline.
chapters=[
 ("INTRODUCTION",[("1.1 Background","Waste sorting involves identifying materials such as paper, plastic, glass, metal, clothing, batteries, and organic waste. A person may be uncertain about an item, especially when it is damaged, mixed-material, or visually similar to another class. EcoSort AI explores whether a supervised image classifier can provide a useful first-pass suggestion through a simple browser workflow."),("1.2 Problem statement","A practical classifier must accept images reliably, preprocess them consistently, map model outputs to readable labels, and communicate uncertainty. Camera access also depends on user permission and a secure browser context. The project joins model training and the application workflow into one system."),("1.3 Existing and proposed system","Manual identification depends on prior knowledge, packaging, or separate local guidance. The proposed system uses a trained classifier to return likely categories from a user-selected photograph. Its output is informational and does not replace local disposal rules."),("1.4 Advantages and limitations","The application provides upload and still-camera workflows, ranked predictions, and a responsive interface. It performs single-image, single-label classification and does not identify multiple objects, determine material composition, or provide jurisdiction-specific recycling advice.")]),
 ("OBJECTIVES AND SCOPE",[("2.1 Objectives","Prepare a twelve-class image dataset; train a MobileNetV2 transfer-learning classifier; provide image upload and camera capture; return ranked category scores; document saved model performance and system limitations."),("2.2 Scope","The configured classes are battery, biological, brown glass, cardboard, clothes, green glass, metal, paper, plastic, shoes, trash, and white glass. The web application accepts supported image formats, preprocesses one image, and returns up to five candidate labels."),("2.3 Exclusions","The system does not perform object detection, segmentation, multi-item localization, hazardous-waste assessment, recycling-symbol reading, or local policy lookup. Predictions should be treated as suggestions.")]),
 ("FEASIBILITY STUDY",[("3.1 Operational feasibility","The interface guides users through photo upload or camera start, live preview, capture, review, and explicit prediction. Upload remains available when camera access is unavailable."),("3.2 Technical feasibility","The repository includes Python 3.11 runtime configuration, TensorFlow/Keras model files, Flask routes, browser UI, dataset, and metadata. Hosted camera access requires HTTPS; local development can use localhost."),("3.3 Economic feasibility","The project uses open-source frameworks. Hosting, storage, and compute costs depend on deployment scale and are not measured in the repository."),("3.4 Risk and feasibility boundary","The saved model and web code make the prototype demonstrable, but the repository does not contain a formal independent evaluation or documented live-host test. These should be recorded before claiming production readiness.")]),
 ("TOOLS AND TECHNOLOGIES",[("4.1 Software stack","The project uses Python for model and web code; TensorFlow/Keras for training and inference; MobileNetV2 for visual features; NumPy and Pillow for image preparation; Flask for routes; and HTML, CSS, and JavaScript for the browser interface."),("4.2 Runtime and deployment","Python is pinned to 3.11.9. Gunicorn serves the Flask application in the configured hosted environment. A manifest and service worker provide a progressive web application shell; inference still runs on the server."),("4.3 Technology roles","The following figure summarizes each tool's role in the project. The technology choices support a complete training-to-inference path rather than a standalone model demonstration.")]),
 ("DATA FLOW DIAGRAM",[("5.1 Context diagram — Level 0","The user supplies an image and receives ranked category suggestions. The EcoSort AI system contains validation, preprocessing, model inference, and result mapping."),("5.2 Detailed data flow — Level 1","The application validates the image and model choice, decodes the image as RGB, resizes it to 224 × 224, runs inference, maps scores to class labels, and returns the ranked result. Invalid inputs produce an error response."),("5.3 DFD symbols","External entity: user. Process: validation, preprocessing, prediction. Data store: saved Keras model and metadata. Data flow: image input and ranked prediction response.")]),
 ("SYSTEM DESIGN AND IMPLEMENTATION",[("6.1 Dataset","Saved metadata contains 15,515 examples distributed across twelve classes. Clothes is the largest category, with 5,325 images; brown glass is the smallest, with 607. This imbalance can affect class-level results."),("6.2 Preprocessing and model","Images are represented as RGB and resized to 224 × 224. The classifier uses ImageNet-pretrained MobileNetV2 with a task-specific twelve-class softmax head. Training includes augmentation, class weighting, early stopping, and fine-tuning of the final backbone layers."),("6.3 Modules and application workflow","Dataset module: reads images from class folders and assigns labels. Training module: splits data, applies augmentation and class weights, trains the classifier, and saves model metadata. Inference module: loads the saved Keras model and maps its output indices to category names. Upload module: checks file extension and selected model, decodes the image with Pillow, resizes it, and returns up to five scores. Camera module: accepts a captured still image as a data URL; prediction starts only after user confirmation. Interface module: presents upload/camera controls, status feedback, top category, and alternatives."),("6.4 Interfaces","The application offers a clear upload mode and a camera mode with preview, capture, retake, and prediction controls. Screenshots below document the project interface; camera preview graphics are explanatory captures, not proof of a live deployment.")]),
 ("RESULTS AND DISCUSSION",[("7.1 Dataset distribution","The class table lists the recorded dataset counts and proportions. The saved model metadata reports validation accuracy of 94.68% and validation loss of 0.1720."),("7.2 Interpretation","The metrics summarize the selected validation split. They do not establish performance on an independent test set, new cameras, mixed waste scenes, or every class equally. No per-class precision/recall or confusion matrix is recorded in the repository."),("7.3 Verification status","The repository contains code and saved model artifacts, but no formal execution log for all application tests. Recommended checks include supported/unsupported image upload, camera permission, capture and retake, model selection, narrow-screen layout, and hosted smoke testing.")]),
 ("CONCLUSION AND FUTURE SCOPE",[("8.1 Conclusion","EcoSort AI demonstrates an end-to-end twelve-class waste-image classifier integrated with a Flask web application. It combines a prepared dataset, transfer learning, server-side inference, photo upload, and a user-triggered camera capture flow. The saved validation metrics are encouraging for the project split but need independent evaluation."),("8.2 Future scope","Document dataset provenance and licensing; curate a held-out test set; publish per-class precision, recall, F1, and confusion matrix; improve minority-class data; capture live interface screenshots; test on target phones; and connect predictions to verified local disposal guidance."),("8.3 Limitations","Predictions may be wrong for ambiguous, mixed, poorly lit, or unfamiliar images. The model is not a hazardous-waste authority or a substitute for local waste-management rules.")]),
 ("REFERENCES AND APPENDIX",[("9.1 References","Project-specific facts are based on the project repository, saved metadata, application source, notebook, and interface files. Technical references follow."),("9.2 Appendix","A repository map and submission checklist help verify the delivered project files and academic front matter.")])]

# Expand the concise outline into the complete academic chapter structure from the supplied sample.
chapters = [
 ("INTRODUCTION",[
  ("1.1 Background","Waste sorting requires identifying paper, plastic, glass, metal, clothing, batteries, and organic materials. A person may be unsure about an item when it is damaged, mixed-material, unfamiliar, or visually similar to another category. Image classification can provide a quick first-pass suggestion, provided the image is clear and the limits of the model are understood."),
  ("1.2 Problem statement","Manual identification depends on prior knowledge, packaging labels, or searching for local disposal instructions. A computer-vision aid must do more than calculate a class: it must accept images safely, preprocess them consistently, map scores to the correct labels, and communicate uncertainty. Camera use also depends on browser permission and secure-origin rules."),
  ("1.3 Objectives","The project prepares a twelve-class labeled image pipeline; trains a MobileNetV2 transfer-learning model; provides browser-based photo upload and camera capture; returns ranked category scores; and documents the implemented system, measured validation result, limitations, and further evaluation needs."),
  ("1.4 Scope","The model performs single-image, single-label classification over twelve configured waste categories. The Flask application supports image upload and explicitly captured camera stills. The output is an informational category suggestion, not a local disposal instruction."),
  ("1.5 Existing system","Without the prototype, users may rely on visual judgment, package labels, or external waste information. These methods can be inconvenient when the user does not recognize a material or the item combines several materials."),
  ("1.6 Proposed system","EcoSort AI combines labeled images, a trained neural network, an inference service, and a responsive interface. The application validates the request, prepares the image, obtains model scores, and shows the highest-ranked labels. Camera users preview and review a captured still before requesting prediction."),
  ("1.7 Advantages and limitations","The system makes a trained classifier accessible through a browser and exposes more than one candidate label. It does not perform object detection, segmentation, material-composition analysis, contamination grading, barcode reading, or jurisdiction-specific guidance. A model suggestion should not determine safe handling of batteries or hazardous waste.")]),
 ("FEASIBILITY STUDY",[
  ("2.1 Operational feasibility","The two user paths are straightforward: select a supported image and submit it, or grant camera permission, capture a still, review it, and request prediction. The upload path remains available if a camera is absent or denied. The app is suitable for coursework demonstration and evaluation."),
  ("2.2 Technical feasibility","The repository contains the source notebook, trained Keras models, metadata, Flask application, HTML/CSS/JavaScript interface, dataset folders, and deployment configuration. The runtime is pinned to Python 3.11.9. Camera access on hosted pages requires HTTPS; localhost is suitable for local development."),
  ("2.3 Economic feasibility","The framework stack is open source. The repository does not contain a cost study. Actual hosting, storage, and inference costs depend on the selected service, traffic, worker count, and compute size."),
  ("2.4 Schedule and resource feasibility","The repository has completed artifacts for data, model, and application integration. No time-sheet or resource ledger is supplied, so this report does not claim exact development hours or project cost. A final demonstration should record device, runtime, model version, and test outcomes."),
  ("2.5 Risk assessment","Primary risks include class imbalance, unknown dataset provenance, incorrect classification, insecure deployment configuration, camera permission failures, model-loading or hosting limits, and user overreliance. The report separates saved validation measures from independent field performance.")]),
 ("TOOLS AND TECHNOLOGIES USED",[
  ("3.1 Development environment","The repository uses Python 3.11.9 for the training notebook and Flask server. The configured deployment uses Gunicorn as the WSGI server. Visual development and source files are stored in the project repository."),
  ("3.2 Python","Python coordinates data loading, model training, image decoding, prediction, and HTTP request handling. It provides one language across the notebook and server-side application."),
  ("3.3 TensorFlow and Keras","TensorFlow/Keras provides image loading, model construction, optimization, validation, serialization, and inference. The saved .keras files are loaded by the Flask application."),
  ("3.4 MobileNetV2","MobileNetV2 is used as a pretrained feature-extraction backbone. The original classification top is removed and a project-specific twelve-class head is added. It is an efficient architecture, but the current web application performs inference on the server rather than locally on the phone."),
  ("3.5 NumPy and Pillow","NumPy represents image data as arrays for model input. Pillow decodes uploaded and camera images, converts them to RGB, and resizes them to the configured dimensions."),
  ("3.6 Flask","Flask serves the main browser page, handles multipart image upload, validates model selection, and exposes the camera prediction endpoint. The application limits request size and checks supported file extensions."),
  ("3.7 HTML, CSS, and JavaScript","HTML defines upload and camera controls; CSS provides the responsive visual layout; JavaScript switches modes, requests camera permission, captures still frames, and posts the selected image to Flask."),
  ("3.8 Progressive Web App and deployment tools","A manifest and service worker provide web-app metadata and application-shell caching. The trained model is not cached for local inference; prediction still requires the server. Render configuration, runtime pins, and Gunicorn startup settings document deployment intent."),
  ("3.9 Technology summary","The figure summarizes the software tools and the responsibility each has in this implementation.")]),
 ("DATA FLOW DIAGRAM",[
  ("4.1 DFD overview","A data-flow diagram represents the system boundary, external entities, processes, data stores, and movement of information. In this project the user submits a photograph; the application validates and prepares it; the model produces class scores; and the interface returns ranked labels."),
  ("4.2 DFD Level 0","At context level the user is the external entity and EcoSort AI is the classification process. The user sends an image and receives category suggestions. The drawing shows the main inputs and outputs without internal implementation details."),
  ("4.3 DFD Level 1","The detailed diagram separates request validation, image preparation, model prediction, and output mapping. The saved model and metadata act as data stores. An invalid image or model selection returns an error rather than a classification."),
  ("4.4 Data dictionary","Image input is a JPG, JPEG, PNG, BMP, WEBP upload or a captured JPEG still. The prepared tensor is RGB at 224 × 224 pixels. The model returns scores for twelve ordered class labels. The response contains up to five labels and their scores."),
  ("4.5 DFD symbols","External entity means a user outside the application boundary. A process transforms input data. A data store represents the saved model and metadata. An arrow represents image, model information, error, or result transfer.")]),
 ("SYSTEM DESIGN AND IMPLEMENTATION",[
  ("5.1 System overview","The system has a browser interface, Flask request layer, image-preparation helper, cached Keras model loader, saved model artifacts, and class metadata. The layers separate user interaction, validation, preprocessing, inference, and display."),
  ("5.2 Dataset description","The repository records 15,515 images in twelve category folders. The table presents each class count and share. The clothes category is the largest; brown glass is the smallest. This imbalance is one reason the training notebook uses class weights."),
  ("5.3 Data preprocessing","Training images are loaded from class folders, resized to 224 × 224, divided into training and validation subsets using an 80/20 setting and seed 42, batched, and prefetched. Training augmentation includes horizontal flip, rotation, zoom, and contrast variation. Inference converts each user image to RGB and resizes it to the metadata dimensions."),
  ("5.4 Model methodology","The model begins with ImageNet-pretrained MobileNetV2. Its original top is omitted; global average pooling, dropout, and a twelve-unit softmax head are attached. The training sequence first trains the classifier head with the backbone frozen, then fine-tunes the final thirty backbone layers at a lower learning rate. Early stopping and learning-rate reduction monitor validation behavior."),
  ("5.5 Application modules","Dataset module supplies labeled training examples. Training module builds the model and saves weights and metadata. Inference module loads the selected .keras artifact and maps output indices using class_names. Upload module validates file extension and model selection. Camera module accepts a still image data URL. User-interface module presents the preview and prediction results."),
  ("5.6 Upload workflow","The user chooses a model and image. The server checks that the file exists, verifies the extension, confirms the model is available, opens the image with Pillow, converts it to RGB, resizes it, runs inference, sorts the output, and renders the top category with alternatives."),
  ("5.7 Camera workflow","The browser requests video permission and displays a preview. The user captures a frame, reviews the still, and presses the prediction control. JavaScript sends a JPEG data URL to the camera endpoint. The camera is not classified continuously; permission and browser compatibility affect operation."),
  ("5.8 User interface and screenshots","The interface includes upload and camera modes, model selection, file input, preview, status feedback, capture, retake, and prediction controls. The images below document the interface design and dataset examples. Interface previews should not be treated as proof of a live hosted service."),
  ("5.9 Data and model artifacts","The models directory contains trained Keras files, metadata with class names, input size, class counts and validation measures, and a training history plot. The metadata class order must remain consistent with the model output units."),
  ("5.10 Deployment configuration","The repository pins Python 3.11.9, installs packages from requirements.txt, and configures Gunicorn to serve app:app. The service accepts requests up to the configured limit. Deployment files are configuration evidence; current production status and hosted performance are not asserted here.")]),
 ("USER EXPERIENCE AND OPERATION",[
  ("6.1 Photo upload","Choose a single, clear image in a supported format. Select an available model, submit the image, then review the leading category and ranked alternatives. The score is a model output and not a guarantee of correctness."),
  ("6.2 Camera capture","Use HTTPS or localhost, allow camera access, frame one object, capture a still, inspect it, and request prediction. Retake the image if it is blurry or the item is too small. Upload is an alternative when camera access is unavailable."),
  ("6.3 Error handling","The application reports missing image, unsupported format, invalid model, unreadable image, oversized camera data, and prediction errors. The camera endpoint checks the model allow-list and image data URL before inference."),
  ("6.4 Accessibility and mobile use","The interface uses browser form controls, labels, status messaging, responsive styling, and touch-friendly controls. Camera access and PWA installation vary by platform; these behaviors should be checked on the actual target devices."),
  ("6.5 Operational safeguards","The Flask configuration limits request size; filename extension and model selection are checked. The repository does not implement user accounts or a database for image history. A deployed service should explain request and log handling and avoid retaining user images without a stated purpose.")]),
 ("DATABASE AND MODEL RESULTS",[
  ("7.1 Data storage","The application does not use a relational database. Its persistent inference artifacts are model files and a JSON metadata file. Training data remains organized in labeled class directories. The browser sends images in a request; no project database stores prediction history."),
  ("7.2 Class distribution","The class distribution table records all twelve folders and the total. Clothes represents about one third of the examples, while brown and green glass have fewer than 650 examples each. A high aggregate score can obscure weaker minority-class performance."),
  ("7.3 Saved model results","The repository metadata records 94.68% validation accuracy and 0.1720 validation loss. These values belong to the project validation split and selected saved model. They are not independent test results, and no class-wise precision, recall, F1, or confusion matrix is supplied."),
  ("7.4 Training history","The training-history chart illustrates saved accuracy and loss curves. These curves help review model development but do not replace an untouched test set or deployment evaluation."),
  ("7.5 Verification plan","The repository does not contain a formal test execution report. Recommended checks include valid and invalid uploads, model selection, camera permission, still capture and retake, top-five response, narrow-screen layout, hosted startup, and consistent label ordering. These are planned checks, not claimed completed passes."),
  ("7.6 Result interpretation","The validation result is a useful coursework benchmark for the saved split. New photographs may differ in lighting, background, viewpoint, composition, camera, or class balance. Users should use local waste guidance for the final disposal decision.")]),
 ("CONCLUSION AND FUTURE SCOPE",[
  ("8.1 Conclusion","EcoSort AI links a twelve-class image dataset and MobileNetV2 classifier to a Flask application with upload and capture-review-predict camera flows. The saved metadata reports 94.68% validation accuracy and 0.172 validation loss on the project split. The system demonstrates an end-to-end academic prototype while retaining a clear boundary between validation and real-world performance."),
  ("8.2 Limitations","The application predicts one configured label for one image. It may fail on ambiguous, mixed-material, unfamiliar, poorly lit, or multi-object scenes. Dataset provenance and licensing are not fully documented in the project snapshot. Independent, per-class test results are not recorded."),
  ("8.3 Future scope","Document dataset source and licence; audit labels and near duplicates; collect more examples for minority classes; set aside a held-out test set; report class-wise metrics and confusion matrix; improve calibrated low-confidence handling; capture genuine live deployment screenshots; test across phones and browsers; and add verified local collection guidance."),
  ("8.4 Further evaluation","A follow-up evaluation should predefine the test set and capture conditions, report sample counts per category, preserve error examples, measure request latency separately from model accuracy, and record browser/device and deployment versions.")]),
 ("REFERENCES AND APPENDICES",[
  ("9.1 References","[1] M. Sandler, A. Howard, M. Zhu, A. Zhmoginov, and L.-C. Chen, MobileNetV2: Inverted Residuals and Linear Bottlenecks, CVPR, 2018.\n[2] TensorFlow, Keras MobileNetV2 API documentation.\n[3] Pallets Projects, Flask documentation.\n[4] EcoSort AI repository: application source, training notebook, metadata, static and template files, model artifacts, runtime and deployment configuration."),
  ("9.2 Repository map","app.py: Flask routes and inference. models/: Keras models, metadata and history. garbage_classification Dataset/: labeled image folders. garbage_classification_model.ipynb: training notebook. templates/ and static/: browser interface, styles, camera controls and PWA shell."),
  ("9.3 Submission checklist","Confirm student names, university roll numbers, programme, session, approved title, mentor spellings, signatures, dataset source and licence. Verify final page numbering and contents after Word updates fields. The group message specifies one hard-bound copy and a 12 November 2026 submission date."),
  ("9.4 Team contribution record","The repository does not contain a verified contribution log. Team members should review and complete their actual responsibilities before submission; this report does not assign individual credit without confirmation.")])]

for ci,(name,sections) in enumerate(chapters,1):
    pb(); para(f"CHAPTER {ci}: {name}","Heading 1")
    for st,body in sections:
        para(st,"Heading 2"); para(body)
        if ci==3 and st.startswith("3.9"): pic(tech_path,5.4,"Figure 3.1. Tools and technologies used in EcoSort AI.")
        if ci==4 and st.startswith("4.2"): pic(BUILD/"dfd_level0.png",5.2,"Figure 4.1. Context-level data-flow diagram.")
        if ci==4 and st.startswith("4.3"): pic(BUILD/"dfd_level1.png",5.2,"Figure 4.2. Detailed data-flow diagram.")
        if ci==5 and st.startswith("5.2"):
            tbl([["Waste category","Images","Share"]]+[[n.replace("-"," ").title(),str(meta['class_counts'][n]),f"{meta['class_counts'][n]/sum(meta['class_counts'].values())*100:.1f}%"] for n in meta["class_names"]]+[["Total","15,515","100.0%"]],[2.7,1.0,1.0],9)
        if ci==5 and st.startswith("5.4"): pic(ROOT/"models"/"training_history.png",5.0,"Figure 5.1. Saved training and validation history.")
        if ci==5 and st.startswith("5.8"):
            pic(BUILD/"upload_screen_detail.png",4.8,"Figure 5.2. Upload interface preview.")
            pic(BUILD/"camera_screen_detail.png",4.8,"Figure 5.3. Camera workflow preview.")
            pic(BUILD/"dataset_sample_gallery.png",5.4,"Figure 5.4. Representative images from the twelve dataset categories.")
        if ci==7 and st.startswith("7.3"):
            tbl([["Metric","Recorded value","Interpretation"],["Validation accuracy",f"{meta['validation_accuracy']*100:.2f}%","Saved validation split."],["Validation loss",f"{meta['validation_loss']:.4f}","Saved validation split."],["Output categories","12","Configured model labels."],["Dataset images","15,515","Sum of metadata class counts."]],[1.8,1.4,2.6],9)
    if ci==4:
        tbl([["DFD element","Meaning in EcoSort AI"],["External entity","User submitting an image and viewing results."],["Processes","Validate request, prepare image, predict, map labels."],["Data store","Saved Keras model and class metadata."],["Data flows","Image, validation response, and ranked scores."]],[1.6,4.1],9)

# Add the contents and figure lists in their correct front-matter position.
# They are inserted immediately before the first chapter heading in document order.
first_chapter = next(p for p in doc.paragraphs if p.text.startswith("CHAPTER 1:"))
insert_before = first_chapter._p
def insert_para_before(text,style=None):
    newp=OxmlElement("w:p"); insert_before.addprevious(newp)
    from docx.text.paragraph import Paragraph as DocxParagraph
    pp=DocxParagraph(newp,first_chapter._parent)
    if style: pp.style=style
    pp.add_run(text); return pp
insert_para_before("TABLE OF CONTENTS","Heading 1")
insert_para_before("Chapter page numbers should be refreshed after final pagination.")
insert_para_before("LIST OF FIGURES","Heading 1")
insert_para_before("Figure list: tools and technologies; Level-0 and Level-1 DFDs; training history; upload interface; camera workflow; dataset sample gallery.")

doc.core_properties.title="EcoSort AI — AI-Powered Waste Classification Project Report"
doc.core_properties.author="Project team"
doc.core_properties.subject="B.Tech Computer Science Engineering project report"
doc.save(DOCX_OUT)

# Render the DOCX when office is available; PDF is built from the same narrative with ReportLab below.
styles=getSampleStyleSheet()
styles.add(ParagraphStyle(name="RBody",parent=styles["BodyText"],fontName="Times-Roman",fontSize=9.6,leading=13.2,alignment=TA_JUSTIFY,spaceAfter=6))
styles.add(ParagraphStyle(name="RHead",parent=styles["Heading1"],fontName="Helvetica-Bold",fontSize=17,leading=21,textColor=colors.HexColor("#103E2E"),spaceAfter=7))
styles.add(ParagraphStyle(name="RSub",parent=styles["Heading2"],fontName="Helvetica-Bold",fontSize=11.4,leading=14,textColor=colors.HexColor("#176547"),spaceBefore=6,spaceAfter=3))
styles.add(ParagraphStyle(name="RCell",parent=styles["BodyText"],fontName="Helvetica",fontSize=7.7,leading=9.5,spaceAfter=0))
styles.add(ParagraphStyle(name="RSmall",parent=styles["BodyText"],fontName="Helvetica",fontSize=8,leading=10,textColor=colors.HexColor("#5B6B62"),spaceAfter=4))
styles.add(ParagraphStyle(name="RCover",parent=styles["Title"],fontName="Helvetica-Bold",fontSize=24,leading=28,textColor=colors.HexColor("#103E2E"),alignment=TA_CENTER,spaceAfter=7))

story=[]
def Q(text,sty="RBody"):
    safe=str(text).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    return Paragraph(safe,styles[sty])
def Qtable(rows,widths):
    data=[[Paragraph(str(c).replace("&","&amp;"),styles["RCell"]) for c in r] for r in rows]
    t=Table(data,colWidths=widths,repeatRows=1,hAlign="LEFT")
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#176547")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.35,colors.HexColor("#D7E2DA")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F2F7F3")]),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
    return t
def Qimg(path,w=150*mm,h=None,caption=None):
    im=RLImage(str(path),width=w,height=h)
    im._restrictSize(w,h or 85*mm)
    im.hAlign="CENTER"
    return [im,Q(caption,"RSmall")] if caption else [im]
def Qheader(canv,d):
    canv.saveState()
    if d.page>1:
        canv.setFont("Helvetica-Bold",8);canv.setFillColor(colors.HexColor("#176547"));canv.drawString(18*mm,A4[1]-12*mm,"ECOSORT AI  /  PROJECT REPORT")
        canv.setFont("Helvetica",8);canv.setFillColor(colors.HexColor("#5B6B62"));canv.drawRightString(A4[0]-18*mm,A4[1]-12*mm,"Garbage Classification with Deep Learning")
        canv.setStrokeColor(colors.HexColor("#D7E2DA"));canv.line(18*mm,A4[1]-15*mm,A4[0]-18*mm,A4[1]-15*mm)
        canv.drawRightString(A4[0]-18*mm,10*mm,str(d.page))
    canv.restoreState()

# cover
story += [Spacer(1,3*mm),*Qimg(ROOT/"static"/"garbage-ai-logo.jpg",22*mm,h=22*mm),Q("IK GUJRAL PUNJAB TECHNICAL UNIVERSITY, HOSHIARPUR CAMPUS","RSmall"),Q("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING","RSmall"),Spacer(1,5*mm),Q("PROJECT REPORT","RSmall"),Q("EcoSort AI","RCover"),Q("AI-Powered Waste Classification","RHead"),Q("Garbage Classification with Deep Learning","RSmall"),Spacer(1,4*mm),Q("Submitted in partial fulfilment of the requirements for the award of the degree of Bachelor of Technology in Computer Science Engineering."),Spacer(1,4*mm),Qtable([["SUBMITTED BY","SUBMITTED TO"],["Arshpreet Singh<br/>Sarabjit Singh<br/>Pooja<br/>Roll Nos.: 2323174, 2323218, 2323200","Dr. Brijesh Bakariya<br/>Project Mentor and HOD, CSE<br/>Er. Satyam Sharma<br/>Project Co-Mentor"]],[78*mm,78*mm]),Spacer(1,4*mm),Q("Semester: 7th  |  Academic session: 2027  |  Submission date: ____________________","RSmall"),Q("Hard-bound project report submission","RSmall"),PageBreak()]
story += [Q("DECLARATION","RHead"),Q("We, the undersigned project team, declare that this report describes EcoSort AI project work carried out by our team. Project implementation details and results are based on the project repository and saved model metadata. External technical sources are cited in the references. This declaration should be signed by all listed students after review."),Qtable([["Name","Signature","Date"],["Arshpreet Singh","____________________________","____________"],["Sarabjit Singh","____________________________","____________"],["Pooja","____________________________","____________"]],[42*mm,75*mm,35*mm]),Spacer(1,6*mm),Q("Department: Computer Science and Engineering<br/>Institution: IK Gujral Punjab Technical University, Hoshiarpur Campus<br/>Semester: 7th; Academic session: 2027"),PageBreak()]
story += [Q("CERTIFICATE","RHead"),Q("This is to certify that the project entitled EcoSort AI: AI-Powered Waste Classification has been carried out by Arshpreet Singh, Sarabjit Singh, and Pooja in the Department of Computer Science and Engineering, IK Gujral Punjab Technical University, Hoshiarpur Campus, in partial fulfilment of the requirements for the Bachelor of Technology in Computer Science Engineering under the guidance of Dr. Brijesh Bakariya, Project Mentor and Head of Department, and Er. Satyam Sharma, Project Co-Mentor. Certification is subject to institutional review and signatures."),Spacer(1,15*mm),Qtable([["Project Mentor / HOD","Project Co-Mentor","External Examiner"],["Dr. Brijesh Bakariya","Er. Satyam Sharma","Name: ______________"],["Signature: ____________","Signature: ____________","Signature: ____________"],["Date: _________________","Date: _________________","Date: _________________"]],[50*mm,50*mm,50*mm]),Spacer(1,10*mm),Q("Institution seal / remarks<br/><br/>________________________________________________________________________________<br/><br/>________________________________________________________________________________","RSmall"),PageBreak()]
story += [Q("ABSTRACT","RHead"),Q("EcoSort AI classifies a waste-item image into one of twelve dataset categories and presents the result through a responsive web application. It combines a MobileNetV2 transfer-learning model, a Flask inference service, and browser workflows for photo upload and camera capture. Camera users see a live preview, capture and review a still, then explicitly request prediction."),Q("The dataset contains 15,515 images. The model input is 224 × 224 pixels; training uses an 80/20 training-validation split, class weighting, augmentation, early stopping, and fine-tuning. Saved metadata reports 94.68% validation accuracy and 0.172 validation loss. These are validation-split results, not independent test results or guarantees for new photographs."),Q("Keywords: waste classification, computer vision, MobileNetV2, TensorFlow, Flask, image recognition.","RSmall"),PageBreak()]
story += [Q("ACKNOWLEDGEMENTS","RHead"),Q("The team thanks Dr. Brijesh Bakariya, Project Mentor and Head of the Department of Computer Science and Engineering, and Er. Satyam Sharma, Project Co-Mentor, for their guidance. The team acknowledges the Department and the authors of open-source frameworks cited in this report. Dataset source and licence details should be verified before redistribution."),Q("Project team: Arshpreet Singh, Sarabjit Singh, and Pooja."),PageBreak()]

# Contents and figure lists use chapter names; final page numbers are refreshed from rendered PDF below.
story += [Q("TABLE OF CONTENTS","RHead"),Qtable([["CHAPTER NO.","TOPIC","PAGE NO."]]+[[f"CHAPTER {i}",nm,""] for i,(nm,_) in enumerate(chapters,1)]+[["REFERENCES / APPENDIX","References, repository map, and submission checklist",""]],[30*mm,104*mm,18*mm]),PageBreak()]
story += [Q("LIST OF FIGURES","RHead"),Qtable([["FIGURE NO.","TOPIC","PAGE NO."],["Figure 3.1","Tools and technologies",""],["Figure 4.1","Context-level DFD",""],["Figure 4.2","Detailed DFD",""],["Figure 5.1","Training and validation history",""],["Figure 5.2","Upload interface",""],["Figure 5.3","Camera workflow",""],["Figure 5.4","Dataset sample gallery",""]],[30*mm,104*mm,18*mm]),PageBreak()]

for ci,(nm,sections) in enumerate(chapters,1):
    story.append(Q(f"CHAPTER {ci}: {nm}","RHead"))
    for st,bodytext in sections:
        story += [Q(st,"RSub"),Q(bodytext)]
        if ci==3 and st.startswith("3.9"):
            story += Qimg(tech_path,145*mm,h=78*mm,caption="Figure 3.1. Tools and technologies used in EcoSort AI.")
        if ci==4 and st.startswith("4.2"):
            story += Qimg(BUILD/"dfd_level0.png",145*mm,h=76*mm,caption="Figure 5.1. Level-0 context DFD.")
        if ci==4 and st.startswith("4.3"):
            story += Qimg(BUILD/"dfd_level1.png",145*mm,h=76*mm,caption="Figure 5.2. Level-1 DFD.")
        if ci==5 and st.startswith("5.2"):
            story.append(Qtable([["Waste category","Images","Share"]]+[[n.replace("-"," ").title(),str(meta["class_counts"][n]),f"{meta['class_counts'][n]/15515*100:.1f}%"] for n in meta["class_names"]]+[["Total","15,515","100.0%"]],[70*mm,30*mm,30*mm]))
        if ci==5 and st.startswith("5.4"):
            story += Qimg(ROOT/"models"/"training_history.png",135*mm,h=75*mm,caption="Figure 6.1. Saved training and validation history.")
        if ci==5 and st.startswith("5.8"):
            story += Qimg(BUILD/"upload_screen_detail.png",125*mm,h=55*mm,caption="Figure 6.2. Upload-interface preview; implementation reference.")
            story += Qimg(BUILD/"camera_screen_detail.png",125*mm,h=55*mm,caption="Figure 6.3. Camera workflow preview; browser permissions vary.")
            story += Qimg(BUILD/"dataset_sample_gallery.png",135*mm,h=68*mm,caption="Figure 6.4. Representative dataset examples across classes.")
        if ci==7 and st.startswith("7.3"):
            story.append(Qtable([["Metric","Recorded value","Interpretation"],["Validation accuracy",f"{meta['validation_accuracy']*100:.2f}%","Validation split result."],["Validation loss",f"{meta['validation_loss']:.4f}","Validation split result."],["Categories","12","Configured model outputs."],["Dataset","15,515","Recorded class-count sum."]],[44*mm,35*mm,51*mm]))
    if ci==4:
        story.append(Qtable([["DFD symbol","Meaning"],["External entity","User"],["Process","Validation, preprocessing, prediction"],["Data store","Model file and class metadata"],["Data flow","Image input and ranked labels"]],[45*mm,95*mm]))
    if ci==9:
        story.append(Q("Selected references","RSub"))
        for ref in ["[1] M. Sandler et al., “MobileNetV2: Inverted Residuals and Linear Bottlenecks,” CVPR, 2018.","[2] TensorFlow, Keras MobileNetV2 API documentation.","[3] Pallets Projects, Flask documentation.","[4] EcoSort AI repository source: app.py, notebook, model metadata, templates, static files, and deployment configuration."]:
            story.append(Q(ref))
        story.append(Q("Repository map","RSub"))
        story.append(Qtable([["Path","Purpose"],["app.py","Flask routes and inference."],["models/","Saved models and metadata."],["garbage_classification Dataset/","Labeled class image folders."],["templates/ and static/","Browser UI, styles, camera controls, and PWA shell."],["garbage_classification_model.ipynb","Training and validation notebook."]],[55*mm,85*mm]))
        story.append(Q("Before printing, confirm names, roll numbers, academic session, signatures, and dataset source/licence with the team and department. The group instruction specifies hard binding and submission by 12 November 2026."))
    story.append(PageBreak())

pdf=SimpleDocTemplate(str(PDF_OUT),pagesize=A4,leftMargin=19*mm,rightMargin=19*mm,topMargin=18*mm,bottomMargin=17*mm,title="EcoSort AI Project Report",author="EcoSort AI Project Team")
pdf.build(story,onFirstPage=Qheader,onLaterPages=Qheader)
print(DOCX_OUT); print(PDF_OUT)
