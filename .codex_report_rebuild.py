from pathlib import Path
import json
import re
import runpy

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image, KeepTogether
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

ROOT = Path(__file__).resolve().parent
REPORTS = ROOT / "Report"
SOURCE = ROOT / "project_archive" / "presentation-build" / "create_full_report.py"
PDF_SOURCE = REPORTS / "EcoSort_AI_Professional_Project_Report_50_Page_Edition.pdf"
DOCX_OUT = REPORTS / "EcoSort_AI_Project_Report_University_Format_Concise.docx"
PDF_OUT = REPORTS / "EcoSort_AI_Project_Report_Final.pdf"
BUILD = ROOT / ".codex_report_build"
BUILD.mkdir(exist_ok=True)

src = SOURCE.read_text(encoding="utf-8")
src = src.replace('ROOT = Path(r"C:\\Users\\HP\\Desktop\\data science\\7TH-SEMESTER-PROJECT-")', f'ROOT = Path({str(ROOT)!r})')
src = src.replace('BUILD = ROOT / ".ppt-team-build"', f'BUILD = Path({str(BUILD)!r})')
src = src.replace('OUT = ROOT / "Report" / "EcoSort_AI_Professional_Project_Report_50_Page_Edition.pdf"', f'OUT = Path({str(BUILD / "source_full.pdf")!r})')
# Temporary assets used by the original build were not committed; use the saved report page renders instead.
src = src.replace('preview_path = BUILD / "slide-08.png"', 'preview_path = ROOT / "project_archive" / "presentation-build" / "report-pages" / "page-08.png"')
src = src.replace('preview.crop((72, 155, 620, 610)).resize((1096, 910)).save(BUILD / "upload_screen_detail.png")', 'preview.resize((1096, 910)).save(BUILD / "upload_screen_detail.png")')
src = src.replace('preview.crop((660, 155, 1208, 610)).resize((1096, 910)).save(BUILD / "camera_screen_detail.png")', 'preview.resize((1096, 910)).save(BUILD / "camera_screen_detail.png")')
src = src.replace('ROOT / "garbage_classification" / class_name', 'ROOT / "garbage_classification Dataset" / class_name')
src = src.replace('add_image(BUILD/"slide-08.png", 170*mm, 115*mm)', 'add_image(BUILD/"upload_screen_detail.png", 170*mm, 115*mm)')
src = src.replace('add_image(BUILD/"slide-09.png", 170*mm, 125*mm)', 'add_image(BUILD/"camera_screen_detail.png", 170*mm, 125*mm)')
temp_source = BUILD / "generate_source.py"
temp_source.write_text(src, encoding="utf-8")
runpy.run_path(str(temp_source), run_name="__main__")

source_pdf = PdfReader(str(BUILD / "source_full.pdf"))
page_text = [(p.extract_text() or "").replace("\x00", "") for p in source_pdf.pages]

# Verify report facts directly from the repository snapshot.
meta = json.loads((ROOT / "models" / "metadata.json").read_text(encoding="utf-8"))
assert sum(meta["class_counts"].values()) == 15515
assert len(meta["class_names"]) == 12
assert abs(meta["validation_accuracy"] - 0.9468256235122681) < 1e-9

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
sec.top_margin, sec.bottom_margin = Inches(.78), Inches(.7)
sec.left_margin, sec.right_margin = Inches(.85), Inches(.78)

GREEN = "176547"
DEEP = "103E2E"
TEAL = "1686A0"
PALE = "F2F7F3"
INK = "1C2922"
MUTED = "5B6B62"
LINE = "D7E2DA"

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Times New Roman"
normal.font.size = Pt(11)
normal.font.color.rgb = RGBColor.from_string(INK)
normal.paragraph_format.line_spacing = 1.2
normal.paragraph_format.space_after = Pt(6)
for stylename, size, color in [("Title", 25, DEEP), ("Heading 1", 18, DEEP), ("Heading 2", 13, GREEN), ("Heading 3", 11, TEAL)]:
    s = styles[stylename]
    s.font.name = "Times New Roman"
    s.font.size = Pt(size)
    s.font.bold = True
    s.font.color.rgb = RGBColor.from_string(color)
    s.paragraph_format.keep_with_next = True
    s.paragraph_format.space_before = Pt(10 if stylename != "Title" else 0)
    s.paragraph_format.space_after = Pt(5)

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)

def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge in kwargs:
            tag = "w:" + edge
            element = borders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                borders.append(element)
            for key in ["val", "sz", "space", "color"]:
                if key in kwargs[edge]:
                    element.set(qn("w:" + key), str(kwargs[edge][key]))

def add_field(paragraph, instr):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    text = OxmlElement("w:instrText"); text.set(qn("xml:space"), "preserve"); text.text = instr
    separate = OxmlElement("w:fldChar"); separate.set(qn("w:fldCharType"), "separate")
    result = OxmlElement("w:t"); result.text = "1"
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, text, separate, result, end])

def set_header_footer(section):
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = header.add_run("ECOSORT AI   |   PROJECT REPORT")
    r.bold = True; r.font.size = Pt(8); r.font.color.rgb = RGBColor.from_string(GREEN)
    r2 = header.add_run("                                      Garbage Classification with Deep Learning")
    r2.font.size = Pt(8); r2.font.color.rgb = RGBColor.from_string(MUTED)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    rr = footer.add_run("Page "); rr.font.size = Pt(8)
    add_field(footer, "PAGE")

set_header_footer(sec)

def p(text="", style=None, align=None, bold=False, italic=False):
    para = doc.add_paragraph(style=style)
    if align is not None: para.alignment = align
    run = para.add_run(text)
    run.bold = bold; run.italic = italic
    return para

def page_break(): doc.add_page_break()

def add_table(rows, widths=None, font=9):
    table = doc.add_table(rows=0, cols=len(rows[0]))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    if widths:
        for row in table.rows:
            for i, w in enumerate(widths): row.cells[i].width = Inches(w)
    for ri, row_data in enumerate(rows):
        cells = table.add_row().cells
        if ri == 0: set_repeat_table_header(table.rows[-1])
        for ci, value in enumerate(row_data):
            cell = cells[ci]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            if ri == 0: shade(cell, GREEN)
            elif ri % 2 == 0: shade(cell, PALE)
            cell.text = ""
            para = cell.paragraphs[0]
            para.paragraph_format.space_after = Pt(2)
            run = para.add_run(str(value)); run.font.name = "Times New Roman"; run.font.size = Pt(font)
            if ri == 0:
                run.bold = True; run.font.color.rgb = RGBColor(255,255,255)
            else: run.font.color.rgb = RGBColor.from_string(INK)
            set_cell_border(cell, top={"val":"single","sz":"4","color":LINE}, bottom={"val":"single","sz":"4","color":LINE}, left={"val":"single","sz":"4","color":LINE}, right={"val":"single","sz":"4","color":LINE})
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table

def add_image(path, width=6.0):
    para = doc.add_paragraph(); para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.add_run().add_picture(str(path), width=Inches(width))

def title(text): p(text, "Heading 1")
def subsection(text): p(text, "Heading 2")

# Cover follows the sample's university, award, submitted by/to, title layout.
for _ in range(2): doc.add_paragraph()
logo = ROOT / "static" / "garbage-ai-logo.jpg"
add_image(logo, 1.45)
p("IK GUJRAL PUNJAB TECHNICAL UNIVERSITY, HOSHIARPUR CAMPUS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
p("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
for _ in range(2): doc.add_paragraph()
p("PROJECT REPORT", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
p("EcoSort AI", "Title", WD_ALIGN_PARAGRAPH.CENTER)
p("AI-Powered Waste Classification", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
p("Garbage Classification with Deep Learning", align=WD_ALIGN_PARAGRAPH.CENTER)
for _ in range(2): doc.add_paragraph()
p("Submitted in partial fulfilment of the requirements for the award of the degree of", align=WD_ALIGN_PARAGRAPH.CENTER)
p("Bachelor of Technology in Computer Science Engineering", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
for _ in range(2): doc.add_paragraph()
cover = add_table([
    ["SUBMITTED BY", "SUBMITTED TO"],
    ["Arshpreet Singh\nSarabjit Singh\nPooja\nUniversity Roll Nos.: 2323174, 2323218, 2323200", "Dr. Brijesh Bakariya\nProject Mentor and Head, CSE\nEr. Satyam Sharma\nProject Co-Mentor"],
], [3.1,3.1], 10)
p("Semester: 7th   |   Academic session: 2027", align=WD_ALIGN_PARAGRAPH.CENTER)
p("Submission date: ____________________", align=WD_ALIGN_PARAGRAPH.CENTER)
p("Hard-bound project report submission", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)
page_break()

# Declaration
title("DECLARATION")
p("We, the undersigned project team, declare that this report describes the EcoSort AI project work carried out by our team. Project implementation details and reported model metrics are based on the project repository and its saved metadata. External technical sources are listed in the references. This declaration is to be signed by all listed students after review.")
subsection("Student signatures")
add_table([["Name", "Signature", "Date"], ["Arshpreet Singh", "____________________________", "____________"], ["Sarabjit Singh", "____________________________", "____________"], ["Pooja", "____________________________", "____________"]], [2.0,2.8,1.2], 10)
p("Department: Computer Science and Engineering\nInstitution: IK Gujral Punjab Technical University, Hoshiarpur Campus\nProject title: EcoSort AI — AI-Powered Waste Classification\nProgramme: Bachelor of Technology in Computer Science Engineering\nSemester: 7th   |   Academic session: 2027")
page_break()

# Certificate page
title("CERTIFICATE")
p("This is to certify that the project entitled “EcoSort AI: AI-Powered Waste Classification” has been carried out by Arshpreet Singh, Sarabjit Singh, and Pooja in the Department of Computer Science and Engineering, IK Gujral Punjab Technical University, Hoshiarpur Campus, in partial fulfilment of the requirements for the Bachelor of Technology in Computer Science Engineering. The project was completed under the guidance of Dr. Brijesh Bakariya, Project Mentor and Head of Department, and Er. Satyam Sharma, Project Co-Mentor. This certificate is subject to institutional review and authorized signatures.")
for _ in range(4): doc.add_paragraph()
add_table([["Project Mentor / HOD", "Project Co-Mentor", "External Examiner"], ["Dr. Brijesh Bakariya", "Er. Satyam Sharma", "Name: __________________"], ["Signature: ______________", "Signature: ______________", "Signature: ______________"], ["Date: ___________________", "Date: ___________________", "Date: ___________________"]], [2.1,2.1,1.8], 9)
p("Institution seal / remarks", "Heading 2")
for _ in range(3): p("________________________________________________________________________________")
page_break()

# Abstract and preliminaries
title("ABSTRACT")
p("EcoSort AI is a deep-learning project that classifies a waste-item image into one of twelve supported dataset categories and presents the result through a responsive web application. The system combines a MobileNetV2 transfer-learning model with a Flask server and a browser interface supporting photo upload and camera capture. Camera users see a live preview, capture and review a still, then explicitly request prediction; the application does not classify the video stream continuously.")
p("The project dataset contains 15,515 images across battery, biological, brown glass, cardboard, clothes, green glass, metal, paper, plastic, shoes, trash, and white glass. Images use a 224 × 224 input size. The training notebook uses an 80/20 training-validation split with seed 42, class weighting, augmentation, early stopping, and fine-tuning of the final MobileNetV2 layers. Saved metadata reports 94.68% validation accuracy and 0.172 validation loss. These metrics describe the project's validation split, not an independent test or a guarantee for new photographs. The result is an end-to-end prototype joining data preparation, training, web inference, upload, and camera workflows.")
p("Keywords: waste classification, computer vision, transfer learning, MobileNetV2, TensorFlow, Flask, image recognition", italic=True)
page_break()
title("ACKNOWLEDGEMENTS")
p("The project team expresses sincere gratitude to Dr. Brijesh Bakariya, Project Mentor and Head of the Department of Computer Science and Engineering, IK Gujral Punjab Technical University, Hoshiarpur Campus, for his academic guidance and support. The team also thanks Er. Satyam Sharma, Project Co-Mentor, for his guidance and feedback during the project work.")
p("The team acknowledges the Department and the University for the academic environment in which the project was developed, and the authors and maintainers of the open-source frameworks cited in this report. Dataset source and licence information should be confirmed before redistribution.")
p("Project team\nArshpreet Singh\nSarabjit Singh\nPooja")
page_break()

# TOC fields, editable and auto-updatable in Word.
title("TABLE OF CONTENTS")
tocp = doc.add_paragraph()
fld = OxmlElement("w:fldSimple"); fld.set(qn("w:instr"), 'TOC \\o "1-2" \\h \\z \\u')
run = OxmlElement("w:r"); txt = OxmlElement("w:t"); txt.text = "Open in Word and update this field to refresh page numbers."; run.append(txt); fld.append(run); tocp._p.append(fld)
page_break()

title("LIST OF FIGURES")
figure_rows = [["FIGURE NO.", "TOPIC", "PAGE NO."], ["Figure 4.1", "Level-0 data-flow diagram", ""], ["Figure 4.2", "Level-1 data-flow diagram", ""], ["Figure 8.1", "Model development workflow", ""], ["Figure 9.1", "MobileNetV2 classifier pipeline", ""], ["Figure 10.1", "System architecture and request flow", ""], ["Figure 12.1", "EcoSort AI upload workflow preview", ""], ["Figure 12.2", "EcoSort AI camera capture workflow preview", ""], ["Figure 12.3", "Representative dataset class gallery", ""], ["Figure 13.1", "Training and validation history", ""]]
add_table(figure_rows, [1.2,4.2,.6], 9)
page_break()

# Extract numbered chapters from the project-authored report, then keep only the focused material needed
# for a regular academic report.
chapter_names = {
    1:"Introduction", 2:"Objectives and Scope", 3:"Feasibility Study", 4:"Tools and Technologies",
    5:"Data Flow Diagram", 6:"System Design and Implementation", 7:"Results and Discussion",
    8:"Conclusion and Future Scope", 9:"References and Appendix"
}
chapter_start = {}
for i, txt in enumerate(page_text):
    m = re.search(r"(?m)^\s*(\d{2})\s*\|\s*([^\n]+)", txt)
    if m:
        num = int(m.group(1))
        if num not in chapter_start:
            chapter_start[num] = i
chapter_start[39] = next((i for i,t in enumerate(page_text) if re.search(r"(?i)\b39\s*\|\s*Appendices",t)), len(page_text)-3)

# Basic inline-flow recreation of diagrams/screenshots. Use existing report-owned visuals for interface panels.
media = {
    "upload": BUILD / "upload_screen_detail.png",
    "camera": BUILD / "camera_screen_detail.png",
    "gallery": BUILD / "dataset_sample_gallery.png",
    "history": ROOT / "models" / "training_history.png",
}

def add_paragraph_text(text):
    text = text.strip()
    if not text: return
    if re.fullmatch(r"\d+\s*\|\s*\d+", text): return
    if re.fullmatch(r"\d+\s*\|\s*\d+\s*\|\s*\d+", text): return
    if text in {"PAGE", "CHAPTER", "TOPIC", "FIGURE NO.", "PAGE NO."}: return
    if text.startswith(("Chapter page references are generated", "Chapter page references will update")): return
    if re.fullmatch(r"\d+(\.\d+)*\s+.+", text) and len(text) < 120:
        depth = text.split(" ",1)[0].count(".")
        p(text, "Heading 2" if depth <= 1 else "Heading 3")
        return
    if re.fullmatch(r"\d+(\.\d+)?", text): return
    if text in {"Table of Contents", "List of Figure's", "List of Figures", "Figure's", "References", "Acknowledgements"}: return
    p(text)

source_chapters = {1:1, 2:2, 3:4, 4:6, 5:11, 6:10, 7:13, 8:38, 9:39}
source_ends = {1:2, 2:4, 3:6, 4:7, 5:12, 6:13, 7:14, 8:39, 9:46}
for number in range(1,10):
    page_break()
    title(f"CHAPTER {number}: {chapter_names[number].upper()}")
    source_number = source_chapters[number]
    start = chapter_start.get(source_number)
    if start is None: continue
    next_source = source_ends[number]
    end = chapter_start.get(next_source, len(page_text))
    if number == 8:
        end = next((i for i in range(start+1,len(page_text)) if "References" in page_text[i]), len(page_text))
    content = "\n".join(page_text[start:end])
    content = re.sub(r"ECOSORT AI\s*/\s*PROJECT REPORT.*?(?=\n|$)", "", content, flags=re.I)
    content = re.sub(r"\n\s*\d+\s*\|\s*\d+\s*\n\s*\d+\s*$", "", content)
    content = re.sub(r"(?m)^\s*\d{2}\s*\|\s*[^\n]+\n", "", content)
    lines = [x.strip() for x in content.splitlines()]
    lines = [x for x in lines if x and not re.fullmatch(r"\d+",x) and not re.fullmatch(r"\d+\s*\|\s*\d+(?:\s*\|\s*\d+)?",x)]
    # Keep only short representative material from each chapter; detailed technical explanation stays in the repo.
    filtered=[]
    para_count=0
    for line in lines:
        if re.match(r"\d+(?:\.\d+){1,2}\s+",line):
            if para_count < 3: filtered.append(line)
            continue
        if len(line)<100 and re.match(r"^(?:Figure|Table|Test area|Metric|Class|Component|State|Deployment|Upload|Camera|Privacy|Advantages|Limitations|Future|Conclusion|The repository|The project|The saved|The current|This project|EcoSort|At inference|For camera|For upload|The app|The model|Model|Training|Validation|Dataset|Python|Flask|TensorFlow|MobileNetV2|Pillow|NumPy|HTML|CSS|JavaScript|PWA|Open-source|The system|In the|When the|A user|The user|One image|The workflow|The interface|This chapter|Each class|Class weighting|The notebook|The route|The browser|A request|The backend|The frontend|The response|Image|Camera|Upload|System|The purpose|A future|Recommended|Current safeguards|The application|Operational|Economic|Technical|Feasibility|Scope|Project objectives)",line,re.I):
            if para_count < 3:
                filtered.append(line); para_count+=1
            continue
        if para_count < 3 and len(line)>120:
            filtered.append(line); para_count+=1
    lines=filtered
    # For chapter 12, insert the project-created UI previews at the figure captions.
    for line in lines:
        if re.search(r"Figure 13\.1", line) and media["history"].exists():
            add_image(media["history"], 5.8)
        # Known tables from the source chapter are rendered from compact data derived from the repository.
        if source_number == 5 and line.startswith("5.1 Class counts"):
            counts = meta["class_counts"]
            names = ["battery","biological","brown-glass","cardboard","clothes","green-glass","metal","paper","plastic","shoes","trash","white-glass"]
            add_table([["Class","Images","Share"]]+[[n.replace("-"," ").title(), f"{counts[n]:,}", f"{counts[n]/sum(counts.values())*100:.1f}%"] for n in names]+[["Total","15,515","100.0%"]], [2.5,1,1], 9)
        if source_number == 13 and line.startswith("13.1 Saved validation metrics"):
            add_table([["Metric","Saved value","Interpretation"],["Validation accuracy",f"{meta['validation_accuracy']*100:.2f}%","Recorded result on the validation split."],["Validation loss",f"{meta['validation_loss']:.4f}","Recorded loss on the validation split."],["Output categories","12","Configured classes."],["Dataset size","15,515","Sum of counts in saved class metadata."]],[1.5,1.2,3.0],8)
        if source_number == 33 and line.startswith("33.1 Recorded counts"):
            add_table([["Group","Count","Share"],["Clothes","5,325","34.3%"],["Shoes","1,977","12.7%"],["Paper","1,050","6.8%"],["Biological","985","6.3%"],["Other seven classes","5,233","33.7%"],["Total","15,515","100.0%"]],[2.7,1,1],8)
        # Turn standalone paragraphs and headings into Word-native text.
        if line.startswith("Figure "):
            p(line, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)
        elif re.fullmatch(r"(?:\d{1,2}\.\d+(?:\.\d+)?\s+|[A-Z][A-Za-z /&-]{4,70}$).+", line) and len(line) < 95:
            add_paragraph_text(line)
        else:
            add_paragraph_text(line)

        # Keep the report at ordinary project length: concise summaries instead of repeated technical chapters.
        if number == 3:
            p("3.1 Operational feasibility", "Heading 2")
            p("The browser interface provides a direct image-upload workflow and a camera workflow with preview, capture, review, and explicit prediction. The prototype can support a demonstration and learning use case, subject to normal server and device availability.")
            p("3.2 Technical feasibility", "Heading 2")
            p("The project uses Python 3.11, TensorFlow/Keras, MobileNetV2, Flask, NumPy, and Pillow. The repository includes trained model files, metadata, the labeled dataset, and deployment configuration. Browser camera access requires a secure origin such as HTTPS or localhost.")
            p("3.3 Economic feasibility", "Heading 2")
            p("The software stack is based on open-source frameworks. Operating costs depend on the host, compute requirements, storage, and traffic; no cost benchmark is recorded in the repository.")
        if number == 5:
            p("The data-flow diagram shows one image moving from the user through validation and preprocessing to the saved classifier, then returning ranked category scores to the interface.")
            p("User → Upload or camera capture → Flask validation → RGB resize to 224 × 224 → MobileNetV2 classifier → Ranked categories → User")
        if number == 9:
            p("The system accepts one image at a time, supports JPG, PNG, BMP, and WEBP uploads, validates model selection, and limits request size. Predictions are model suggestions; the system does not provide local recycling rules or guarantee classification on mixed or unfamiliar objects.")

# References from the source PDF, preserved as a separate academic section.
page_break(); title("REFERENCES")
ref_text = next((t for t in page_text if re.search(r"\[1\].*MobileNetV2",t)), "")
for ref in re.split(r"(?=\[\d+\])", ref_text):
    ref = ref.strip()
    if ref.startswith("["):
        p(ref)

# Finalize editable file with common Word defaults and metadata.
doc.core_properties.title = "EcoSort AI — AI-Powered Waste Classification Project Report"
doc.core_properties.author = "Arshpreet Singh; Sarabjit Singh; Pooja"
doc.core_properties.subject = "Bachelor of Technology, Computer Science Engineering, 7th Semester"
doc.save(DOCX_OUT)
print(f"Created {DOCX_OUT}")

# Print-ready PDF is built from the same concise chapter structure as the editable Word file.

pdf_styles = getSampleStyleSheet()
pdf_styles.add(ParagraphStyle(name="ReportTitle", parent=pdf_styles["Title"], fontName="Helvetica-Bold", fontSize=25, leading=30, textColor=colors.HexColor("#103E2E"), alignment=TA_CENTER, spaceAfter=10))
pdf_styles.add(ParagraphStyle(name="ReportHeading", parent=pdf_styles["Heading1"], fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=colors.HexColor("#103E2E"), spaceAfter=9))
pdf_styles.add(ParagraphStyle(name="ReportBody", parent=pdf_styles["BodyText"], fontName="Times-Roman", fontSize=10.4, leading=14.3, alignment=TA_JUSTIFY, spaceAfter=8))
pdf_styles.add(ParagraphStyle(name="ReportSmall", parent=pdf_styles["BodyText"], fontName="Helvetica", fontSize=9, leading=12, textColor=colors.HexColor("#31423A"), spaceAfter=5))
pdf_styles.add(ParagraphStyle(name="ReportCell", parent=pdf_styles["BodyText"], fontName="Helvetica", fontSize=8, leading=10, spaceAfter=0))
pdf_styles.add(ParagraphStyle(name="ReportTOC", parent=pdf_styles["BodyText"], fontName="Helvetica", fontSize=9, leading=13, spaceAfter=2))

pdf_story=[]
def rp(txt, sty="ReportBody"): return Paragraph(txt.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"),pdf_styles[sty])
def pdf_header(canvas, doc):
    canvas.saveState()
    if doc.page > 1:
        canvas.setFont("Helvetica-Bold",8); canvas.setFillColor(colors.HexColor("#176547")); canvas.drawString(20*mm,A4[1]-13*mm,"ECOSORT AI  /  PROJECT REPORT")
        canvas.setFont("Helvetica",8); canvas.setFillColor(colors.HexColor("#5B6B62")); canvas.drawRightString(A4[0]-20*mm,A4[1]-13*mm,"Garbage Classification with Deep Learning")
        canvas.setStrokeColor(colors.HexColor("#D7E2DA")); canvas.line(20*mm,A4[1]-16*mm,A4[0]-20*mm,A4[1]-16*mm)
        canvas.line(20*mm,15*mm,A4[0]-20*mm,15*mm)
        canvas.setFont("Helvetica",8); canvas.drawRightString(A4[0]-20*mm,10*mm,f"{doc.page}")
    canvas.restoreState()

pdf_story += [Spacer(1,10*mm), Image(str(logo),width=34*mm,height=34*mm), Spacer(1,8*mm), rp("IK GUJRAL PUNJAB TECHNICAL UNIVERSITY, HOSHIARPUR CAMPUS","ReportSmall"), rp("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING","ReportSmall"), Spacer(1,10*mm), rp("PROJECT REPORT","ReportSmall"), rp("EcoSort AI","ReportTitle"), rp("AI-Powered Waste Classification","ReportSmall"), rp("Garbage Classification with Deep Learning","ReportSmall"), Spacer(1,8*mm), rp("Submitted in partial fulfilment of the requirements for the award of the degree of Bachelor of Technology in Computer Science Engineering.","ReportSmall"), Spacer(1,8*mm)]
cover_data=[[rp("SUBMITTED BY","ReportCell"),rp("SUBMITTED TO","ReportCell")],[rp("Arshpreet Singh<br/>Sarabjit Singh<br/>Pooja<br/>Roll Nos.: 2323174, 2323218, 2323200","ReportCell"),rp("Dr. Brijesh Bakariya<br/>Project Mentor and HOD, CSE<br/>Er. Satyam Sharma<br/>Project Co-Mentor","ReportCell")]]
ct=Table(cover_data,colWidths=[78*mm,78*mm],hAlign="CENTER")
ct.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F7F3")),("GRID",(0,0),(-1,-1),.4,colors.HexColor("#D7E2DA")),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),9),("RIGHTPADDING",(0,0),(-1,-1),9),("TOPPADDING",(0,0),(-1,-1),9),("BOTTOMPADDING",(0,0),(-1,-1),9)]))
pdf_story += [ct,Spacer(1,7*mm),rp("Semester: 7th  |  Academic session: 2027  |  Submission date: ____________________","ReportSmall"),rp("Hard-bound project report submission","ReportSmall"),PageBreak()]

def pdf_table(rows,widths):
    data=[[Paragraph(str(cell).replace("&","&amp;"),pdf_styles["ReportCell"]) for cell in row] for row in rows]
    tb=Table(data,colWidths=widths,repeatRows=1,hAlign="LEFT")
    tb.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#176547")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.35,colors.HexColor("#D7E2DA")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F2F7F3")]),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
    return tb

pdf_story += [rp("DECLARATION","ReportHeading"),rp("We, the undersigned project team, declare that this report describes the EcoSort AI project work carried out by our team. Project implementation details and reported model metrics are based on the project repository and its saved metadata. External technical sources are listed in the references. This declaration is to be signed by all listed students after review."),pdf_table([["Name","Signature","Date"],["Arshpreet Singh","____________________________","____________"],["Sarabjit Singh","____________________________","____________"],["Pooja","____________________________","____________"]],[42*mm,75*mm,35*mm]),Spacer(1,8*mm),rp("Department: Computer Science and Engineering<br/>Institution: IK Gujral Punjab Technical University, Hoshiarpur Campus<br/>Semester: 7th  |  Academic session: 2027","ReportSmall"),PageBreak()]
pdf_story += [rp("CERTIFICATE","ReportHeading"),rp("This is to certify that the project entitled “EcoSort AI: AI-Powered Waste Classification” has been carried out by Arshpreet Singh, Sarabjit Singh, and Pooja in the Department of Computer Science and Engineering, IK Gujral Punjab Technical University, Hoshiarpur Campus, in partial fulfilment of the requirements for the Bachelor of Technology in Computer Science Engineering. The project was completed under the guidance of Dr. Brijesh Bakariya, Project Mentor and Head of Department, and Er. Satyam Sharma, Project Co-Mentor. This certificate is subject to institutional review and authorized signatures."),Spacer(1,18*mm),pdf_table([["Project Mentor / HOD","Project Co-Mentor","External Examiner"],["Dr. Brijesh Bakariya","Er. Satyam Sharma","Name: ________________"],["Signature: ____________","Signature: ____________","Signature: ____________"],["Date: _________________","Date: _________________","Date: _________________"]],[51*mm,51*mm,49*mm]),Spacer(1,18*mm),rp("Institution seal / remarks<br/><br/>________________________________________________________________________________<br/><br/>________________________________________________________________________________","ReportSmall"),PageBreak()]
pdf_story += [rp("ABSTRACT","ReportHeading"),rp("EcoSort AI is a deep-learning project that classifies a waste-item image into one of twelve supported dataset categories and presents the result through a responsive web application. The system combines a MobileNetV2 transfer-learning model with a Flask server and a browser interface supporting photo upload and camera capture. Camera users see a live preview, capture and review a still, then explicitly request prediction; the application does not classify the video stream continuously."),rp("The project dataset contains 15,515 images across battery, biological, brown glass, cardboard, clothes, green glass, metal, paper, plastic, shoes, trash, and white glass. Images use a 224 × 224 input size. The training notebook uses an 80/20 training-validation split with seed 42, class weighting, augmentation, early stopping, and fine-tuning of the final MobileNetV2 layers. Saved metadata reports 94.68% validation accuracy and 0.172 validation loss. These metrics describe the project's validation split, not an independent test or a guarantee for new photographs."),rp("Keywords: waste classification, computer vision, transfer learning, MobileNetV2, TensorFlow, Flask, image recognition.","ReportSmall"),PageBreak()]
pdf_story += [rp("ACKNOWLEDGEMENTS","ReportHeading"),rp("The project team expresses sincere gratitude to Dr. Brijesh Bakariya, Project Mentor and Head of the Department of Computer Science and Engineering, IK Gujral Punjab Technical University, Hoshiarpur Campus, for his academic guidance and support. The team also thanks Er. Satyam Sharma, Project Co-Mentor, for his guidance and feedback during the project work."),rp("The team acknowledges the Department and the University for the academic environment in which the project was developed, and the authors and maintainers of the open-source frameworks cited in this report. Dataset source and licence information should be confirmed before redistribution."),rp("Project team<br/>Arshpreet Singh<br/>Sarabjit Singh<br/>Pooja"),PageBreak()]

toc_rows=[["CHAPTER NO.","TOPIC","PAGE NO."]]
for n,name in chapter_names.items():
    toc_rows.append([f"CHAPTER {n}",name,str(8+n)])
toc_rows.append(["REFERENCES","Technical and project references","17"])
toc_rows.append(["APPENDICES","Repository, terminology, and submission records","18"])
pdf_story += [rp("TABLE OF CONTENTS","ReportHeading"),pdf_table(toc_rows,[31*mm,104*mm,18*mm]),PageBreak()]
fig_rows=[["FIGURE NO.","TOPIC","PAGE NO."],["Figure 5.1","Classifier data-flow diagram","13"],["Figure 6.1","MobileNetV2 model pipeline","14"],["Figure 7.1","Training and validation history","15"]]
pdf_story += [rp("LIST OF FIGURES","ReportHeading"),pdf_table(fig_rows,[30*mm,105*mm,18*mm]),PageBreak()]

# Matching concise chapter text, using extracted project chapters as source.
for n in range(1,10):
    if n > 1: pdf_story.append(PageBreak())
    pdf_story.append(rp(f"CHAPTER {n}: {chapter_names[n].upper()}","ReportHeading"))
    s_num=source_chapters[n]
    s=chapter_start.get(s_num)
    e=chapter_start.get(source_ends[n],len(page_text))
    if n==8: e=next((i for i in range(s+1,len(page_text)) if "References" in page_text[i]),len(page_text))
    raw="\n".join(page_text[s:e])
    raw=re.sub(r"ECOSORT AI\s*/\s*PROJECT REPORT.*?(?=\n|$)","",raw,flags=re.I)
    raw=re.sub(r"(?m)^\s*\d{2}\s*\|\s*[^\n]+\n","",raw)
    lines=[x.strip() for x in raw.splitlines() if x.strip() and not re.fullmatch(r"\d+(?:\s*\|\s*\d+)?",x.strip())]
    para_count=0
    for line in lines:
        if line.startswith("Figure 13.1") and media["history"].exists():
            pdf_story.append(Image(str(media["history"]),width=150*mm,height=78*mm))
        if re.match(r"\d+(?:\.\d+){1,2}\s+",line):
            if para_count<3: pdf_story.append(rp(line,"ReportSmall"))
        else:
            if para_count<3: pdf_story.append(rp(line)); para_count+=1
    if n==3:
        pdf_story += [rp("3.1 Operational feasibility","ReportSmall"),rp("The browser interface provides photo upload and camera capture with preview, review, and user-triggered prediction. It supports an academic demonstration subject to device and server availability."),rp("3.2 Technical feasibility","ReportSmall"),rp("Python, TensorFlow/Keras, MobileNetV2, Flask, NumPy, and Pillow form the software stack. Camera access requires HTTPS or localhost."),rp("3.3 Economic feasibility","ReportSmall"),rp("The software stack is open source. Hosting and compute costs depend on deployment needs and are not benchmarked in the repository.")]
    if n==5: pdf_story.append(rp("User → Upload or camera capture → Flask validation → RGB resize to 224 × 224 → MobileNetV2 classifier → Ranked categories → User","ReportSmall"))
    if n==9:
        pdf_story += [rp("Project map","ReportSmall"),pdf_table([["Path","Purpose"],["app.py","Flask routes, image processing, and inference."],["models/metadata.json","Class labels, image size, counts, and validation metrics."],["garbage_classification_model.ipynb","Dataset preparation, model training, and evaluation."],["templates/ and static/","Browser interface, camera workflow, styles, and PWA shell."]],[50*mm,100*mm]),rp("Submission checklist","ReportSmall"),rp("Confirm student names and roll numbers, verify dataset provenance and licensing, record live demonstration results, and obtain required mentor/department signatures before printing and hard binding.")]
    # Next chapter begins with a page break at the top of this loop.

pdfdoc=SimpleDocTemplate(str(PDF_OUT),pagesize=A4,leftMargin=20*mm,rightMargin=20*mm,topMargin=19*mm,bottomMargin=19*mm,title="EcoSort AI — AI-Powered Waste Classification Project Report",author="Arshpreet Singh; Sarabjit Singh; Pooja")
pdfdoc.build(pdf_story,onFirstPage=pdf_header,onLaterPages=pdf_header)
print(f"Created {PDF_OUT}")
