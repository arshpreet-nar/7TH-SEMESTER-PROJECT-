# EcoSort AI report format contract

## Reference inventory

| Reference | Purpose | Integrity / scope |
|---|---|---|
| `C:\Users\HP\Desktop\front page 1 - Manpreet.docx` | Cover layout and university seal | SHA-256 `312fdae7d026c509a012650bcc0cec48c7385bcb4848ffda8787466367a891ca`; 1 section; Letter page; retained unchanged. Reuse the institutional seal only; replace student/project fields. |
| `C:\Users\HP\Desktop\middle pages 2.docx` | Declaration, certification signatures, acknowledgement | SHA-256 not recorded in current extraction; 1 section; Letter page; retained unchanged. Use its declaration/acknowledgement sequence and signature slots; replace all Manpreet/leave-project details. |
| `C:\Users\HP\Desktop\Table of Contents 3.docx` | Contents and figure-list table pattern | SHA-256 `312fdae7d026c509a012650bcc0cec48c7385bcb4848ffda8787466367a891ca` was recorded for the cover source; TOC SHA must be freshly inventoried before later template reuse. One A4 section, three tables: contents, figures, and continuation. Use chapter/topic/page columns and figure/topic/page columns. |
| `C:\Users\HP\Desktop\sample  project report f.docx` | Academic chapter sequence and body typography | SHA-256 `5a45e730cc059af7bba880f96b855f781b148ccb7e4fb2fb48748db81b17e5c7`; one Letter section; retained unchanged. Its E-Cops text, Java stack and project claims do not apply. |
| `C:\Users\HP\Desktop\project file manpreet.docx` | Separate example of a project body and screenshot use | SHA-256 `c2bcde8f7850a812dfbea81f6fd75d2299841266b6cbba5096bd2a42ec0d9eaf8`; two Letter sections; retained unchanged. It describes Automation Leave Management and is not a source of EcoSort facts or images. |
| `C:\Users\HP\Desktop\Sample Project Complete file.pdf` | Visual sample of cover, declaration, acknowledgement, contents, figures, and chapter pages | 57 pages, A4 page box approximately 595.4 × 841.8 pt; retained unchanged. MedSwift content is excluded. |
| `C:\Users\HP\Desktop\project report format Instructions.doc` | Required report section checklist and print directions | Legacy Word document; 5 pages as observed in Word. Read-only reference. Checklist transcribed into the content map below. |

DOCX page counts vary by Word renderer and are not stored as stable document metadata. The reference files remain unchanged. Exact full-page reference renders were unavailable after Word's background automation was denied and interactive review was stopped; the contract therefore records observable section geometry, text structure, embedded seal, and the matching 57-page PDF sample.

## Page system and typography

- Cover and body pages: US Letter portrait (8.5 × 11 in). The cover template has 1 in margins on all sides.
- Inner-page pattern from the project DOCX and contents file: top 0.98 in, bottom 0.49 in, left 1.38 in for binding, right 0.49 in; header/footer distances approximately 0.49–0.50 in.
- Body: Times New Roman, 12 pt, justified, 1.5 line spacing as shown in the project sample. Chapter titles: centered, bold, approximately 18 pt. Numbered section titles: bold, approximately 16 pt.
- Cover title: centered 24 pt bold; project title 16 pt bold; degree line 14 pt bold; author/university fields centered. Use the university seal extracted from the supplied cover reference.
- Declaration title: centered Times New Roman 16 pt bold; include team declaration, student signatures, certification statement, examiner slots and Head of Department seal/signature area.
- Contents/list tables: three columns for chapter/figure/table identifier, topic, and page. Keep page numbers aligned in the final column; use compact but readable text to keep the contents on one or two real pages.
- One-sided printing; consistent page numbering. Omit a visible number on the cover, use Roman numerals for preliminary pages and Arabic numbering for the report body.

## Required content map

1. Cover page: EcoSort AI title, B.Tech CSE wording, team names, blank roll-number/session/date slots, mentor/co-mentor, Department and IKGPTU Hoshiarpur Campus.
2. Declaration and certification/signature slots.
3. Acknowledgements.
4. Abstract and keywords.
5. About the institution/department and project context (the relevant “About Company/Industry/Institute” slot; no company is claimed).
6. Table of contents, list of tables, list of figures, abbreviations/nomenclature.
7. Main chapters: introduction (objectives, scope, existing/proposed system, advantages/limitations, modules); feasibility; tools/technologies; DFD Level 0, 1 and 2; data/database design; implementation and interface screenshots; results/discussion; conclusion/future scope.
8. References, project datasheet (applicable), and appendices.

The supplied application uses file-based dataset/model artifacts and does not configure a relational database. The report must state this explicitly and provide a data/artifact dictionary with field names, types and identifiers where present. Do not invent database tables. The app has no login interface; mark login screenshots not applicable and document the actual home/upload, camera and result/error views available in the project.

## Editable content slots and exclusions

- Project facts come from the EcoSort repository, `models/metadata.json`, application code and the EcoSort source report.
- Reuse only the university seal from the Manpreet cover. Do not copy Manpreet's name, roll number, leave-system prose, database screenshots, screenshots, PHP/XAMPP stack, or claims into EcoSort.
- Preserve unknown roll numbers, approved academic session, signature/date, and dataset provenance as blank or clearly pending fields; do not guess.
- Keep references and claims traceable to the repository or listed technical sources. Distinguish validation metrics from independent test results.
- Include actual EcoSort DFD and technology visuals; add a level-2 inference flow and an honest data-persistence description. Do not add empty pages to satisfy length.

## Fidelity / QA gates

- Render the final PDF to page images and visually inspect page patterns, cover, signature page, contents/lists, chapters, tables, DFDs, screenshots, references and appendices.
- Confirm Letter page size, readable typography, no blank filler, no clipped tables/figures, consistent body numbering, accurate page references, and a physical length greater than 50 pages only through substantive sections/front matter.
- The Word hand-bound copy is page-faithful if the DOCX renderer cannot be recovered; disclose that its page images are not directly editable.
