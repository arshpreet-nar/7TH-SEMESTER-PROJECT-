import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const workspaceDir = "C:\\Users\\HP\\Desktop\\data science\\7TH-SEMESTER-PROJECT-";
const buildDir = path.join(workspaceDir, ".ppt-team-build");
const outputDir = path.join(workspaceDir, "Report");
const finalPath = path.join(outputDir, "EcoSort_AI_Short_College_Submission.pptx");
const skillDir = "C:\\Users\\HP\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.1015.12040\\skills\\presentations";
const runtimeHelpers = await import(pathToFileURL(path.join(skillDir, "container_tools/runtime_helpers.mjs")).href);
const utils = await import(pathToFileURL(path.join(skillDir, "container_tools/artifact_tool_utils.mjs")).href);
const { Presentation, PresentationFile } = await runtimeHelpers.importRuntimeModule("@oai/artifact-tool");
const { finalizePresentation, resolvePresentationFont, applyPresentationChartFont, makeNativeBulletParagraphs } = utils;

const W = 1280, H = 720;
const C = { deep: "#0D3B2B", forest: "#145C43", green: "#25865E", mint: "#E0F2E7", lime: "#B6DC37", blue: "#1886A2", bluePale: "#E2F3F7", coral: "#E76F57", coralPale: "#FCEBE6", ink: "#17251D", muted: "#5E6F65", line: "#D7E2D9", pale: "#F5F8F4", white: "#FFFFFF", darkMuted: "#C9DED0" };
const FONT = resolvePresentationFont({ fontFamily: "Arial" });
const presentation = Presentation.create({ slideSize: { width: W, height: H } });

function rect(slide, name, x, y, w, h, fill, stroke = "none", radius = 0) {
  return slide.shapes.add({ geometry: radius ? "roundRect" : "rect", name, position: { left: x, top: y, width: w, height: h }, fill, line: { style: "solid", fill: stroke, width: stroke === "none" ? 0 : 1 }, ...(radius ? { borderRadius: radius } : {}) });
}
function line(slide, name, x1, y1, x2, y2, color = C.line, width = 2) {
  return slide.shapes.add({ geometry: "line", name, position: { left: x1, top: y1, width: x2 - x1, height: y2 - y1 }, fill: "none", line: { style: "solid", fill: color, width } });
}
function text(slide, name, value, x, y, w, h, size, color = C.ink, options = {}) {
  const shape = slide.shapes.add({ geometry: "textbox", name, position: { left: x, top: y, width: w, height: h }, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  shape.text = value;
  shape.text.style = { typeface: FONT, fontSize: size, bold: Boolean(options.bold), color, alignment: options.align || "left", verticalAlignment: options.valign || "top", autoFit: "shrinkText", wrap: "square", insets: options.insets ?? 0 };
  return shape;
}
function bullets(slide, name, items, x, y, w, h, size = 22, color = C.ink, gap = 12) {
  const shape = slide.shapes.add({ geometry: "textbox", name, position: { left: x, top: y, width: w, height: h }, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  shape.text = makeNativeBulletParagraphs(items, { marginLeftPoints: 18, hangingPoints: 9, spaceAfterPoints: gap });
  shape.text.style = { typeface: FONT, fontSize: size, color, autoFit: "shrinkText", wrap: "square", insets: 0 };
  return shape;
}
function footer(slide, n, dark = false) {
  const color = dark ? C.darkMuted : C.muted;
  line(slide, `footer-rule-${n}`, 64, 675, 1216, 675, dark ? "#35634D" : C.line, 1);
  text(slide, `footer-brand-${n}`, "ECOSORT AI  /  COLLEGE PROJECT", 64, 685, 560, 20, 13, color, { bold: true });
  text(slide, `footer-num-${n}`, `${String(n).padStart(2, "0")} / 07`, 1120, 684, 95, 20, 13, color, { bold: true, align: "right" });
}
function title(slide, value, n, subtitle = "") {
  slide.background.fill = C.pale;
  rect(slide, `top-accent-${n}`, 0, 0, W, 8, C.forest);
  text(slide, `title-${n}`, value, 64, 42, 1140, 58, 42, C.deep, { bold: true });
  if (subtitle) text(slide, `subtitle-${n}`, subtitle, 66, 103, 1120, 45, 20, C.muted);
  footer(slide, n);
}
function notes(slide, source, cue) { slide.speakerNotes.text = `Project source: ${source}\nPresenter cue: ${cue}`; }
async function imageBytes(file) { const b = await fs.readFile(file); return new Uint8Array(b.buffer, b.byteOffset, b.byteLength); }

// 1. Cover
{
  const s = presentation.slides.add(); s.background.fill = C.deep;
  rect(s, "cover-left-rule", 0, 0, 15, H, C.lime); rect(s, "cover-right-field", 875, 0, 405, H, "#F3F7F1");
  s.images.add({ blob: await imageBytes(path.join(workspaceDir, "static", "garbage-ai-logo.jpg")), contentType: "image/jpeg", alt: "EcoSort AI recycling logo", fit: "contain", geometry: "ellipse", position: { left: 900, top: 155, width: 350, height: 350 } });
  text(s, "cover-eyebrow", "COLLEGE PROJECT SUBMISSION", 76, 92, 710, 30, 17, C.lime, { bold: true });
  text(s, "cover-title", "AI-Powered Waste\nClassification", 76, 160, 740, 138, 49, C.white, { bold: true });
  text(s, "cover-subtitle", "EcoSort AI", 80, 325, 680, 48, 31, "#BCE5C9", { bold: true });
  line(s, "cover-divider", 80, 400, 760, 400, "#35634D", 2);
  text(s, "cover-team-label", "PROJECT TEAM", 80, 430, 400, 25, 15, C.darkMuted, { bold: true });
  text(s, "cover-team", "Arshpreet Singh   |   Sarbjit Singh   |   Pooja", 80, 468, 740, 36, 20, C.white);
  text(s, "cover-field-label", "Deep Learning  /  Computer Vision  /  Sustainability", 80, 605, 740, 28, 17, C.darkMuted);
  line(s, "cover-footer-rule-dark", 64, 675, 860, 675, "#35634D", 1); line(s, "cover-footer-rule-light", 900, 675, 1216, 675, C.line, 1);
  text(s, "cover-footer-brand", "ECOSORT AI  /  COLLEGE PROJECT", 64, 685, 560, 20, 13, C.darkMuted, { bold: true });
  text(s, "cover-footer-number", "01 / 07", 1120, 684, 95, 20, 13, C.muted, { bold: true, align: "right" });
  notes(s, "Project repository and static/garbage-ai-logo.jpg.", "Introduce the team and the goal: classify a waste photo and support better sorting decisions.");
}
// 2. Objectives
{
  const s = presentation.slides.add(); title(s, "Project Objectives", 2, "Build a practical image-classification workflow that is easy to use on a phone.");
  const items = [
    ["Classify waste images", "Predict one of 12 dataset categories from a photo."],
    ["Provide a simple experience", "Support photo upload and camera capture with review before prediction."],
    ["Show useful model output", "Return the most likely category and ranked confidence scores."],
    ["Make access mobile-friendly", "Present the app as an installable, responsive web experience."],
  ];
  for (let i = 0; i < items.length; i++) {
    const y = 172 + i * 112; rect(s, `objective-num-bg-${i}`, 82, y, 52, 52, i % 2 ? C.blue : C.forest, "none", 26);
    text(s, `objective-num-${i}`, String(i + 1), 82, y + 5, 52, 40, 22, C.white, { bold: true, align: "center", valign: "middle" });
    text(s, `objective-heading-${i}`, items[i][0], 165, y + 2, 410, 38, 25, C.deep, { bold: true });
    text(s, `objective-copy-${i}`, items[i][1], 165, y + 43, 970, 48, 20, C.muted);
  }
  notes(s, "README.md, app.py, templates/index.html.", "These are the project goals; distinguish them from outcomes already measured.");
}
// 3. Methodology
{
  const s = presentation.slides.add(); title(s, "Methodology", 3, "Dataset preparation and transfer learning form the model-development pipeline.");
  const stages = [
    { x: 72, w: 258, color: C.forest, head: "1. Dataset", body: "15,515 labeled images\n12 waste classes" },
    { x: 368, w: 258, color: C.blue, head: "2. Prepare", body: "80 / 20 train-validation split\nRGB, resized to 224 x 224" },
    { x: 664, w: 258, color: C.coral, head: "3. Train", body: "MobileNetV2 transfer learning\nAugmentation and class weights" },
    { x: 960, w: 248, color: C.forest, head: "4. Evaluate", body: "Validation accuracy and loss\nSave best model checkpoint" },
  ];
  for (let i = 0; i < stages.length; i++) {
    const p = stages[i]; rect(s, `method-stage-${i}`, p.x, 215, p.w, 190, C.white, C.line, 5);
    rect(s, `method-color-${i}`, p.x, 215, p.w, 9, p.color);
    text(s, `method-head-${i}`, p.head, p.x + 18, 250, p.w - 36, 42, 24, p.color, { bold: true });
    text(s, `method-body-${i}`, p.body, p.x + 18, 310, p.w - 36, 68, 19, C.ink);
    if (i < stages.length - 1) line(s, `method-connector-${i}`, p.x + p.w + 8, 310, stages[i + 1].x - 8, 310, C.muted, 2);
  }
  rect(s, "method-model-band", 72, 460, 1136, 128, C.mint, "none", 5);
  text(s, "method-model-title", "Model approach", 96, 480, 250, 34, 23, C.deep, { bold: true });
  text(s, "method-model-body", "A pretrained MobileNetV2 feature extractor is paired with a new 12-class softmax head. The final 30 backbone layers are then fine-tuned at a lower learning rate.", 96, 522, 1070, 52, 20, C.ink);
  notes(s, "garbage_classification_model.ipynb, models/metadata.json, README.md.", "Describe the sequence and note that the 80/20 split is the reported project setup.");
}
// 4. Implementation
{
  const s = presentation.slides.add(); title(s, "Implementation", 4, "A responsive web interface connects photo capture to server-side model inference.");
  const blocks = [
    { x: 72, w: 330, color: C.blue, head: "User interface", body: "Responsive web app\nUpload or camera preview\nCapture, review, then predict" },
    { x: 475, w: 330, color: C.forest, head: "Flask backend", body: "Validate image request\nPreprocess to model input\nReturn ranked predictions" },
    { x: 878, w: 330, color: C.coral, head: "Classifier", body: "TensorFlow / Keras model\n12 class probabilities\nTop-five result display" },
  ];
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i]; rect(s, `implementation-box-${i}`, b.x, 205, b.w, 205, C.white, C.line, 5);
    rect(s, `implementation-bar-${i}`, b.x, 205, 8, 205, b.color);
    text(s, `implementation-head-${i}`, b.head, b.x + 28, 237, b.w - 52, 39, 25, b.color, { bold: true });
    text(s, `implementation-body-${i}`, b.body, b.x + 28, 295, b.w - 52, 92, 20, C.ink);
    if (i < blocks.length - 1) line(s, `implementation-link-${i}`, b.x + b.w + 12, 306, blocks[i + 1].x - 12, 306, C.blue, 3);
  }
  rect(s, "implementation-cam-band", 72, 455, 1136, 132, C.bluePale, "none", 5);
  text(s, "implementation-camera-title", "Camera prediction flow", 96, 478, 290, 34, 23, C.deep, { bold: true });
  text(s, "implementation-camera-body", "Live camera preview remains visible. The user captures and reviews a still image, then explicitly starts prediction. The captured image is sent to POST /api/predict-frame.", 96, 520, 1070, 55, 20, C.ink);
  notes(s, "app.py and templates/index.html; /api/predict-frame endpoint.", "Clarify this is a mobile-friendly web app and camera frames are not continuously classified.");
}
// 5. Results
{
  const s = presentation.slides.add(); title(s, "Results", 5, "Reported validation metrics for the selected model checkpoint.");
  rect(s, "results-acc-card", 72, 175, 500, 177, C.mint, "none", 5);
  text(s, "results-acc-value", "94.68%", 100, 198, 430, 72, 58, C.forest, { bold: true });
  text(s, "results-acc-label", "Validation accuracy", 104, 278, 420, 35, 22, C.deep, { bold: true });
  rect(s, "results-loss-card", 614, 175, 594, 177, C.bluePale, "none", 5);
  text(s, "results-loss-value", "0.172", 644, 198, 500, 72, 58, C.blue, { bold: true });
  text(s, "results-loss-label", "Validation loss", 648, 278, 480, 35, 22, C.deep, { bold: true });
  text(s, "results-dataset-title", "Evaluation context", 72, 403, 430, 38, 25, C.deep, { bold: true });
  bullets(s, "results-context", ["15,515 images across 12 classes", "80 / 20 training-validation split", "Metrics are from validation data, not an independent real-world test"], 72, 452, 1090, 143, 20, C.ink, 7);
  rect(s, "results-caveat", 72, 605, 1136, 42, C.coralPale, "none", 4);
  text(s, "results-caveat-copy", "Prediction quality may vary with lighting, image quality, backgrounds, and unfamiliar objects.", 92, 611, 1090, 30, 17, C.deep, { bold: true, valign: "middle" });
  notes(s, "models/metadata.json and README.md. Validation accuracy 0.9468256; validation loss 0.1720355.", "Do not describe validation accuracy as real-world accuracy; no independent test result is reported.");
}
// 6. Outcomes and team roles
{
  const s = presentation.slides.add(); title(s, "Outcomes and Team Roles", 6, "The project delivers a working end-to-end prototype and divides work across model, application, and user experience.");
  text(s, "outcome-heading", "Project outcomes", 72, 170, 500, 36, 25, C.forest, { bold: true });
  bullets(s, "outcome-list", ["12-class waste image classifier", "Photo upload and capture-review-predict workflow", "Mobile-friendly installable web application", "Validation benchmark with documented limitations"], 72, 216, 500, 230, 19, C.ink, 5);
  line(s, "team-divider", 620, 170, 620, 610, C.line, 2);
  text(s, "team-heading", "Team contributions", 665, 170, 520, 36, 25, C.blue, { bold: true });
  const team = [
    ["Arshpreet Singh", "Dataset preparation, model training, and evaluation"],
    ["Sarbjit Singh", "Flask backend, prediction API, and model integration"],
    ["Pooja", "User interface, camera workflow testing, and documentation"],
  ];
  for (let i = 0; i < team.length; i++) {
    const y = 226 + i * 112;
    text(s, `team-name-${i}`, team[i][0], 665, y, 500, 34, 22, C.deep, { bold: true });
    text(s, `team-role-${i}`, team[i][1], 665, y + 39, 510, 54, 18, C.muted);
  }
  text(s, "team-role-note", "Role allocation is proposed; confirm it matches each member's actual contribution.", 665, 570, 510, 45, 15, C.muted);
  notes(s, "Project repository and team names supplied by the user. Role split is proposed from project work areas.", "Confirm role allocation with the team before submission.");
}
// 7. Closing
{
  const s = presentation.slides.add(); s.background.fill = C.deep; rect(s, "closing-accent", 0, 0, W, 10, C.lime);
  text(s, "closing-title", "Thank You", 80, 180, 1120, 90, 64, C.white, { bold: true, align: "center" });
  text(s, "closing-subtitle", "EcoSort AI  |  AI-Powered Waste Classification", 120, 295, 1040, 45, 27, "#BCE5C9", { align: "center" });
  line(s, "closing-rule", 355, 380, 925, 380, "#35634D", 2);
  text(s, "closing-team", "Arshpreet Singh     Sarbjit Singh     Pooja", 130, 425, 1020, 45, 24, C.white, { align: "center" });
  text(s, "closing-questions", "Questions", 120, 525, 1040, 55, 31, C.lime, { bold: true, align: "center" });
  footer(s, 7, true); notes(s, "Project repository.", "Invite questions about methodology, results, or the app workflow.");
}

await fs.mkdir(buildDir, { recursive: true }); await fs.mkdir(outputDir, { recursive: true });
for (let i = 0; i < presentation.slides.items.length; i++) {
  const blob = await presentation.export({ slide: presentation.slides.items[i], format: "png", scale: 1 });
  await fs.writeFile(path.join(buildDir, `short-slide-${String(i + 1).padStart(2, "0")}.png`), new Uint8Array(await blob.arrayBuffer()));
}
const montage = await presentation.export({ format: "webp", montage: true });
await fs.writeFile(path.join(buildDir, "short-montage.webp"), new Uint8Array(await montage.arrayBuffer()));
const candidatePath = path.join(buildDir, "short-candidate.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);
const result = await finalizePresentation({
  workspaceDir, candidatePath, finalPath,
  pythonExecutable: "C:\\Users\\HP\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe",
  integrityValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-bullet-geometry", "--validate-heading-fit"],
  explicitTotalSlideCount: 7, requiredNativeChartOwnerSlides: [], fontPolicy: { basis: "design", families: [FONT] }, verifyArtifactToolImport: true,
  receiptPath: path.join(buildDir, "short-validation.json"),
});
console.log(JSON.stringify(result, null, 2));
