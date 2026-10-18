import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const workspaceDir = "C:\\Users\\HP\\Desktop\\data science\\7TH-SEMESTER-PROJECT-";
const buildDir = path.join(workspaceDir, ".ppt-team-build");
const outputDir = path.join(workspaceDir, "Report");
const finalPath = path.join(outputDir, "EcoSort_AI_Final_College_Presentation_Revision.pptx");
const skillDir = "C:\\Users\\HP\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.1015.12040\\skills\\presentations";
const runtimeHelpers = await import(pathToFileURL(path.join(skillDir, "container_tools/runtime_helpers.mjs")).href);
const utils = await import(pathToFileURL(path.join(skillDir, "container_tools/artifact_tool_utils.mjs")).href);
const { Presentation, PresentationFile } = await runtimeHelpers.importRuntimeModule("@oai/artifact-tool");
const { finalizePresentation, resolvePresentationFont, applyPresentationChartFont, makeNativeBulletParagraphs } = utils;

const W = 1280;
const H = 720;
const C = {
  deep: "#0D3B2B",
  forest: "#145C43",
  green: "#25865E",
  mint: "#E0F2E7",
  lime: "#B6DC37",
  blue: "#1886A2",
  bluePale: "#E2F3F7",
  coral: "#E76F57",
  coralPale: "#FCEBE6",
  ink: "#17251D",
  muted: "#5E6F65",
  line: "#D7E2D9",
  pale: "#F5F8F4",
  white: "#FFFFFF",
  darkMuted: "#C9DED0",
};
const FONT = resolvePresentationFont({ fontFamily: "Arial" });
const presentation = Presentation.create({ slideSize: { width: W, height: H } });

function rect(slide, name, x, y, w, h, fill, stroke = "none", radius = 0) {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: stroke, width: stroke === "none" ? 0 : 1 },
    ...(radius ? { borderRadius: radius } : {}),
  });
}

function line(slide, name, x1, y1, x2, y2, color = C.line, width = 2) {
  return slide.shapes.add({
    geometry: "line",
    name,
    position: { left: x1, top: y1, width: x2 - x1, height: y2 - y1 },
    fill: "none",
    line: { style: "solid", fill: color, width },
  });
}

function text(slide, name, value, x, y, w, h, size, color = C.ink, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = value;
  shape.text.style = {
    typeface: FONT,
    fontSize: size,
    bold: Boolean(options.bold),
    color,
    alignment: options.align || "left",
    verticalAlignment: options.valign || "top",
    autoFit: "shrinkText",
    wrap: "square",
    insets: options.insets ?? 0,
    ...(options.lineSpacing ? { lineSpacing: options.lineSpacing } : {}),
  };
  return shape;
}

function bullets(slide, name, items, x, y, w, h, size = 21, color = C.ink, gap = 9) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = makeNativeBulletParagraphs(items, {
    marginLeftPoints: 18,
    hangingPoints: 9,
    spaceAfterPoints: gap,
  });
  shape.text.style = {
    typeface: FONT,
    fontSize: size,
    color,
    autoFit: "shrinkText",
    wrap: "square",
    insets: 0,
  };
  return shape;
}

function addFooter(slide, n, dark = false) {
  const color = dark ? C.darkMuted : C.muted;
  line(slide, `footer-rule-${n}`, 64, 675, 1216, 675, dark ? "#35634D" : C.line, 1);
  text(slide, `footer-brand-${n}`, "ECOSORT AI  /  COLLEGE PROJECT", 64, 685, 560, 20, 13, color, { bold: true });
  text(slide, `footer-number-${n}`, String(n).padStart(2, "0"), 1175, 684, 40, 20, 13, color, { bold: true, align: "right" });
}

function addTitle(slide, title, n, subtitle = "") {
  slide.background.fill = C.pale;
  rect(slide, `top-accent-${n}`, 0, 0, W, 8, C.forest);
  text(slide, `slide-title-${n}`, title, 64, 42, 1140, 58, 42, C.deep, { bold: true });
  if (subtitle) text(slide, `slide-subtitle-${n}`, subtitle, 66, 103, 1120, 42, 20, C.muted);
  addFooter(slide, n);
}

function addNotes(slide, source, presenter = "") {
  slide.speakerNotes.text = `Project source: ${source}${presenter ? `\nPresenter cue: ${presenter}` : ""}`;
}

async function imageBytes(filePath) {
  const bytes = await fs.readFile(filePath);
  return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function addNumberStep(slide, n, x, y, label, detail, color = C.forest, width = 250) {
  rect(slide, `step-number-bg-${n}-${x}`, x, y, 46, 46, color, "none", 23);
  text(slide, `step-number-${n}-${x}`, String(n), x, y + 4, 46, 36, 23, C.white, { bold: true, align: "center", valign: "middle" });
  text(slide, `step-label-${n}-${x}`, label, x, y + 62, width, 38, 22, C.deep, { bold: true });
  text(slide, `step-detail-${n}-${x}`, detail, x, y + 105, width, 88, 18, C.muted);
}

async function addAppPreview(slide, name, x, y, w, h, mode) {
  rect(slide, `${name}-frame`, x, y, w, h, C.white, C.line, 6);
  rect(slide, `${name}-chrome`, x, y, w, 26, "#E8EEE9", "none", 6);
  for (let i = 0; i < 3; i++) rect(slide, `${name}-dot-${i}`, x + 14 + i * 17, y + 9, 8, 8, [C.coral, C.lime, C.green][i], "none", 4);
  rect(slide, `${name}-brand-mark`, x + 18, y + 42, 30, 30, C.forest, "none", 15);
  text(slide, `${name}-brand`, "EcoSort AI", x + 58, y + 43, 170, 26, 17, C.deep, { bold: true });
  text(slide, `${name}-eyebrow`, "AI-POWERED WASTE SORTING", x + 20, y + 84, w - 40, 18, 10, C.green, { bold: true });
  text(slide, `${name}-headline`, mode === "upload" ? "Know where it goes." : "Point. Scan. Sort.", x + 20, y + 106, w - 40, 31, 22, C.deep, { bold: true });
  text(slide, `${name}-subhead`, mode === "upload" ? "Choose a photo to identify a waste item." : "Capture an item, review it, then predict.", x + 20, y + 140, w - 40, 25, 12, C.muted);
  const tabY = y + 176;
  rect(slide, `${name}-tab-active`, x + 18, tabY, (w - 42) / 2, 28, mode === "upload" ? C.mint : C.pale, "none", 4);
  rect(slide, `${name}-tab-inactive`, x + 24 + (w - 42) / 2, tabY, (w - 42) / 2, 28, mode === "upload" ? C.pale : C.mint, "none", 4);
  text(slide, `${name}-tab-upload`, "Upload a photo", x + 20, tabY + 6, (w - 46) / 2, 17, 11, mode === "upload" ? C.deep : C.muted, { bold: mode === "upload", align: "center" });
  text(slide, `${name}-tab-camera`, "Scan live", x + 26 + (w - 42) / 2, tabY + 6, (w - 46) / 2, 17, 11, mode === "camera" ? C.deep : C.muted, { bold: mode === "camera", align: "center" });
  const panelY = tabY + 40;
  rect(slide, `${name}-left-panel`, x + 18, panelY, w * 0.47, h - 257, C.white, C.line, 4);
  text(slide, `${name}-panel-title`, mode === "upload" ? "What do you want to sort?" : "Live camera", x + 30, panelY + 12, w * 0.43, 22, 13, C.deep, { bold: true });
  if (mode === "upload") text(slide, `${name}-model`, "Model  ·  best_garbage_classifier.keras", x + 30, panelY + 42, w * 0.43, 20, 9, C.muted);
  if (mode === "upload") {
    rect(slide, `${name}-drop`, x + 30, panelY + 70, w * 0.41, 84, C.pale, C.line, 4);
    text(slide, `${name}-drop-copy`, "Choose a clear photo\nJPG, PNG, BMP or WEBP", x + 43, panelY + 91, w * 0.37, 44, 11, C.muted, { align: "center", valign: "middle" });
    rect(slide, `${name}-action`, x + 30, panelY + 165, w * 0.41, 30, C.forest, "none", 4);
    text(slide, `${name}-action-copy`, "Find the category", x + 33, panelY + 172, w * 0.40, 19, 11, C.white, { bold: true, align: "center" });
  } else {
    text(slide, `${name}-camera-instructions`, "Center one item in the camera preview.", x + 30, panelY + 43, w * 0.43, 28, 10, C.muted);
    rect(slide, `${name}-capture`, x + 30, panelY + 82, w * 0.41, 42, C.forest, "none", 4);
    text(slide, `${name}-capture-label`, "Capture photo", x + 33, panelY + 92, w * 0.40, 22, 12, C.white, { bold: true, align: "center" });
    text(slide, `${name}-review-note`, "Review the captured image, then select Predict image.", x + 30, panelY + 140, w * 0.43, 43, 10, C.muted);
  }
  const rx = x + w * 0.52;
  rect(slide, `${name}-result-panel`, rx, panelY, w * 0.44, h - 257, C.pale, C.line, 4);
  text(slide, `${name}-result-kicker`, mode === "upload" ? "SCAN RESULT" : "CAPTURE REVIEW", rx + 10, panelY + 10, w * 0.40, 16, 9, C.green, { bold: true });
  const imgPath = path.join(workspaceDir, "garbage_classification", "plastic", "plastic1.jpg");
  slide.images.add({ blob: await imageBytes(imgPath), contentType: "image/jpeg", alt: "Example plastic item from the project's classification dataset, shown in the app interface preview", fit: "cover", position: { left: rx + 10, top: panelY + 33, width: w * 0.40, height: mode === "upload" ? 90 : 106 } });
  text(slide, `${name}-class`, mode === "upload" ? "Plastic" : "Captured item", rx + 10, panelY + (mode === "upload" ? 129 : 145), w * 0.40, 24, 14, C.deep, { bold: true });
  text(slide, `${name}-confidence`, mode === "upload" ? "Likely category · sample preview" : "Ready for prediction", rx + 10, panelY + (mode === "upload" ? 155 : 169), w * 0.40, 21, 10, C.muted);
}

// 1. Cover
{
  const slide = presentation.slides.add();
  slide.background.fill = C.deep;
  rect(slide, "cover-left-rule", 0, 0, 15, H, C.lime);
  rect(slide, "cover-right-field", 875, 0, 405, H, "#F3F7F1");
  slide.images.add({
    blob: await imageBytes(path.join(workspaceDir, "static", "garbage-ai-logo.jpg")),
    contentType: "image/jpeg",
    alt: "EcoSort AI recycling logo with a smart waste bin",
    fit: "contain",
    geometry: "ellipse",
    position: { left: 900, top: 145, width: 350, height: 350 },
  });
  text(slide, "cover-eyebrow", "COLLEGE PROJECT SUBMISSION", 76, 92, 710, 30, 17, C.lime, { bold: true });
  text(slide, "cover-title", "AI-Powered\nWaste Classification", 76, 155, 740, 150, 51, C.white, { bold: true });
  text(slide, "cover-subtitle", "EcoSort AI", 80, 330, 680, 48, 31, "#BCE5C9", { bold: true });
  line(slide, "cover-divider", 80, 405, 760, 405, "#35634D", 2);
  text(slide, "cover-team-label", "PROJECT TEAM", 80, 435, 400, 25, 15, C.darkMuted, { bold: true });
  text(slide, "cover-team", "Arshpreet Singh   |   Sarbjit Singh   |   Pooja", 80, 472, 740, 36, 20, C.white);
  text(slide, "cover-institution", "IK Gujral Punjab Technical University  /  Hoshiarpur Campus", 80, 528, 740, 24, 16, C.darkMuted);
  text(slide, "cover-department", "Department of Computer Science and Engineering", 80, 552, 740, 24, 16, C.darkMuted);
  text(slide, "cover-mentors", "Project Mentor and HOD: Dr. Brijesh Bakariya  /  Project Co-Mentor: Er. Satyam Sharma", 80, 581, 770, 36, 14, C.darkMuted);
  text(slide, "cover-field-label", "Deep Learning  /  Computer Vision  /  Sustainability", 80, 605, 740, 28, 17, C.darkMuted);
  line(slide, "cover-footer-rule-dark", 64, 675, 860, 675, "#35634D", 1);
  line(slide, "cover-footer-rule-light", 900, 675, 1216, 675, C.line, 1);
  text(slide, "cover-footer-brand", "ECOSORT AI  /  COLLEGE PROJECT", 64, 685, 560, 20, 13, C.darkMuted, { bold: true });
  text(slide, "cover-footer-number", "01", 1175, 684, 40, 20, 13, C.muted, { bold: true, align: "right" });
  addNotes(slide, "Project repository, app brand asset in static/garbage-ai-logo.jpg.", "Introduce the team and state the project's goal in one sentence.");
}

// 2. Project overview
{
  const slide = presentation.slides.add();
  addTitle(slide, "Project Objectives", 2, "A browser-based assistant that predicts the material category of a waste photo.");
  line(slide, "overview-divider", 625, 175, 625, 595, C.line, 2);
  text(slide, "overview-problem-head", "The problem", 72, 175, 480, 42, 30, C.forest, { bold: true });
  text(slide, "overview-problem", "People often face uncertainty when sorting everyday waste. A wrong choice can contaminate otherwise recyclable material.", 72, 232, 480, 130, 24, C.ink);
  text(slide, "overview-goal-head", "Project goal", 680, 175, 500, 42, 30, C.blue, { bold: true });
  text(slide, "overview-goal", "Use a trained image classifier to return a likely waste class and confidence scores through a simple photo or camera workflow.", 680, 232, 500, 132, 24, C.ink);
  text(slide, "overview-delivers-head", "What the project delivers", 72, 410, 500, 38, 25, C.deep, { bold: true });
  bullets(slide, "overview-delivers", [
    "12-class image classification model",
    "Photo upload and ranked prediction results",
    "Camera preview with capture, review, and predict steps",
    "Mobile-friendly installable web app",
  ], 72, 458, 1085, 178, 20, C.ink, 5);
  addNotes(slide, "README.md, app.py, templates/index.html.", "Explain that the app is a decision aid. It does not replace local waste instructions.");
}

// 3. Dataset
{
  const slide = presentation.slides.add();
  addTitle(slide, "Dataset and Waste Classes", 3, "Folder-based image dataset with one directory for each target label.");
  text(slide, "dataset-total", "15,515", 72, 170, 260, 62, 52, C.forest, { bold: true });
  text(slide, "dataset-total-label", "images across 12 classes", 74, 237, 260, 30, 19, C.muted);
  text(slide, "dataset-largest", "Clothes is the largest class", 72, 300, 325, 38, 24, C.deep, { bold: true });
  text(slide, "dataset-imbalance", "5,325 images, about 34% of the dataset. Class weights help address this imbalance during training.", 72, 350, 330, 112, 20, C.muted);
  slide.charts.add("bar", {
    position: { left: 430, top: 150, width: 770, height: 470 },
    title: "Image count by class",
    titlePlacement: "aboveChart",
    titleTextStyle: { fontSize: 20, bold: true, fill: C.deep },
    categories: ["battery", "biological", "brown-glass", "cardboard", "clothes", "green-glass", "metal", "paper", "plastic", "shoes", "trash", "white-glass"],
    series: [{ name: "Images", values: [945, 985, 607, 891, 5325, 629, 769, 1050, 865, 1977, 697, 775], fill: C.green }],
    hasLegend: false,
    barOptions: { direction: "bar", grouping: "clustered", gapWidth: 50 },
    xAxis: { min: 0, max: 6000, majorUnit: 1000, numberFormatCode: "#,##0", textStyle: { fontSize: 13, fill: C.muted }, majorGridlines: { style: "solid", fill: C.line, width: 1 } },
    yAxis: { textStyle: { fontSize: 13, fill: C.ink }, majorGridlines: null },
    chartFill: "none",
    chartLine: { style: "solid", fill: "none", width: 0 },
    plotAreaFill: "none",
    plotAreaLine: { style: "solid", fill: "none", width: 0 },
  });
  applyPresentationChartFont(slide.charts.items.at(-1), { fontFamily: FONT });
  addNotes(slide, "models/metadata.json and README.md. Counts sum to 15,515 images.", "Point out the class imbalance and connect it to class-weighted training.");
}

// 4. Data preparation
{
  const slide = presentation.slides.add();
  addTitle(slide, "Methodology: Image Preparation", 4, "The notebook loads labeled folders into a repeatable training and validation pipeline.");
  line(slide, "prep-flow-line", 145, 245, 1110, 245, C.line, 4);
  addNumberStep(slide, 1, 92, 210, "Organize", "Each class has its own folder. TensorFlow infers labels from folder names.", C.forest, 230);
  addNumberStep(slide, 2, 385, 210, "Split", "80% training and 20% validation with seed 42 for repeatability.", C.blue, 230);
  addNumberStep(slide, 3, 678, 210, "Resize", "Convert images to RGB and resize to 224 x 224 pixels.", C.coral, 230);
  addNumberStep(slide, 4, 971, 210, "Batch", "Load in batches of 32 and prefetch the input pipeline.", C.forest, 230);
  rect(slide, "prep-augmentation-band", 72, 480, 1128, 127, C.mint, "none", 6);
  text(slide, "prep-augmentation-title", "Training-time augmentation", 96, 501, 330, 36, 23, C.deep, { bold: true });
  text(slide, "prep-augmentation-copy", "Random horizontal flip, rotation, zoom, and contrast variation help expose the model to modest changes in object position and image appearance.", 96, 548, 1060, 48, 20, C.ink);
  addNotes(slide, "garbage_classification_model.ipynb, sections 3-8.", "Mention that augmentation is used during training, not to alter the user's captured image in the interface.");
}

// 5. Model architecture
{
  const slide = presentation.slides.add();
  addTitle(slide, "Methodology: MobileNetV2", 5, "Transfer learning reuses visual features and learns a new classifier for the project classes.");
  const x = [70, 310, 560, 840, 1060];
  const widths = [185, 195, 225, 195, 155];
  const labels = [
    ["Input image", "224 x 224\nRGB"],
    ["Augmentation", "Flip, rotate,\nzoom, contrast"],
    ["MobileNetV2", "ImageNet backbone\ninitially frozen"],
    ["Classification head", "Global pooling\nDropout 0.25"],
    ["12 outputs", "Softmax\nclass scores"],
  ];
  const fills = [C.bluePale, C.mint, C.forest, C.coralPale, C.lime];
  const titleColors = [C.deep, C.deep, C.white, C.deep, C.deep];
  const bodyColors = [C.muted, C.muted, C.darkMuted, C.muted, C.ink];
  for (let i = 0; i < labels.length; i++) {
    rect(slide, `model-stage-${i}`, x[i], 250, widths[i], 205, fills[i], "none", 7);
    text(slide, `model-title-${i}`, labels[i][0], x[i] + 14, 278, widths[i] - 28, 58, 22, titleColors[i], { bold: true, align: "center", valign: "middle" });
    text(slide, `model-detail-${i}`, labels[i][1], x[i] + 14, 350, widths[i] - 28, 68, 18, bodyColors[i], { align: "center", valign: "middle" });
    if (i < labels.length - 1) {
      line(slide, `model-connector-${i}`, x[i] + widths[i] + 8, 350, x[i + 1] - 9, 350, C.blue, 3);
    }
  }
  text(slide, "model-headline", "Fine-tuning", 72, 500, 210, 36, 25, C.forest, { bold: true });
  text(slide, "model-finetune-copy", "After training the new head, the notebook unfreezes the final 30 MobileNetV2 layers and fine-tunes them at a lower learning rate.", 72, 547, 1050, 75, 22, C.ink);
  addNotes(slide, "garbage_classification_model.ipynb, section 9 and section 11.", "Explain why transfer learning is useful when building a classifier for a focused dataset.");
}

// 6. Training strategy
{
  const slide = presentation.slides.add();
  addTitle(slide, "Methodology: Training Strategy", 6, "The notebook trains a new classification head, then fine-tunes part of the pretrained backbone.");
  rect(slide, "training-stage-one-band", 72, 170, 550, 190, C.mint, "none", 6);
  text(slide, "training-stage-one-number", "01", 96, 187, 74, 55, 37, C.forest, { bold: true });
  text(slide, "training-stage-one-title", "Train the new head", 183, 194, 390, 36, 26, C.deep, { bold: true });
  text(slide, "training-stage-one-copy", "Freeze MobileNetV2 and learn the 12-class output layer for up to 12 epochs.", 96, 260, 480, 75, 21, C.ink);
  rect(slide, "training-stage-two-band", 658, 170, 550, 190, C.bluePale, "none", 6);
  text(slide, "training-stage-two-number", "02", 682, 187, 74, 55, 37, C.blue, { bold: true });
  text(slide, "training-stage-two-title", "Fine-tune the backbone", 769, 194, 410, 36, 26, C.deep, { bold: true });
  text(slide, "training-stage-two-copy", "Unfreeze the final 30 layers and train for up to 5 more epochs at learning rate 1e-5.", 682, 260, 480, 75, 21, C.ink);
  text(slide, "training-controls-title", "Controls used during training", 72, 410, 600, 38, 26, C.forest, { bold: true });
  bullets(slide, "training-controls", [
    "Class weights address unequal class counts.",
    "Early stopping monitors validation loss and restores best weights.",
    "ReduceLROnPlateau lowers the learning rate when validation loss stalls.",
    "A checkpoint keeps the model with the best validation accuracy.",
  ], 72, 460, 1110, 175, 20, C.ink, 4);
  addNotes(slide, "garbage_classification_model.ipynb, sections 7, 10, and 11.", "The maximum epochs are upper bounds because early stopping can end training sooner.");
}

// 7. Results
{
  const slide = presentation.slides.add();
  addTitle(slide, "Model Results", 7, "Saved validation metrics show how the selected model performed on the project validation split.");
  text(slide, "result-accuracy", "94.68%", 72, 175, 360, 75, 58, C.forest, { bold: true });
  text(slide, "result-accuracy-label", "Validation accuracy", 76, 250, 350, 34, 22, C.muted, { bold: true });
  text(slide, "result-loss", "0.172", 72, 340, 270, 64, 47, C.blue, { bold: true });
  text(slide, "result-loss-label", "Validation loss", 76, 405, 350, 30, 22, C.muted, { bold: true });
  text(slide, "result-caveat", "Validation split only. The repository does not report an independent test score. New images may differ in lighting, background, or object type.", 72, 490, 390, 100, 20, C.ink);
  text(slide, "result-abstention", "Below 40% top score: show an 'Image not confidently classified' popup.", 72, 600, 710, 30, 17, C.coral, { bold: true });
  slide.images.add({
    blob: await imageBytes(path.join(workspaceDir, "models", "training_history.png")),
    contentType: "image/png",
    alt: "Training and validation accuracy and loss curves across model epochs",
    fit: "contain",
    position: { left: 490, top: 160, width: 720, height: 440 },
  });
  addNotes(slide, "models/metadata.json and models/training_history.png. Validation accuracy 0.9468256, validation loss 0.1720355.", "Separate this validation result from an independent real-world test score, which the repository does not report.");
}

// 8. Product workflow
{
  const slide = presentation.slides.add();
  addTitle(slide, "Implementation: App Screens", 8, "Upload or capture a still. Scores below the 40% threshold trigger a low-confidence popup.");
  await addAppPreview(slide, "upload-preview", 72, 160, 548, 450, "upload");
  await addAppPreview(slide, "camera-preview", 660, 160, 548, 450, "camera");
  text(slide, "upload-preview-caption", "UPLOAD SCREEN  /  Interface preview", 82, 620, 520, 24, 15, C.blue, { bold: true });
  text(slide, "camera-preview-caption", "CAMERA SCREEN  /  Interface preview", 670, 620, 520, 24, 15, C.coral, { bold: true });
  addNotes(slide, "templates/index.html and app.py routes / and /api/predict-frame. Interface previews are reconstructed from the project's implemented controls; the Flask/TensorFlow server was not available for live capture.", "The browser shows a live camera preview. The user captures and reviews a still image, then starts one prediction. The app does not continuously classify video frames.");
}

// 9. System architecture
{
  const slide = presentation.slides.add();
  addTitle(slide, "Implementation: Application Architecture", 9, "The browser captures or uploads an image. Flask runs TensorFlow inference and returns ranked labels.");
  const top = 200;
  const boxes = [
    { x: 72, w: 240, fill: C.bluePale, title: "Browser", detail: "Photo upload\nCamera preview\nCaptured still" },
    { x: 382, w: 240, fill: C.mint, title: "Flask app", detail: "Validate request\nChoose saved model\nReturn response" },
    { x: 692, w: 240, fill: C.coralPale, title: "Image pipeline", detail: "Decode image\nRGB conversion\nResize to 224 x 224" },
    { x: 1002, w: 206, fill: C.deep, title: "Keras model", detail: "Class probabilities\nTop five predictions" },
  ];
  for (let i = 0; i < boxes.length; i++) {
    const b = boxes[i];
    rect(slide, `arch-box-${i}`, b.x, top, b.w, 205, b.fill, "none", 6);
    const titleColor = i === 3 ? C.white : C.deep;
    const detailColor = i === 3 ? C.darkMuted : C.muted;
    text(slide, `arch-title-${i}`, b.title, b.x + 18, top + 25, b.w - 36, 40, 24, titleColor, { bold: true, align: "center" });
    text(slide, `arch-detail-${i}`, b.detail, b.x + 16, top + 90, b.w - 32, 90, 18, detailColor, { align: "center", valign: "middle" });
    if (i < boxes.length - 1) line(slide, `arch-link-${i}`, b.x + b.w + 12, top + 103, boxes[i + 1].x - 12, top + 103, C.blue, 3);
  }
  rect(slide, "arch-endpoint-band", 72, 470, 1136, 108, C.white, C.line, 5);
  text(slide, "arch-endpoint-label", "CAMERA API", 94, 492, 145, 28, 15, C.blue, { bold: true });
  text(slide, "arch-endpoint", "POST /api/predict-frame", 255, 485, 370, 36, 24, C.deep, { bold: true });
  text(slide, "arch-endpoint-detail", "Accepts the selected model and a captured JPEG data URL. The response contains the top label and up to five scores.", 255, 527, 890, 38, 18, C.muted);
  addNotes(slide, "app.py, especially predict_image(), load_model(), index(), and predict_frame().", "Describe the difference between the client-side camera stream and the single still image sent for inference.");
}

// 10. Mobile and hosting
{
  const slide = presentation.slides.add();
  addTitle(slide, "Implementation: Mobile App and Deployment", 10, "The project runs as an installable web app. The model stays on the server and needs a network connection.");
  text(slide, "mobile-left-head", "Mobile experience", 72, 170, 490, 42, 29, C.forest, { bold: true });
  bullets(slide, "mobile-left-list", [
    "Responsive layout for phone and desktop screens.",
    "PWA manifest, app icon, and service worker for the web shell.",
    "Rear-facing camera is preferred when the browser supports it.",
    "HTTPS is required for camera access on a hosted site.",
  ], 72, 228, 520, 280, 20, C.ink, 7);
  line(slide, "mobile-column-rule", 632, 176, 632, 590, C.line, 2);
  text(slide, "mobile-right-head", "Hosting configuration", 690, 170, 510, 42, 29, C.blue, { bold: true });
  text(slide, "mobile-runtime", "Python 3.11.9", 690, 232, 460, 45, 28, C.deep, { bold: true });
  text(slide, "mobile-runtime-detail", "Pinned for TensorFlow package compatibility.", 690, 280, 480, 35, 19, C.muted);
  text(slide, "mobile-command", "gunicorn --bind 0.0.0.0:$PORT app:app", 690, 345, 505, 47, 19, C.forest, { bold: true });
  text(slide, "mobile-deploy-detail", "Render reads render.yaml, installs requirements.txt, then runs the Flask app through Gunicorn.", 690, 402, 480, 88, 20, C.ink);
  rect(slide, "mobile-online-note", 690, 525, 485, 80, C.coralPale, "none", 5);
  text(slide, "mobile-online-text", "Camera permissions and predictions depend on a secure, reachable website.", 710, 543, 445, 47, 19, C.deep, { bold: true, valign: "middle" });
  addNotes(slide, "static/manifest.webmanifest, static/service-worker.js, render.yaml, Procfile, runtime.txt, and .python-version.", "Clarify that this is a mobile-friendly PWA, not a native Android APK, and inference is server-side.");
}

// 11. Team roles
{
  const slide = presentation.slides.add();
  addTitle(slide, "Team Roles", 11, "Suggested role allocation mapped to the project's machine learning, application, and presentation work.");
  const people = [
    { x: 72, name: "Arshpreet Singh", role: "Dataset and ML", color: C.forest, tasks: "Dataset review\nMobileNetV2 training\nClass weighting\nModel evaluation" },
    { x: 470, name: "Sarbjit Singh", role: "Backend and integration", color: C.blue, tasks: "Flask routes and APIs\nModel loading and inference\nUpload and camera integration\nDeployment configuration" },
    { x: 868, name: "Pooja", role: "UI, testing, and documentation", color: C.coral, tasks: "Responsive interface\nCapture and review flow\nMobile PWA experience\nTesting and project report" },
  ];
  for (let i = 0; i < people.length; i++) {
    const p = people[i];
    text(slide, `team-name-${i}`, p.name, p.x, 205, 335, 54, 31, C.deep, { bold: true });
    rect(slide, `team-role-underline-${i}`, p.x, 270, 68, 6, p.color);
    text(slide, `team-role-${i}`, p.role, p.x, 300, 335, 54, 22, p.color, { bold: true });
    text(slide, `team-tasks-${i}`, p.tasks, p.x, 376, 335, 150, 21, C.ink);
    if (i < people.length - 1) line(slide, `team-divider-${i}`, p.x + 365, 202, p.x + 365, 545, C.line, 2);
  }
  rect(slide, "team-shared-band", 72, 570, 1136, 49, C.mint, "none", 5);
  text(slide, "team-shared", "Shared work: planning, integration review, project demonstration, and submission preparation.", 94, 580, 1090, 30, 19, C.deep, { bold: true, valign: "middle" });
  addNotes(slide, "Team names supplied by the user. Role assignments are proposed from the project's work areas because contribution history is not recorded in the repository.", "Confirm the role split reflects the team's actual contributions before presenting.");
}

// 12. Limitations and roadmap
{
  const slide = presentation.slides.add();
  addTitle(slide, "Outcomes and Next Steps", 12, "A working end-to-end prototype, with validation results and clear areas for further testing.");
  text(slide, "limits-head", "Project outcomes and limitations", 72, 175, 510, 40, 26, C.coral, { bold: true });
  bullets(slide, "limits-list", [
    "One image receives one dominant class, even when several objects appear.",
    "The 12 labels do not cover every waste material or local disposal rule.",
    "The saved metadata does not include per-class precision, recall, or a confusion matrix.",
    "Softmax scores are not a calibrated guarantee of correctness.",
  ], 72, 230, 520, 305, 20, C.ink, 6);
  line(slide, "limits-divider", 635, 180, 635, 580, C.line, 2);
  text(slide, "next-head", "Recommended improvements", 690, 175, 510, 40, 28, C.blue, { bold: true });
  bullets(slide, "next-list", [
    "Add a held-out test set with per-class metrics and error analysis.",
    "Collect more balanced images across phones, lighting, and backgrounds.",
    "Calibrate the 40% low-confidence threshold on a held-out set.",
    "Explore detection for multiple items and verified local disposal guidance.",
  ], 690, 230, 520, 305, 20, C.ink, 6);
  rect(slide, "limits-conclusion", 72, 570, 1136, 53, C.deep, "none", 5);
  text(slide, "limits-conclusion-text", "The project demonstrates an end-to-end path from labeled images to a usable waste-classification app.", 94, 580, 1088, 34, 20, C.white, { bold: true, valign: "middle" });
  addNotes(slide, "models/metadata.json and garbage_classification_model.ipynb. The repository reports validation metrics but no independent test metrics.", "Discuss next steps as future work, not as features already implemented.");
}

// 13. Closing
{
  const slide = presentation.slides.add();
  slide.background.fill = C.deep;
  rect(slide, "closing-accent", 0, 0, W, 10, C.lime);
  text(slide, "closing-title", "Thank you", 80, 180, 1120, 100, 68, C.white, { bold: true, align: "center" });
  text(slide, "closing-subtitle", "EcoSort AI  |  AI-Powered Waste Classification", 120, 300, 1040, 45, 27, "#BCE5C9", { align: "center" });
  line(slide, "closing-rule", 355, 390, 925, 390, "#35634D", 2);
  text(slide, "closing-team", "Arshpreet Singh     Sarbjit Singh     Pooja", 130, 425, 1020, 45, 24, C.white, { align: "center" });
  text(slide, "closing-question", "Questions", 120, 530, 1040, 55, 31, C.lime, { bold: true, align: "center" });
  addFooter(slide, 13, true);
  addNotes(slide, "Project repository.", "Invite questions about the model, camera capture flow, or validation results.");
}

await fs.mkdir(buildDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
for (let i = 0; i < presentation.slides.items.length; i++) {
  const slide = presentation.slides.items[i];
  const blob = await presentation.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(buildDir, `slide-${String(i + 1).padStart(2, "0")}.png`), new Uint8Array(await blob.arrayBuffer()));
}
const montage = await presentation.export({ format: "webp", montage: true });
await fs.writeFile(path.join(buildDir, "montage.webp"), new Uint8Array(await montage.arrayBuffer()));

const candidatePath = path.join(buildDir, "candidate.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);
const result = await finalizePresentation({
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable: "C:\\Users\\HP\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe",
  integrityValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: [
    "--expected-slide-size-emu", "12192000,6858000",
    "--validate-bullet-geometry",
    "--validate-heading-fit",
  ],
  explicitTotalSlideCount: 13,
  requiredNativeChartOwnerSlides: [3],
  nativeChartTargetApplication: "powerpoint",
  materializeLiteralChartWorkbooks: true,
  fontPolicy: { basis: "design", families: [FONT] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(buildDir, "validation.json"),
});
console.log(JSON.stringify(result, null, 2));
