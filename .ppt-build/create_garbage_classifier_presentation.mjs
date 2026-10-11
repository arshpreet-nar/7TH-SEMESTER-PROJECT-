import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "C:\\Users\\HP\\Desktop\\data science\\7TH-SEMESTER-PROJECT-";
const buildDir = path.join(workspaceDir, ".ppt-build");
const outputDir = path.join(workspaceDir, "deliverables");
const skillDir = "C:\\Users\\HP\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.1009.11318\\skills\\presentations";
const runtimePython = "C:\\Users\\HP\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";
const runtimeNodeModules = "C:\\Users\\HP\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules";
const finalPptx = path.join(outputDir, "Garbage_Classification_Deep_Learning_Project_v3.pptx");

process.env.RUNTIME_NODE_MODULES = runtimeNodeModules;

const C = {
  ink: "#17231F",
  inkSoft: "#486058",
  forest: "#176B52",
  teal: "#159A8A",
  coral: "#E76F51",
  citrus: "#F4BE3B",
  cream: "#F8F5EE",
  mist: "#E6EEE9",
  white: "#FFFFFF",
  lightCoral: "#FCE9E4",
  lightTeal: "#E4F4F0",
  grid: "#D8E1DC",
};

const FONT = "Arial";

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function imageBytes(relativePath) {
  const bytes = await fs.readFile(path.join(workspaceDir, relativePath));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function textbox(slide, text, position, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    typeface: FONT,
    fontSize: 22,
    color: C.ink,
    ...style,
  };
  return shape;
}

function rect(slide, position, fill, line = { style: "solid", fill: "none", width: 0 }, geometry = "rect") {
  return slide.shapes.add({ geometry, position, fill, line });
}

function addTitle(slide, title, subtitle = "") {
  textbox(slide, title, { left: 68, top: 48, width: 940, height: 46 }, {
    fontSize: 33,
    bold: true,
    color: C.ink,
  });
  if (subtitle) {
    textbox(slide, subtitle, { left: 70, top: 102, width: 940, height: 30 }, {
      fontSize: 17,
      color: C.inkSoft,
    });
  }
  rect(slide, { left: 68, top: 138, width: 84, height: 5 }, C.coral);
}

function addSlideNumber(slide, number) {
  textbox(slide, String(number).padStart(2, "0"), { left: 1160, top: 672, width: 50, height: 22 }, {
    fontSize: 12,
    color: C.inkSoft,
    align: "right",
  });
}

function addProcessArrow(slide, x, y, width) {
  const arrow = slide.shapes.add({
    geometry: "rightArrow",
    position: { left: x, top: y, width, height: 22 },
    fill: C.citrus,
    line: { style: "solid", fill: C.citrus, width: 0 },
  });
  return arrow;
}

function addStage(slide, title, caption, x, y, width, tone) {
  rect(slide, { left: x, top: y, width, height: 124 }, C.white, { style: "solid", fill: C.grid, width: 1 }, "roundRect");
  rect(slide, { left: x, top: y, width, height: 9 }, tone, { style: "solid", fill: tone, width: 0 }, "roundRect");
  textbox(slide, title, { left: x + 16, top: y + 27, width: width - 32, height: 32 }, {
    fontSize: title === "Augmentation" || title === "Preprocess" ? 16 : title.length >= 10 ? 18 : 20,
    bold: true,
    color: C.ink,
    align: "center",
  });
  textbox(slide, caption, { left: x + 14, top: y + 68, width: width - 28, height: 42 }, {
    fontSize: 14,
    color: C.inkSoft,
    align: "center",
  });
}

async function main() {
  await fs.mkdir(buildDir, { recursive: true });
  await fs.mkdir(outputDir, { recursive: true });
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  presentation.theme.defaultFont = FONT;

  // Slide 1: cover
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.cream;
    rect(slide, { left: 0, top: 0, width: 24, height: 720 }, C.forest);
    rect(slide, { left: 24, top: 0, width: 11, height: 720 }, C.coral);
    textbox(slide, "DEEP LEARNING PROJECT", { left: 72, top: 88, width: 360, height: 24 }, {
      fontSize: 14,
      bold: true,
      color: C.forest,
      charSpacing: 1.4,
    });
    textbox(slide, "Garbage\nClassification", { left: 68, top: 135, width: 560, height: 182 }, {
      fontSize: 55,
      bold: true,
      color: C.ink,
    });
    textbox(slide, "Image classification with MobileNetV2, TensorFlow, Flask, and Render", { left: 72, top: 343, width: 508, height: 60 }, {
      fontSize: 22,
      color: C.inkSoft,
    });
    rect(slide, { left: 72, top: 444, width: 482, height: 2 }, C.grid);
    textbox(slide, "12 waste classes    /    15,515 images    /    94.68% validation accuracy", { left: 72, top: 470, width: 600, height: 34 }, {
      fontSize: 17,
      bold: true,
      color: C.forest,
    });
    const collage = [
      ["garbage_classification/battery/battery1.jpg", "image/jpeg", 704, 76, 220, 220, "Battery sample"],
      ["garbage_classification/plastic/plastic1.jpg", "image/jpeg", 944, 76, 220, 220, "Plastic sample"],
      ["garbage_classification/paper/paper1.jpg", "image/jpeg", 704, 316, 220, 220, "Paper sample"],
      ["garbage_classification/brown-glass/brown-glass1.jpg", "image/jpeg", 944, 316, 220, 220, "Brown glass sample"],
    ];
    for (const [file, type, left, top, width, height, alt] of collage) {
      slide.images.add({
        blob: await imageBytes(file),
        contentType: type,
        alt,
        fit: "cover",
        position: { left, top, width, height },
        geometry: "roundRect",
        borderRadius: "rounded-lg",
      });
    }
    slide.speakerNotes.text = "Source: project dataset samples and models/metadata.json.";
  }

  // Slide 2: project objective
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.white;
    addTitle(slide, "Project objective", "Convert a trained waste-classification model into a usable browser experience");
    textbox(slide, "The project takes a user-uploaded image and identifies the most likely waste category.", { left: 78, top: 205, width: 525, height: 82 }, {
      fontSize: 26,
      bold: true,
      color: C.ink,
    });
    const objectives = [
      ["01", "Classify waste", "Recognize 12 categories, including plastic, glass, paper, metal, shoes, and clothes."],
      ["02", "Show confidence", "Return the top five predicted labels so users can judge the model's certainty."],
      ["03", "Make it accessible", "Serve saved Keras models through a Flask web application that accepts common image files."],
    ];
    let y = 330;
    for (const [n, title, body] of objectives) {
      textbox(slide, n, { left: 82, top: y, width: 54, height: 30 }, { fontSize: 18, bold: true, color: C.coral });
      textbox(slide, title, { left: 150, top: y - 1, width: 270, height: 28 }, { fontSize: 20, bold: true, color: C.ink });
      textbox(slide, body, { left: 150, top: y + 30, width: 450, height: 48 }, { fontSize: 16, color: C.inkSoft });
      rect(slide, { left: 80, top: y + 91, width: 515, height: 1 }, C.grid);
      y += 105;
    }
    slide.images.add({
      blob: await imageBytes("garbage_classification/trash/trash1.jpg"),
      contentType: "image/jpeg",
      alt: "Trash sample from the project dataset",
      fit: "cover",
      position: { left: 704, top: 182, width: 430, height: 414 },
      geometry: "roundRect",
      borderRadius: "rounded-xl",
    });
    rect(slide, { left: 704, top: 548, width: 430, height: 48 }, C.forest, { style: "solid", fill: C.forest, width: 0 }, "roundRect");
    textbox(slide, "Dataset example: trash", { left: 728, top: 560, width: 360, height: 24 }, { fontSize: 16, bold: true, color: C.white });
    addSlideNumber(slide, 2);
    slide.speakerNotes.text = "Source: app.py and templates/index.html. Image source: project dataset.";
  }

  // Slide 3: dataset chart
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.cream;
    addTitle(slide, "Training dataset", "15,515 labelled images stored in one folder per waste class");
    textbox(slide, "The classes are imbalanced. Class weights were used during training to reduce the influence of overrepresented categories.", { left: 73, top: 165, width: 520, height: 66 }, {
      fontSize: 19,
      color: C.inkSoft,
    });
    slide.charts.add("bar", {
      position: { left: 78, top: 246, width: 738, height: 390 },
      categories: ["clothes", "shoes", "paper", "biological", "battery", "cardboard", "plastic", "white-glass", "metal", "trash", "green-glass", "brown-glass"],
      series: [{
        name: "Images",
        values: [5325, 1977, 1050, 985, 945, 891, 865, 775, 769, 697, 629, 607],
        fill: C.teal,
      }],
      hasLegend: false,
      dataLabels: { showValue: true, position: "outEnd" },
      chartFill: C.cream,
      plotAreaFill: C.cream,
      xAxis: {
        majorGridlines: { style: "solid", fill: C.grid, width: 1 },
        tickLabelColor: C.inkSoft,
        tickLabelFontSize: 11,
      },
      yAxis: { tickLabelColor: C.ink, tickLabelFontSize: 12 },
    });
    rect(slide, { left: 890, top: 262, width: 280, height: 238 }, C.white, { style: "solid", fill: C.grid, width: 1 }, "roundRect");
    textbox(slide, "Dataset highlights", { left: 918, top: 290, width: 220, height: 28 }, { fontSize: 19, bold: true, color: C.ink });
    textbox(slide, "Largest class", { left: 918, top: 344, width: 150, height: 18 }, { fontSize: 14, color: C.inkSoft });
    textbox(slide, "Clothes  5,325", { left: 918, top: 367, width: 200, height: 30 }, { fontSize: 20, bold: true, color: C.forest });
    textbox(slide, "Smallest class", { left: 918, top: 422, width: 150, height: 18 }, { fontSize: 14, color: C.inkSoft });
    textbox(slide, "Brown-glass  607", { left: 918, top: 445, width: 220, height: 30 }, { fontSize: 20, bold: true, color: C.coral });
    textbox(slide, "20% of the dataset is reserved for validation.", { left: 888, top: 536, width: 275, height: 56 }, { fontSize: 17, color: C.inkSoft });
    addSlideNumber(slide, 3);
    slide.speakerNotes.text = "Source: models/metadata.json and garbage_classification_model.ipynb.";
  }

  // Slide 4: model training pipeline
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.white;
    addTitle(slide, "Model training pipeline", "MobileNetV2 transfer learning with augmentation, class weighting, and fine-tuning");
    const stageY = 245;
    addStage(slide, "Dataset folders", "12 labelled classes", 75, stageY, 172, C.forest);
    addProcessArrow(slide, 252, stageY + 50, 34);
    addStage(slide, "Augmentation", "Flip, rotation, zoom", 296, stageY, 172, C.teal);
    addProcessArrow(slide, 473, stageY + 50, 34);
    addStage(slide, "MobileNetV2", "ImageNet base model", 517, stageY, 172, C.coral);
    addProcessArrow(slide, 694, stageY + 50, 34);
    addStage(slide, "Fine-tuning", "Last layers trainable", 738, stageY, 172, C.citrus);
    addProcessArrow(slide, 915, stageY + 50, 34);
    addStage(slide, "Keras models", "Saved .keras artifacts", 959, stageY, 172, C.forest);

    const settings = [
      ["224 x 224", "RGB input size"],
      ["32", "Batch size"],
      ["12 + 5", "Training and fine-tune epochs"],
      ["20%", "Validation split"],
    ];
    let x = 100;
    for (const [value, label] of settings) {
      textbox(slide, value, { left: x, top: 473, width: 190, height: 32 }, { fontSize: 23, bold: true, color: C.forest, align: "center" });
      textbox(slide, label, { left: x, top: 510, width: 190, height: 26 }, { fontSize: 14, color: C.inkSoft, align: "center" });
      x += 260;
    }
    slide.images.add({
      blob: await imageBytes("garbage_classification/shoes/shoes1.jpg"),
      contentType: "image/jpeg",
      alt: "Shoe image used as an example input",
      fit: "cover",
      position: { left: 1020, top: 560, width: 108, height: 86 },
      geometry: "roundRect",
      borderRadius: "rounded-lg",
    });
    textbox(slide, "Example input", { left: 900, top: 588, width: 100, height: 22 }, { fontSize: 14, color: C.inkSoft, align: "right" });
    addSlideNumber(slide, 4);
    slide.speakerNotes.text = "Source: garbage_classification_model.ipynb. The notebook uses MobileNetV2, data augmentation, class weights, EarlyStopping, and ReduceLROnPlateau.";
  }

  // Slide 5: validation results
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.cream;
    addTitle(slide, "Validation results", "The best saved model reaches 94.68% validation accuracy with a 0.172 validation loss");
    slide.images.add({
      blob: await imageBytes("models/training_history.png"),
      contentType: "image/png",
      alt: "Training and validation accuracy and loss across epochs",
      fit: "contain",
      position: { left: 56, top: 174, width: 708, height: 420 },
      geometry: "rect",
    });
    rect(slide, { left: 818, top: 184, width: 330, height: 160 }, C.forest, { style: "solid", fill: C.forest, width: 0 }, "roundRect");
    textbox(slide, "94.68%", { left: 850, top: 219, width: 270, height: 56 }, { fontSize: 42, bold: true, color: C.white, align: "center" });
    textbox(slide, "validation accuracy", { left: 850, top: 282, width: 270, height: 26 }, { fontSize: 17, color: C.white, align: "center" });
    rect(slide, { left: 818, top: 374, width: 330, height: 160 }, C.lightCoral, { style: "solid", fill: C.lightCoral, width: 0 }, "roundRect");
    textbox(slide, "0.172", { left: 850, top: 409, width: 270, height: 50 }, { fontSize: 38, bold: true, color: C.coral, align: "center" });
    textbox(slide, "validation loss", { left: 850, top: 468, width: 270, height: 24 }, { fontSize: 17, color: C.inkSoft, align: "center" });
    textbox(slide, "3,103 validation images", { left: 818, top: 575, width: 330, height: 28 }, { fontSize: 18, bold: true, color: C.ink, align: "center" });
    textbox(slide, "Training history includes 16 recorded epochs.", { left: 818, top: 610, width: 330, height: 28 }, { fontSize: 15, color: C.inkSoft, align: "center" });
    addSlideNumber(slide, 5);
    slide.speakerNotes.text = "Source: models/metadata.json and models/training_history.png. Validation dataset count comes from garbage_classification_model.ipynb.";
  }

  // Slide 6: web application flow
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.white;
    addTitle(slide, "Web application integration", "Flask loads the saved model once and returns ranked probabilities for every valid upload");
    slide.images.add({
      blob: await imageBytes("garbage_classification/testing/paper.jpg"),
      contentType: "image/jpeg",
      alt: "Paper image used as an example upload",
      fit: "cover",
      position: { left: 74, top: 206, width: 234, height: 304 },
      geometry: "roundRect",
      borderRadius: "rounded-xl",
    });
    textbox(slide, "Example upload", { left: 74, top: 528, width: 234, height: 26 }, { fontSize: 16, bold: true, color: C.ink, align: "center" });
    textbox(slide, "JPG, PNG, BMP, or WEBP\nMaximum file size: 8 MB", { left: 74, top: 559, width: 234, height: 48 }, { fontSize: 14, color: C.inkSoft, align: "center" });
    const flow = [
      ["Upload", "User chooses an image", C.forest],
      ["Validate", "Extension and size checks", C.teal],
      ["Preprocess", "RGB resize to 224 x 224", C.coral],
      ["Predict", "Keras model returns probabilities", C.citrus],
      ["Display", "Top five labels and confidence", C.forest],
    ];
    let x = 362;
    for (let i = 0; i < flow.length; i += 1) {
      const [title, caption, tone] = flow[i];
      addStage(slide, title, caption, x, 264, 142, tone);
      if (i < flow.length - 1) addProcessArrow(slide, x + 146, 316, 24);
      x += 166;
    }
    textbox(slide, "The app defaults to best_garbage_classifier.keras and lets the user select either saved model.", { left: 365, top: 452, width: 735, height: 48 }, {
      fontSize: 19,
      color: C.ink,
      align: "center",
    });
    rect(slide, { left: 432, top: 540, width: 598, height: 1 }, C.grid);
    textbox(slide, "Core stack: Flask  |  TensorFlow / Keras  |  Pillow  |  NumPy", { left: 418, top: 565, width: 630, height: 28 }, { fontSize: 17, bold: true, color: C.forest, align: "center" });
    addSlideNumber(slide, 6);
    slide.speakerNotes.text = "Source: app.py and templates/index.html. Example image source: garbage_classification/testing/paper.jpg.";
  }

  // Slide 7: Render deployment
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.cream;
    addTitle(slide, "Live deployment on Render", "The repository contains the service definition and start command for a public Python web service");
    const stages = [
      ["GitHub repository", "Source code, models, and configuration", C.forest],
      ["Render build", "pip install -r requirements.txt", C.teal],
      ["Gunicorn service", "Starts Flask on Render's assigned port", C.coral],
      ["Public web app", "Browser uploads an image for classification", C.citrus],
    ];
    let x = 80;
    for (let i = 0; i < stages.length; i += 1) {
      const [title, caption, tone] = stages[i];
      addStage(slide, title, caption, x, 238, 216, tone);
      if (i < stages.length - 1) addProcessArrow(slide, x + 220, 290, 45);
      x += 282;
    }
    textbox(slide, "Deployment files", { left: 79, top: 456, width: 220, height: 30 }, { fontSize: 20, bold: true, color: C.ink });
    const deployFiles = [
      ["render.yaml", "Defines build command, Gunicorn start command, and Python environment variable"],
      ["runtime.txt", "Records Python 3.11.9 for compatible deployment tooling"],
      [".python-version", "Pins the project runtime to Python 3.11.9"],
      ["Procfile", "Provides a process command for compatible hosts"],
    ];
    let fy = 508;
    for (const [file, description] of deployFiles) {
      textbox(slide, file, { left: 82, top: fy, width: 180, height: 24 }, { fontSize: 17, bold: true, color: C.forest });
      textbox(slide, description, { left: 265, top: fy, width: 690, height: 24 }, { fontSize: 15, color: C.inkSoft });
      fy += 38;
    }
    addSlideNumber(slide, 7);
    slide.speakerNotes.text = "Source: render.yaml, runtime.txt, .python-version, Procfile, requirements.txt, and README.md.";
  }

  // Slide 8: deployment issue and fix
  {
    const slide = presentation.slides.add();
    slide.background.fill = C.white;
    addTitle(slide, "Render deployment issue and fix", "Pin Python 3.11.9 because the default Python 3.14.3 build cannot install the requested TensorFlow package");
    rect(slide, { left: 82, top: 194, width: 490, height: 292 }, C.lightCoral, { style: "solid", fill: C.lightCoral, width: 0 }, "roundRect");
    textbox(slide, "Why the build failed", { left: 116, top: 230, width: 360, height: 32 }, { fontSize: 25, bold: true, color: C.coral });
    textbox(slide, "Render selected Python 3.14.3 by default. The package installer then reported that no compatible tensorflow>=2.15 distribution was available.", { left: 116, top: 290, width: 400, height: 90 }, { fontSize: 19, color: C.ink });
    textbox(slide, "Observed build error:\nNo matching distribution found for tensorflow>=2.15", { left: 116, top: 405, width: 410, height: 55 }, { fontSize: 15, bold: true, color: C.inkSoft });
    rect(slide, { left: 640, top: 194, width: 558, height: 292 }, C.lightTeal, { style: "solid", fill: C.lightTeal, width: 0 }, "roundRect");
    textbox(slide, "Configured solution", { left: 677, top: 230, width: 380, height: 32 }, { fontSize: 25, bold: true, color: C.forest });
    textbox(slide, "1. Set the Render environment variable PYTHON_VERSION to 3.11.9.\n\n2. Keep runtime.txt and .python-version set to 3.11.9.\n\n3. Clear the build cache, then deploy again.", { left: 677, top: 288, width: 445, height: 145 }, { fontSize: 19, color: C.ink });
    rect(slide, { left: 180, top: 544, width: 920, height: 82 }, C.forest, { style: "solid", fill: C.forest, width: 0 }, "roundRect");
    textbox(slide, "Expected result after redeploy: TensorFlow installs on Python 3.11.9, and Gunicorn starts the Flask app.", { left: 220, top: 564, width: 840, height: 44 }, { fontSize: 17, bold: true, color: C.white, align: "center" });
    addSlideNumber(slide, 8);
    slide.speakerNotes.text = "Source: Render build log supplied for this project, render.yaml, runtime.txt, .python-version, and README.md.";
  }

  for (const [index, slide] of presentation.slides.items.entries()) {
    await writeBlob(path.join(buildDir, `slide-${String(index + 1).padStart(2, "0")}.png`), await presentation.export({ slide, format: "png", scale: 1 }));
  }
  await writeBlob(path.join(buildDir, "deck-montage.webp"), await presentation.export({ format: "webp", montage: true }));

  const { finalizePresentation } = await import(pathToFileURL(path.join(skillDir, "container_tools", "artifact_tool_utils.mjs")).href);
  const candidatePath = path.join(buildDir, "candidate.pptx");
  await (await PresentationFile.exportPptx(presentation)).save(candidatePath);
  const result = await finalizePresentation({
    explicitTotalSlideCount: 8,
    requiredNativeChartOwnerSlides: [3],
    materializeLiteralChartWorkbooks: true,
    workspaceDir,
    candidatePath,
    finalPath: finalPptx,
    pythonExecutable: runtimePython,
    integrityValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_package_integrity.py"),
    layoutValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_layout_geometry.py"),
    layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-bullet-geometry", "--validate-heading-fit"],
    fontPolicy: { basis: "design", families: [FONT] },
    verifyArtifactToolImport: true,
    receiptPath: path.join(buildDir, "Garbage_Classification_Deep_Learning_Project_v3.validation.json"),
  });
  console.log(JSON.stringify({ finalPptx, result }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
