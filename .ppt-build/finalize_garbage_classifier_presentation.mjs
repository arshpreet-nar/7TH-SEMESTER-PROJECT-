import path from "node:path";
import { pathToFileURL } from "node:url";

const workspaceDir = "C:\\Users\\HP\\Desktop\\data science\\7TH-SEMESTER-PROJECT-";
const skillDir = "C:\\Users\\HP\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.1009.11318\\skills\\presentations";
const runtimePython = "C:\\Users\\HP\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";

process.env.RUNTIME_NODE_MODULES = "C:\\Users\\HP\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules";

const { finalizePresentation } = await import(pathToFileURL(
  path.join(skillDir, "container_tools", "artifact_tool_utils.mjs"),
).href);

const result = await finalizePresentation({
  explicitTotalSlideCount: 8,
  requiredNativeChartOwnerSlides: [3],
  workspaceDir,
  candidatePath: path.join(workspaceDir, ".chart-data-vjP03w", "candidate.pptx"),
  finalPath: path.join(workspaceDir, "deliverables", "Garbage_Classification_Deep_Learning_Project_final.pptx"),
  pythonExecutable: runtimePython,
  integrityValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-bullet-geometry", "--validate-heading-fit"],
  fontPolicy: { basis: "design", families: ["Arial"] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(workspaceDir, ".ppt-build", "Garbage_Classification_Deep_Learning_Project_final.validation.json"),
});

console.log(JSON.stringify(result, null, 2));
