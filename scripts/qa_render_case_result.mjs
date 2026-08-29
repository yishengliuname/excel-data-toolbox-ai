import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  FileBlob,
  SpreadsheetFile,
} from "file:///C:/Users/liuyisheng/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const caseDir = path.join(scriptDir, "..", "outputs", "deepseek_v4_production_case");
const inputPath = path.join(caseDir, "华辰商贸_2026H1_AI自动执行结果.xlsx");
const previewDir = path.join(caseDir, "result_previews");
await fs.mkdir(previewDir, { recursive: true });

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheetNames = workbook.worksheets.items.map((sheet) => sheet.name);
const manifest = [];
for (let index = 0; index < sheetNames.length; index += 1) {
  const sheetName = sheetNames[index];
  const safeName = sheetName.replace(/[\\/:*?"<>|]/g, "_");
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 0.55,
    format: "png",
  });
  const filename = `${String(index + 1).padStart(2, "0")}_${safeName}.png`;
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(path.join(previewDir, filename), bytes);
  manifest.push({ index: index + 1, sheet: sheetName, filename, bytes: bytes.length });
}

const inspection = await workbook.inspect({
  kind: "workbook,sheet,table,formula,match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  maxChars: 20000,
  tableMaxRows: 4,
  tableMaxCols: 12,
  options: { maxResults: 500 },
});
const inspectionText = inspection.ndjson ?? String(inspection);
await fs.writeFile(path.join(previewDir, "inspection.ndjson"), inspectionText, "utf8");
await fs.writeFile(path.join(previewDir, "manifest.json"), JSON.stringify(manifest, null, 2), "utf8");

const formulaErrors = (inspectionText.match(/#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/g) ?? []).length;
console.log(JSON.stringify({ workbook: inputPath, sheetCount: sheetNames.length, formulaErrors, manifest }, null, 2));
