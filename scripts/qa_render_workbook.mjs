import fs from "node:fs/promises";
import path from "node:path";
import {
  FileBlob,
  SpreadsheetFile,
} from "file:///C:/Users/liuyisheng/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const [inputPath, outputDirectory] = process.argv.slice(2);
if (!inputPath || !outputDirectory) {
  throw new Error("usage: qa_render_workbook.mjs <input.xlsx> <preview-directory>");
}

await fs.mkdir(outputDirectory, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const manifest = [];
for (const [index, sheet] of workbook.worksheets.items.entries()) {
  const safeName = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 0.7,
    format: "png",
  });
  const filename = `${String(index + 1).padStart(2, "0")}_${safeName}.png`;
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(path.join(outputDirectory, filename), bytes);
  manifest.push({ sheet: sheet.name, filename, bytes: bytes.length });
}

const inspection = await workbook.inspect({
  kind: "workbook,sheet,table,formula,match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  maxChars: 12000,
  tableMaxRows: 5,
  tableMaxCols: 14,
  options: { maxResults: 300 },
});
const inspectionText = inspection.ndjson ?? String(inspection);
await fs.writeFile(path.join(outputDirectory, "inspection.ndjson"), inspectionText, "utf8");
await fs.writeFile(
  path.join(outputDirectory, "manifest.json"),
  JSON.stringify(manifest, null, 2),
  "utf8",
);
const formulaErrors = (
  inspectionText.match(/#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/g) ?? []
).length;
console.log(JSON.stringify({ sheetCount: manifest.length, formulaErrors, manifest }, null, 2));
