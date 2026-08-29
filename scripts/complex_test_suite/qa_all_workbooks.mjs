import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import sharp from "sharp";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const here = path.dirname(fileURLToPath(import.meta.url));
const projectDir = path.resolve(here, "..", "..");
const outputDir = path.join(projectDir, "outputs", "complex_test_suite_20260822");
const qaDir = path.join(outputDir, "qa_all");
await fs.mkdir(qaDir, { recursive: true });

const workbooks = [
  ["case01_input", "案例01_订单回款经营诊断_输入.xlsx"],
  ["case01_result", "案例01_订单回款经营诊断_标准结果.xlsx"],
  ["case02_input", "案例02_连锁库存销售异常_输入.xlsx"],
  ["case02_result", "案例02_连锁库存销售异常_标准结果.xlsx"],
  ["case03_input", "案例03_考勤薪资核验脱敏_输入.xlsx"],
  ["case03_result", "案例03_考勤薪资核验脱敏_标准结果.xlsx"],
];

function colName(index) {
  let value = index + 1;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function escapeXml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

async function createContactSheets(label, previews) {
  const contactFiles = [];
  const columns = 4;
  const rows = 4;
  const tileWidth = 360;
  const tileHeight = 245;
  const pageSize = columns * rows;
  for (let page = 0; page * pageSize < previews.length; page += 1) {
    const pageItems = previews.slice(page * pageSize, (page + 1) * pageSize);
    const composites = [];
    for (let index = 0; index < pageItems.length; index += 1) {
      const item = pageItems[index];
      const resized = await sharp(item.path)
        .resize(tileWidth - 12, tileHeight - 38, { fit: "contain", background: "#FFFFFF" })
        .flatten({ background: "#FFFFFF" })
        .png()
        .toBuffer();
      const x = (index % columns) * tileWidth + 6;
      const y = Math.floor(index / columns) * tileHeight + 32;
      composites.push({ input: resized, left: x, top: y });
      const labelSvg = Buffer.from(`<svg width="${tileWidth - 12}" height="26"><rect width="100%" height="100%" fill="#16324F"/><text x="8" y="18" fill="white" font-family="Microsoft YaHei, sans-serif" font-size="13">${escapeXml(`${item.index}. ${item.sheet}`)}</text></svg>`);
      composites.push({ input: labelSvg, left: x, top: Math.floor(index / columns) * tileHeight + 4 });
    }
    const contactPath = path.join(qaDir, `${label}_contact_${String(page + 1).padStart(2, "0")}.png`);
    await sharp({ create: { width: columns * tileWidth, height: rows * tileHeight, channels: 4, background: "#EDF2F7" } })
      .composite(composites)
      .png()
      .toFile(contactPath);
    contactFiles.push(contactPath);
  }
  return contactFiles;
}

const summary = [];
for (const [label, filename] of workbooks) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(outputDir, filename)));
  const previewDir = path.join(qaDir, label);
  await fs.mkdir(previewDir, { recursive: true });
  const sheets = workbook.worksheets.items.map((sheet) => sheet.name);
  const previews = [];
  for (let index = 0; index < sheets.length; index += 1) {
    const sheetName = sheets[index];
    const sheet = workbook.worksheets.getItem(sheetName);
    const used = sheet.getUsedRange();
    const previewRange = `A1:${colName(Math.max(0, used.columnCount - 1))}${Math.min(used.rowCount, 35)}`;
    const blob = await workbook.render({ sheetName, range: previewRange, scale: 0.55, format: "png" });
    const safeName = sheetName.replace(/[\\/:*?"<>|]/g, "_");
    const previewPath = path.join(previewDir, `${String(index + 1).padStart(2, "0")}_${safeName}.png`);
    await fs.writeFile(previewPath, new Uint8Array(await blob.arrayBuffer()));
    previews.push({ index: index + 1, sheet: sheetName, path: previewPath, range: previewRange });
  }
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    maxChars: 16000,
  });
  const errorLines = (errors.ndjson ?? String(errors)).split(/\r?\n/).filter((line) => line.includes('"kind":"match"'));
  const overview = await workbook.inspect({ kind: "workbook,sheet,table", maxChars: 12000, tableMaxRows: 3, tableMaxCols: 10 });
  await fs.writeFile(path.join(previewDir, "overview.ndjson"), overview.ndjson ?? String(overview), "utf8");
  await fs.writeFile(path.join(previewDir, "formula_errors.ndjson"), errors.ndjson ?? String(errors), "utf8");
  const contactSheets = await createContactSheets(label, previews);
  summary.push({ label, filename, sheetCount: sheets.length, formulaErrorCount: errorLines.length, previews, contactSheets });
}

await fs.writeFile(path.join(qaDir, "qa_summary.json"), JSON.stringify(summary, null, 2), "utf8");
console.log(JSON.stringify(summary.map(({ label, filename, sheetCount, formulaErrorCount, contactSheets }) => ({ label, filename, sheetCount, formulaErrorCount, contactSheets })), null, 2));
