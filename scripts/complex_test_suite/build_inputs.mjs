import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const here = path.dirname(fileURLToPath(import.meta.url));
const projectDir = path.resolve(here, "..", "..");
const outputDir = path.join(projectDir, "outputs", "complex_test_suite_20260822");
const qaDir = path.join(outputDir, "qa_inputs");
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(qaDir, { recursive: true });

let seed = 20260822;
function random() {
  seed = (seed * 1664525 + 1013904223) >>> 0;
  return seed / 4294967296;
}
const pick = (items) => items[Math.floor(random() * items.length)];
const pad = (value, size) => String(value).padStart(size, "0");
const round2 = (value) => Math.round((value + Number.EPSILON) * 100) / 100;
function datePlus(date, days) {
  const copy = new Date(date);
  copy.setUTCDate(copy.getUTCDate() + days);
  return copy;
}
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

const palette = {
  navy: "#16324F",
  blue: "#2F5D8A",
  teal: "#167D7F",
  green: "#2F855A",
  amber: "#F6AD55",
  lightBlue: "#D9EAF7",
  lightAmber: "#FFF8E8",
  red: "#C53030",
  lightRed: "#FFF5F5",
  ink: "#243447",
  grid: "#E5EAF0",
};

function writeDataSheet(sheet, headers, rows, tableName, widths = {}) {
  const matrix = [headers, ...rows];
  sheet.getRangeByIndexes(0, 0, matrix.length, headers.length).values = matrix;
  sheet.getRangeByIndexes(0, 0, 1, headers.length).format = {
    fill: palette.navy,
    font: { bold: true, color: "#FFFFFF", size: 10 },
    rowHeight: 28,
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "medium", color: palette.navy },
  };
  if (rows.length) {
    sheet.getRangeByIndexes(1, 0, rows.length, headers.length).format = {
      font: { color: palette.ink, size: 9 },
      rowHeight: 20,
      verticalAlignment: "center",
      borders: { insideHorizontal: { style: "hair", color: palette.grid } },
    };
  }
  const endCell = `${colName(headers.length - 1)}${rows.length + 1}`;
  sheet.tables.add(`A1:${endCell}`, true, tableName).style = "TableStyleMedium2";
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
  headers.forEach((_, index) => {
    sheet.getRangeByIndexes(0, index, rows.length + 1, 1).format.columnWidth = widths[index] ?? 15;
  });
}

function writeCover(sheet, { title, subtitle, prompt, facts, boundary }) {
  sheet.showGridLines = false;
  sheet.getRange("A1:J2").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1:J2").format = {
    fill: palette.navy,
    font: { bold: true, color: "#FFFFFF", size: 20 },
    verticalAlignment: "center",
  };
  sheet.getRange("A4:J4").merge();
  sheet.getRange("A4").values = [[subtitle]];
  sheet.getRange("A4:J4").format = {
    fill: palette.lightBlue,
    font: { italic: true, color: palette.navy, size: 11 },
    verticalAlignment: "center",
  };
  sheet.getRange("A6:B10").values = [["检查维度", "设计值"], ...facts];
  sheet.getRange("A6:B6").format = { fill: palette.green, font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("A6:B10").format.borders = { preset: "outside", style: "thin", color: "#A0AEC0" };
  sheet.getRange("D6:J6").merge();
  sheet.getRange("D6").values = [["程序执行边界"]];
  sheet.getRange("D6:J6").format = { fill: palette.green, font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("D7:J10").merge();
  sheet.getRange("D7").values = [[boundary]];
  sheet.getRange("D7:J10").format = { fill: "#F1F8F4", wrapText: true, verticalAlignment: "top", font: { color: "#285943", size: 10 } };
  sheet.getRange("A12:J12").merge();
  sheet.getRange("A12").values = [["复制到“AI 一句话完成”的测试指令"]];
  sheet.getRange("A12:J12").format = { fill: palette.amber, font: { bold: true, color: "#5F370E", size: 12 } };
  sheet.getRange("A13:J18").merge();
  sheet.getRange("A13").values = [[prompt]];
  sheet.getRange("A13:J18").format = {
    fill: palette.lightAmber,
    font: { color: "#4A3B20", size: 11 },
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "outside", style: "thin", color: palette.amber },
  };
  sheet.getRange("A20:J20").merge();
  sheet.getRange("A20").values = [["操作：上传本工作簿 → 只选择业务数据表，不选择 00_测试说明/验收口径 → 粘贴上述指令 → 仅生成计划 → 核对标准话术 → 确认执行。"]];
  sheet.getRange("A20:J20").format = { fill: "#FEE2E2", font: { bold: true, color: "#9B2C2C" }, wrapText: true };
  sheet.getRange("A1:J20").format.columnWidth = 15;
  sheet.getRange("A1:A20").format.columnWidth = 24;
  sheet.getRange("B1:B20").format.columnWidth = 18;
  sheet.getRange("D1:D20").format.columnWidth = 22;
  sheet.getRange("A13:J18").format.rowHeight = 27;
  sheet.getRange("A20:J20").format.rowHeight = 34;
  sheet.freezePanes.freezeRows(4);
}

async function verifyAndExport(workbook, filename, label) {
  const sheets = workbook.worksheets.items.map((sheet) => sheet.name);
  const rendered = [];
  for (const sheetName of sheets) {
    const sheet = workbook.worksheets.getItem(sheetName);
    const used = sheet.getUsedRange();
    const previewRange = `A1:${colName(Math.max(0, used.columnCount - 1))}${Math.min(used.rowCount, 40)}`;
    const blob = await workbook.render({ sheetName, range: previewRange, scale: sheetName === "00_测试说明" ? 0.9 : 0.7, format: "png" });
    const safe = sheetName.replace(/[\\/:*?"<>|]/g, "_");
    const previewName = `${label}_${String(rendered.length + 1).padStart(2, "0")}_${safe}.png`;
    await fs.writeFile(path.join(qaDir, previewName), new Uint8Array(await blob.arrayBuffer()));
    rendered.push({ sheet: sheetName, preview: previewName });
  }
  const inspection = await workbook.inspect({
    kind: "workbook,sheet,table,formula,match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    maxChars: 18000,
    tableMaxRows: 4,
    tableMaxCols: 12,
    options: { useRegex: true, maxResults: 300 },
  });
  await fs.writeFile(path.join(qaDir, `${label}_inspection.ndjson`), inspection.ndjson ?? String(inspection), "utf8");
  const output = await SpreadsheetFile.exportXlsx(workbook);
  const outputPath = path.join(outputDir, filename);
  await output.save(outputPath);
  return { outputPath, sheets, rendered };
}

async function copyOrderCase() {
  const sourcePath = path.join(projectDir, "outputs", "deepseek_v4_production_case", "华辰商贸_2026H1订单回款经营诊断_高难度案例.xlsx");
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
  return verifyAndExport(workbook, "案例01_订单回款经营诊断_输入.xlsx", "case01");
}

function buildInventoryWorkbook() {
  const workbook = Workbook.create();
  const cover = workbook.worksheets.add("00_测试说明");
  const eastSheet = workbook.worksheets.add("华东销售出库");
  const southSheet = workbook.worksheets.add("华南销售出库");
  const receiptSheet = workbook.worksheets.add("采购入库");
  const openingSheet = workbook.worksheets.add("月初库存");
  const closingSheet = workbook.worksheets.add("月末盘点");
  const skuSheet = workbook.worksheets.add("SKU主数据");
  const checkSheet = workbook.worksheets.add("验收口径");

  const categories = ["数码配件", "居家日用", "办公用品", "食品饮料", "个护清洁", "季节商品"];
  const suppliers = ["启明供应链", "恒泰实业", "远航商贸", "嘉禾制造", "星海经贸", "云帆供应链"];
  const stores = [
    ["E01", "上海虹桥店", "华东"], ["E02", "杭州城西店", "华东"], ["E03", "南京新街口店", "华东"],
    ["E04", "苏州园区店", "华东"], ["E05", "宁波鄞州店", "华东"], ["S01", "深圳南山店", "华南"],
    ["S02", "广州天河店", "华南"], ["S03", "厦门思明店", "华南"], ["S04", "佛山禅城店", "华南"],
    ["S05", "东莞松山湖店", "华南"],
  ];
  const skuHeader = ["SKU", "商品标准名称", "品类", "供应商", "标准进价", "标准售价", "安全库存", "是否在售"];
  const skus = [];
  for (let i = 1; i <= 36; i += 1) {
    const category = categories[(i - 1) % categories.length];
    const cost = round2(18 + i * 4.7 + (i % 5) * 6.3);
    skus.push([`SKU-${pad(i, 3)}`, `${category}-${["基础款", "升级款", "畅销款"][i % 3]}-${pad(i, 2)}`, category, suppliers[i % suppliers.length], cost, round2(cost * (1.35 + (i % 4) * 0.08)), 12 + (i % 8) * 4, i % 17 === 0 ? "停产" : "在售"]);
  }

  const inventoryHeader = ["盘点月份", "门店编码", "门店名称", "区域", "SKU", "商品名称", "库存数量", "库存金额"];
  const opening = [];
  const closing = [];
  for (let storeIndex = 0; storeIndex < stores.length; storeIndex += 1) {
    for (let skuIndex = 0; skuIndex < skus.length; skuIndex += 1) {
      const store = stores[storeIndex];
      const sku = skus[skuIndex];
      const quantity = 18 + ((storeIndex * 17 + skuIndex * 11) % 85);
      opening.push([new Date(Date.UTC(2026, 3, 1)), store[0], store[1], store[2], sku[0], sku[1], quantity, round2(quantity * sku[4])]);
      const drift = ((storeIndex * 7 + skuIndex * 13) % 41) - 20;
      const closingQty = Math.max(0, quantity + drift);
      closing.push([new Date(Date.UTC(2026, 5, 30)), store[0], store[1], store[2], sku[0], sku[1], closingQty, round2(closingQty * sku[4])]);
    }
  }
  const removedClosing = closing.splice(0, 6);
  for (let i = 1; i <= 4; i += 1) {
    const store = stores[5 + i];
    closing.push([new Date(Date.UTC(2026, 5, 30)), store[0], store[1], store[2], `NEW-${pad(i, 3)}`, `新品待建档-${i}`, 20 + i * 3, 1200 + i * 100]);
  }
  for (const index of [44, 109, 188, 277, 330]) closing.push([...closing[index]]);

  const salesHeader = ["出库单号", "销售日期", "月份", "门店编码", "门店名称", "区域", "SKU", "商品名称原值", "销售数量", "销售单价", "折扣率", "销售金额", "渠道", "会员编号"];
  const eastSales = [];
  const southSales = [];
  for (let i = 1; i <= 1440; i += 1) {
    const date = datePlus(new Date(Date.UTC(2026, 3, 1)), Math.floor(random() * 91));
    const storeIndex = i <= 720 ? Math.floor(random() * 5) : 5 + Math.floor(random() * 5);
    const store = stores[storeIndex];
    const sku = skus[Math.floor(random() * skus.length)];
    let qty = 1 + Math.floor(random() * 18);
    if (i % 43 === 0) qty = -1 * (1 + (i % 3));
    if ([317, 988, 1331].includes(i)) qty = 420 + (i % 19);
    const discount = [0, 0.03, 0.05, 0.1, 0.15][i % 5];
    const amount = round2(qty * sku[5] * (1 - discount));
    const noisyName = i % 67 === 0 ? sku[1].replace("升级款", "升級款") : i % 41 === 0 ? ` ${sku[1]} ` : sku[1];
    const row = [`OUT-${pad(i, 7)}`, date, `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1, 2)}`, store[0], store[1], store[2], sku[0], noisyName, qty, sku[5], discount, amount, pick(["门店零售", "小程序", "团购", "到家平台"]), `M${pad(1 + (i * 37) % 620, 5)}`];
    (i <= 720 ? eastSales : southSales).push(row);
  }
  for (const index of [12, 78, 166, 255, 399, 511, 608, 690, 715]) eastSales.push([...eastSales[index]]);
  for (const index of [20, 105, 211, 318, 466, 587, 701]) southSales.push([...southSales[index]]);

  const receiptHeader = ["入库单号", "入库日期", "门店编码", "门店名称", "区域", "SKU", "入库数量", "采购单价", "含税金额", "供应商原值", "状态"];
  const receipts = [];
  for (let i = 1; i <= 620; i += 1) {
    const date = datePlus(new Date(Date.UTC(2026, 3, 1)), Math.floor(random() * 91));
    const store = stores[Math.floor(random() * stores.length)];
    const sku = skus[Math.floor(random() * skus.length)];
    let quantity = 10 + Math.floor(random() * 90);
    if (i % 127 === 0) quantity = -5;
    if (i % 149 === 0) quantity = 900;
    const skuValue = i % 79 === 0 ? "" : sku[0];
    const price = round2(sku[4] * (1 + ((i % 7) - 3) * 0.008));
    receipts.push([`IN-${pad(i, 7)}`, date, store[0], store[1], store[2], skuValue, quantity, price, round2(quantity * price), i % 53 === 0 ? `${sku[3]}公司` : sku[3], i % 31 === 0 ? "待质检" : "已入库"]);
  }
  for (const index of [14, 99, 208, 333, 444, 579]) receipts.push([...receipts[index]]);

  const prompt = "老板让我看看二季度十家门店的库存和销售是不是有问题。把华东、华南两张销售出库合起来，先清理完全重复行；检查出库单号、门店编码和 SKU 不能为空且出库单号唯一，销售数量允许退货但只能在 -20 到 500，销售金额在 -10 万到 100 万。用 SKU 主数据补齐品类、供应商、标准进价和安全库存；按月份和区域看销售趋势，做品类×门店透视、SKU 销售贡献和 IQR 异常交易。采购入库按门店+SKU 汇总；月初库存和月末盘点按门店编码+SKU 比对，新增、缺失、数量或金额变化、重复键都分开列。任何新品未建档、重复库存键、负数采购和异常大单都只标记，不要自动修正原表。";
  writeCover(cover, {
    title: "案例 02｜连锁门店库存、采购与销售异常诊断",
    subtitle: "验证：多表追加 → 清洗验收 → 主数据查找 → 趋势/透视/贡献/异常 → 新旧库存比对",
    prompt,
    facts: [["销售原始行数", eastSales.length + southSales.length], ["完全重复销售", 16], ["库存删除/新增键", `${removedClosing.length} / 4`], ["月末重复库存键", 5]],
    boundary: "退货负数量不等于错误；只有超出授权范围、缺关键键、主数据未建档、重复库存键和异常大单进入核验。程序不得自行推导库存损耗责任，也不得覆盖原始盘点。",
  });
  writeDataSheet(eastSheet, salesHeader, eastSales, "EastSalesTable", {0: 17, 1: 13, 2: 10, 3: 12, 4: 18, 5: 10, 6: 12, 7: 27, 8: 11, 9: 13, 10: 10, 11: 15, 12: 13, 13: 13});
  writeDataSheet(southSheet, salesHeader, southSales, "SouthSalesTable", {0: 17, 1: 13, 2: 10, 3: 12, 4: 18, 5: 10, 6: 12, 7: 27, 8: 11, 9: 13, 10: 10, 11: 15, 12: 13, 13: 13});
  writeDataSheet(receiptSheet, receiptHeader, receipts, "ReceiptsTable", {0: 17, 1: 13, 2: 12, 3: 18, 4: 10, 5: 12, 6: 11, 7: 13, 8: 15, 9: 20, 10: 12});
  writeDataSheet(openingSheet, inventoryHeader, opening, "OpeningInventoryTable", {0: 13, 1: 12, 2: 18, 3: 10, 4: 12, 5: 27, 6: 12, 7: 15});
  writeDataSheet(closingSheet, inventoryHeader, closing, "ClosingInventoryTable", {0: 13, 1: 12, 2: 18, 3: 10, 4: 12, 5: 27, 6: 12, 7: 15});
  writeDataSheet(skuSheet, skuHeader, skus, "SkuMasterTable", {0: 12, 1: 27, 2: 15, 3: 20, 4: 13, 5: 13, 6: 12, 7: 11});
  writeDataSheet(checkSheet, ["验收项", "设计值", "预期处理"], [
    ["两区销售合并", eastSales.length + southSales.length, "合并后保留来源，再去除 16 行完全重复"],
    ["销售数量极端值", 3, "IQR/范围规则标记，不自动改写"],
    ["采购负数", 4, "质量失败或异常明细"],
    ["采购 SKU 缺失", 7, "质量失败并进入核验"],
    ["库存键删除", 6, "旧有新无"], ["库存键新增", 4, "旧无新有"], ["月末重复键", 5, "重复键隔离"],
  ], "InventoryAcceptanceTable", {0: 23, 1: 18, 2: 42});

  for (const sheet of [eastSheet, southSheet]) {
    sheet.getRange(`B2:B${sheet.getUsedRange().rowCount}`).format.numberFormat = "yyyy-mm-dd";
    sheet.getRange(`J2:J${sheet.getUsedRange().rowCount}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
    sheet.getRange(`K2:K${sheet.getUsedRange().rowCount}`).format.numberFormat = "0.0%";
    sheet.getRange(`L2:L${sheet.getUsedRange().rowCount}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
    sheet.getRange(`I2:I${sheet.getUsedRange().rowCount}`).conditionalFormats.add("cellIs", { operator: "lessThan", formula: 0, format: { font: { color: palette.red, bold: true }, fill: palette.lightRed } });
  }
  receiptSheet.getRange(`B2:B${receipts.length + 1}`).format.numberFormat = "yyyy-mm-dd";
  receiptSheet.getRange(`H2:I${receipts.length + 1}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
  for (const sheet of [openingSheet, closingSheet]) {
    sheet.getRange(`A2:A${sheet.getUsedRange().rowCount}`).format.numberFormat = "yyyy-mm-dd";
    sheet.getRange(`H2:H${sheet.getUsedRange().rowCount}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
  }
  skuSheet.getRange(`E2:F${skus.length + 1}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
  return workbook;
}

function buildHrWorkbook() {
  const workbook = Workbook.create();
  const cover = workbook.worksheets.add("00_测试说明");
  const masterSheet = workbook.worksheets.add("员工主数据");
  const attendanceSheet = workbook.worksheets.add("2026-07考勤明细");
  const juneSheet = workbook.worksheets.add("2026-06薪资");
  const julySheet = workbook.worksheets.add("2026-07薪资");
  const budgetSheet = workbook.worksheets.add("部门薪资预算");
  const checkSheet = workbook.worksheets.add("验收口径");

  const departments = ["销售一部", "销售二部", "供应链", "客服中心", "产品研发", "财务部", "人力资源", "门店运营"];
  const cities = ["上海", "杭州", "南京", "深圳", "广州", "成都", "武汉", "北京"];
  const surname = ["王", "李", "张", "刘", "陈", "杨", "赵", "黄", "周", "吴"];
  const given = ["晨", "宇", "欣", "敏", "浩", "婷", "凯", "宁", "磊", "雪", "博", "静"];
  const masterHeader = ["员工编号", "姓名", "身份证号", "手机号", "邮箱", "部门", "工作城市", "入职日期", "在职状态", "银行卡号"];
  const employees = [];
  for (let i = 1; i <= 180; i += 1) {
    const dept = departments[(i - 1) % departments.length];
    const city = cities[(i * 3) % cities.length];
    employees.push([
      `E${pad(i, 4)}`, `${surname[i % surname.length]}${given[(i * 5) % given.length]}${i % 13 === 0 ? "·" : ""}`,
      `310101${19800101 + (i % 25) * 10000 + (i % 12) * 100 + (i % 27)}${pad((i * 37) % 10000, 4)}`.slice(0, 18),
      `139${pad((10000000 + i * 9277) % 100000000, 8)}`, `employee${pad(i, 4)}@example.test`, dept, city,
      new Date(Date.UTC(2017 + (i % 9), i % 12, 1 + (i % 27))), i <= 3 ? "已离职" : "在职", `622202${pad((100000000000 + i * 7919) % 1000000000000, 12)}`,
    ]);
  }

  const attendanceHeader = ["员工编号", "考勤日期", "班次", "应出勤小时", "实际出勤小时", "加班小时", "迟到分钟", "缺勤小时", "请假类型", "数据来源"];
  const attendance = [];
  const weekdays = [];
  for (let day = 1; day <= 31; day += 1) {
    const date = new Date(Date.UTC(2026, 6, day));
    if (![0, 6].includes(date.getUTCDay())) weekdays.push(date);
  }
  let attendanceIndex = 0;
  for (const employee of employees) {
    for (const date of weekdays) {
      attendanceIndex += 1;
      let actual = 8;
      let overtime = attendanceIndex % 9 === 0 ? 2 : attendanceIndex % 23 === 0 ? 4 : 0;
      let late = attendanceIndex % 41 === 0 ? 18 : 0;
      let absence = attendanceIndex % 97 === 0 ? 4 : 0;
      let leaveType = absence ? "事假" : "";
      if (attendanceIndex % 701 === 0) actual = 16.5;
      if (attendanceIndex % 953 === 0) actual = -1;
      if (attendanceIndex % 607 === 0) late = 420;
      if (absence) actual = 4;
      const employeeId = attendanceIndex % 457 === 0 ? "" : employee[0];
      attendance.push([employeeId, date, attendanceIndex % 29 === 0 ? "晚班" : "标准班", 8, actual, overtime, late, absence, leaveType, attendanceIndex % 11 === 0 ? "门禁补录" : "考勤机"]);
    }
  }
  for (const index of [15, 188, 407, 712, 1011, 1399, 1755, 2190, 2777, 3188, 3566, 3988]) attendance.push([...attendance[index]]);

  const salaryHeader = ["员工编号", "姓名", "基本工资", "岗位津贴", "加班工资", "考勤扣款", "绩效奖金", "社保公积金", "个税", "实发工资", "银行卡号", "版本时间"];
  function salaryRow(employee, monthIndex, revision = 0) {
    const i = Number(employee[0].slice(1));
    const base = 5200 + (i % 12) * 430 + (departments.indexOf(employee[5]) % 3) * 700;
    const allowance = 400 + (i % 5) * 180;
    const overtimePay = round2((i % 7) * 115.5 + monthIndex * 20);
    const deduction = i % 29 === 0 ? 360 : i % 17 === 0 ? 120 : 0;
    const performance = 500 + (i % 9) * 260 + revision * 100;
    const social = round2(base * 0.105);
    const tax = round2(Math.max(0, (base + allowance + overtimePay + performance - social - 5000) * 0.1));
    let net = round2(base + allowance + overtimePay - deduction + performance - social - tax);
    if (monthIndex === 7 && [44, 119, 166].includes(i)) net = round2(net * 4.8);
    if (monthIndex === 7 && [57, 148].includes(i)) net = -850;
    return [employee[0], employee[1], base, allowance, overtimePay, deduction, performance, social, tax, net, employee[9], new Date(Date.UTC(2026, monthIndex - 1, 25 + revision))];
  }
  const june = employees.slice(0, 170).map((employee) => salaryRow(employee, 6));
  const july = employees.slice(3).map((employee) => salaryRow(employee, 7));
  for (const employeeIndex of [21, 54, 88, 121, 150]) july.push(salaryRow(employees[employeeIndex], 7, 1));

  const budgetHeader = ["月份", "部门", "预算人数", "实发工资预算", "加班工资预算", "负责人"];
  const budgets = departments.map((department, index) => [new Date(Date.UTC(2026, 6, 1)), department, 24, 235000 + index * 12000, 16000 + index * 900, `${surname[index]}经理`]);

  const prompt = "帮我核验 7 月考勤和工资。考勤先去掉完全重复行，检查员工编号非空、实际出勤 0–16 小时、加班 0–12 小时、迟到 0–360 分钟、缺勤 0–8 小时；按员工编号汇总实际出勤、加班、迟到和缺勤，再用员工主数据补姓名、部门、城市。7 月工资按员工编号去重，只保留版本时间最新的一条，检查员工编号唯一、实发工资 0–10 万；补齐部门、身份证、手机号、邮箱和银行卡。按部门汇总实发工资、加班工资和人数，再与部门薪资预算按部门匹配；把 6 月和 7 月工资按员工编号比对，列出新增、离职、工资变化和重复键；对实发工资做 IQR 异常检测和数值相关性分析。最后输出一份姓名、身份证、手机号、邮箱和银行卡都脱敏的交付表。缺员工编号、超长工时、负工资、异常高薪和重复版本必须人工核验，不得自动修正。";
  writeCover(cover, {
    title: "案例 03｜考勤、薪资、预算核验与隐私脱敏",
    subtitle: "验证：大表清洗验收 → 员工级汇总 → 主数据查找 → 跨月比对 → 异常/相关性 → 多字段脱敏",
    prompt,
    facts: [["员工主数据", employees.length], ["考勤原始行数", attendance.length], ["7 月工资版本重复", 5], ["工资极端/负值", "3 / 2"]],
    boundary: "工资差异只做识别，不推断责任或擅自调薪；考勤异常、缺员工编号、重复工资版本、离职/新增人员及预算差异均需 HR 人工确认。所有姓名与账户信息均为虚构。",
  });
  writeDataSheet(masterSheet, masterHeader, employees, "EmployeeMasterTable", {0: 12, 1: 13, 2: 21, 3: 16, 4: 27, 5: 14, 6: 12, 7: 13, 8: 11, 9: 22});
  writeDataSheet(attendanceSheet, attendanceHeader, attendance, "AttendanceTable", {0: 12, 1: 13, 2: 11, 3: 13, 4: 14, 5: 12, 6: 12, 7: 12, 8: 12, 9: 14});
  writeDataSheet(juneSheet, salaryHeader, june, "JunePayrollTable", {0: 12, 1: 13, 2: 13, 3: 13, 4: 13, 5: 13, 6: 13, 7: 14, 8: 12, 9: 14, 10: 22, 11: 18});
  writeDataSheet(julySheet, salaryHeader, july, "JulyPayrollTable", {0: 12, 1: 13, 2: 13, 3: 13, 4: 13, 5: 13, 6: 13, 7: 14, 8: 12, 9: 14, 10: 22, 11: 18});
  writeDataSheet(budgetSheet, budgetHeader, budgets, "PayrollBudgetTable", {0: 13, 1: 14, 2: 12, 3: 16, 4: 16, 5: 13});
  writeDataSheet(checkSheet, ["验收项", "设计值", "预期处理"], [
    ["考勤完全重复", 12, "清洗删除但保留报告"], ["考勤员工号缺失", 8, "质量失败"],
    ["实际出勤超范围", 9, "质量失败/人工核验"], ["迟到分钟超范围", 6, "质量失败/人工核验"],
    ["6 月有、7 月无", 3, "跨月比对删除项"], ["7 月新增", 10, "跨月比对新增项"],
    ["7 月重复工资版本", 5, "只保留版本时间最新一条"], ["异常高薪/负工资", "3 / 2", "异常标记，不自动改写"],
  ], "HrAcceptanceTable", {0: 24, 1: 18, 2: 42});

  masterSheet.getRange(`H2:H${employees.length + 1}`).format.numberFormat = "yyyy-mm-dd";
  attendanceSheet.getRange(`B2:B${attendance.length + 1}`).format.numberFormat = "yyyy-mm-dd";
  for (const sheet of [juneSheet, julySheet]) {
    const count = sheet.getUsedRange().rowCount;
    sheet.getRange(`C2:J${count}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
    sheet.getRange(`L2:L${count}`).format.numberFormat = "yyyy-mm-dd hh:mm";
    sheet.getRange(`J2:J${count}`).conditionalFormats.add("cellIs", { operator: "lessThan", formula: 0, format: { font: { color: palette.red, bold: true }, fill: palette.lightRed } });
  }
  budgetSheet.getRange(`A2:A${budgets.length + 1}`).format.numberFormat = "yyyy-mm";
  budgetSheet.getRange(`D2:E${budgets.length + 1}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
  return workbook;
}

const outputs = [];
outputs.push(await copyOrderCase());
outputs.push(await verifyAndExport(buildInventoryWorkbook(), "案例02_连锁库存销售异常_输入.xlsx", "case02"));
outputs.push(await verifyAndExport(buildHrWorkbook(), "案例03_考勤薪资核验脱敏_输入.xlsx", "case03"));
await fs.writeFile(path.join(outputDir, "input_manifest.json"), JSON.stringify(outputs, null, 2), "utf8");
console.log(JSON.stringify(outputs.map((item) => ({ output: item.outputPath, sheets: item.sheets.length })), null, 2));
