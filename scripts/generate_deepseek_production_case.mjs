import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import {
  SpreadsheetFile,
  Workbook,
} from "file:///C:/Users/liuyisheng/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const outputDir = new URL("../outputs/deepseek_v4_production_case/", import.meta.url);
await fs.mkdir(outputDir, { recursive: true });

// A deterministic generator makes the acceptance case reproducible.
let seed = 20260822;
function random() {
  seed = (seed * 1664525 + 1013904223) >>> 0;
  return seed / 4294967296;
}
function pick(items) {
  return items[Math.floor(random() * items.length)];
}
function round2(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}
function pad(value, size) {
  return String(value).padStart(size, "0");
}
function datePlus(date, days) {
  const copy = new Date(date);
  copy.setUTCDate(copy.getUTCDate() + days);
  return copy;
}
function monthKey(date) {
  return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1, 2)}`;
}

const regions = ["华东", "华南", "华北", "西南", "华中", "东北"];
const channels = ["直营网店", "经销商", "电商平台", "企业团购"];
const segments = ["战略客户", "成长客户", "普通客户", "新客户"];
const surnames = ["辰光", "远航", "鑫达", "博远", "新锐", "嘉禾", "启明", "恒泰", "云帆", "瑞景", "安和", "融创", "卓越", "华盛", "佳信", "星海"];
const industries = ["商贸", "科技", "供应链", "实业", "餐饮管理", "连锁零售", "电子商务", "智能制造"];
const cityByRegion = {
  华东: ["上海", "杭州", "苏州", "南京"], 华南: ["广州", "深圳", "厦门", "佛山"],
  华北: ["北京", "天津", "石家庄", "济南"], 西南: ["成都", "重庆", "昆明", "贵阳"],
  华中: ["武汉", "长沙", "郑州", "南昌"], 东北: ["沈阳", "大连", "长春", "哈尔滨"],
};

const customerHeader = ["客户ID", "客户标准名称", "区域", "城市", "客户层级", "行业", "客户状态", "授信额度", "账期天数"];
const customers = [];
for (let i = 1; i <= 96; i += 1) {
  const region = regions[(i - 1) % regions.length];
  const city = cityByRegion[region][(i - 1) % 4];
  const company = `${city}${surnames[(i * 7) % surnames.length]}${industries[(i * 5) % industries.length]}有限公司`;
  customers.push([
    `C${pad(i, 4)}`, company, region, city, segments[(i * 3) % segments.length],
    industries[(i * 5) % industries.length], i % 19 === 0 ? "暂停合作" : "正常",
    50000 + (i % 12) * 25000, [15, 30, 45, 60][i % 4],
  ]);
}

const productHeader = ["SKU", "商品标准名称", "品类", "子品类", "标准含税价", "单位成本", "是否在售"];
const categories = ["办公设备", "办公耗材", "仓储用品", "员工福利"];
const products = [];
for (let i = 1; i <= 24; i += 1) {
  const category = categories[(i - 1) % categories.length];
  const basePrice = [1299, 169, 88, 259][(i - 1) % 4] + i * 17;
  products.push([
    `SKU-${pad(i, 3)}`, `${category}-${["标准款", "专业款", "企业款"][i % 3]}-${pad(i, 2)}`,
    category, `${category}${["A", "B", "C"][i % 3]}组`, basePrice, round2(basePrice * (0.51 + (i % 5) * 0.035)),
    i % 17 === 0 ? "停产" : "在售",
  ]);
}

function noisyCustomerName(standard, orderIndex) {
  const variants = [
    standard,
    standard.replace("有限公司", "公司"),
    standard.replace("有限公司", ""),
    standard.replace("市", ""),
    standard.replace("有限公司", "（有限）公司"),
    ` ${standard} `,
  ];
  let value = variants[orderIndex % variants.length];
  if (orderIndex % 47 === 0) value = value.replace("科技", "科枝");
  if (orderIndex % 71 === 0) value = value.replace("供应链", "供应連");
  return value;
}

const orderHeader = [
  "订单号", "下单日期", "区域月份", "客户ID", "客户名称原值", "区域", "渠道", "SKU", "商品名称原值",
  "数量", "含税单价", "折扣率", "订单金额", "业务员", "订单状态", "手机号", "邮箱",
];
const orders = [];
for (let i = 1; i <= 960; i += 1) {
  const baseDate = new Date(Date.UTC(2026, 0, 1));
  const orderDate = datePlus(baseDate, Math.floor(random() * 181));
  const customerIndex = Math.floor(random() * customers.length);
  const productIndex = Math.floor(random() * products.length);
  const customer = customers[customerIndex];
  const product = products[productIndex];
  const quantity = 1 + Math.floor(random() * (i % 31 === 0 ? 85 : 18));
  const discount = [0, 0.03, 0.05, 0.08, 0.1, 0.15][Math.floor(random() * 6)];
  let amount = round2(quantity * product[4] * (1 - discount));
  let status = pick(["已完成", "已完成", "已完成", "已发货", "待付款"]);
  if (i % 89 === 0) { status = "已退款"; amount = -amount; }
  if (i === 311 || i === 742) amount = round2(amount * 12); // deliberate extreme-value errors
  const customerId = i % 53 === 0 ? "" : customer[0];
  const region = i % 79 === 0 ? "" : customer[2];
  const salesperson = `${["王", "李", "张", "陈", "赵", "周"][customerIndex % 6]}${["敏", "伟", "磊", "静"][i % 4]}`;
  orders.push([
    `SO2026${pad(i, 6)}`, orderDate, `${monthKey(orderDate)}|${customer[2]}`, customerId,
    noisyCustomerName(customer[1], i), region, channels[(i + customerIndex) % channels.length], product[0],
    i % 67 === 0 ? `${product[1]}(旧称)` : product[1], quantity, product[4], discount, amount, salesperson,
    status, `138${pad((10000000 + i * 7919) % 100000000, 8)}`, `buyer${pad(i, 4)}@example.test`,
  ]);
}
// Eight exact duplicates are included to test de-duplication without damaging the originals.
for (const idx of [12, 118, 257, 399, 521, 688, 804, 912]) orders.push([...orders[idx]]);

const paymentHeader = ["回款流水号", "订单号", "到账日期", "付款方原值", "回款金额", "支付渠道", "银行交易号", "备注"];
const payments = [];
let paymentCounter = 1;
for (let i = 0; i < orders.length - 8; i += 1) {
  const order = orders[i];
  if (!(["已完成", "已发货", "已退款"].includes(order[14]))) continue;
  if (i % 23 === 0) continue; // deliberately unpaid
  const days = i % 61 === 0 ? 18 : 1 + (i % 5);
  let amount = order[12];
  if (i % 37 === 0) amount = round2(amount + 0.03); // within tolerance
  if (i % 83 === 0) amount = round2(amount - 19.8); // outside tolerance
  const payer = noisyCustomerName(customers.find((row) => row[0] === order[3])?.[1] ?? order[4], i + 13);
  const split = i % 17 === 0 && amount > 0;
  const chunks = split ? [round2(amount * 0.6), round2(amount - round2(amount * 0.6))] : [amount];
  for (let j = 0; j < chunks.length; j += 1) {
    payments.push([
      `PAY-${pad(paymentCounter, 7)}`, order[0], datePlus(order[1], days + j), payer, chunks[j],
      ["银行转账", "支付宝企业付", "微信商户", "银企直连"][(i + j) % 4],
      `TXN-${pad(paymentCounter * 13, 10)}`, split ? `第${j + 1}笔分期` : "正常回款",
    ]);
    paymentCounter += 1;
  }
}
// Add unmatched bank entries and duplicated transaction rows.
for (let i = 1; i <= 7; i += 1) {
  payments.push([`PAY-${pad(paymentCounter, 7)}`, `SO-UNKNOWN-${i}`, new Date(Date.UTC(2026, 6, i)), `未知付款方${i}`, 800 + i * 137.19, "银行转账", `TXN-UNMATCH-${i}`, "待认领款"]);
  paymentCounter += 1;
}
for (const idx of [19, 106, 245, 417]) payments.push([...payments[idx]]);

const targetHeader = ["区域月份", "月份", "区域", "销售目标", "回款率目标", "毛利率目标"];
const targets = [];
for (let month = 1; month <= 6; month += 1) {
  for (let r = 0; r < regions.length; r += 1) {
    const key = `2026-${pad(month, 2)}|${regions[r]}`;
    targets.push([key, new Date(Date.UTC(2026, month - 1, 1)), regions[r], 270000 + month * 18000 + r * 22000, 0.93 + (r % 3) * 0.01, 0.32 + (r % 2) * 0.015]);
  }
}

const issueHeader = ["验收项", "设计值", "业务含义", "期望处理原则"];
const issueRows = [
  ["订单原始行数", orders.length, "含 8 行完全重复", "清洗后应保留 960 行"],
  ["客户主数据", customers.length, "标准客户维表", "模糊匹配只给建议，歧义项进入人工核验"],
  ["订单金额极端异常", 2, "疑似录入多一位", "只标记，不擅自篡改"],
  ["未付订单（抽样规则）", "约 4%", "已完成/已发货但没有回款", "生成差异明细并保留解释"],
  ["容差内金额差", "0.03 元", "银行尾差", "可按 0.05 元容差自动通过"],
  ["容差外金额差", "19.80 元", "短款或手续费未说明", "必须进入异常/核验"],
  ["超期回款", "18 天", "超过默认 7 天日期窗口", "必须单列"],
  ["待认领款", 7, "订单号无法匹配", "不得强行匹配"],
  ["重复银行流水", 4, "同一笔交易重复导入", "隔离重复键，避免重复计款"],
];

const workbook = Workbook.create();
const cover = workbook.worksheets.add("案例说明");
const orderSheet = workbook.worksheets.add("订单明细");
const paymentSheet = workbook.worksheets.add("回款流水");
const customerSheet = workbook.worksheets.add("客户主数据");
const productSheet = workbook.worksheets.add("商品主数据");
const targetSheet = workbook.worksheets.add("区域月度目标");
const checkSheet = workbook.worksheets.add("验收口径");

function writeDataSheet(sheet, headers, rows, tableName, widths = {}) {
  sheet.getRangeByIndexes(0, 0, rows.length + 1, headers.length).values = [headers, ...rows];
  const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
  headerRange.format = {
    fill: "#16324F",
    font: { bold: true, color: "#FFFFFF", size: 10 },
    rowHeight: 28,
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "medium", color: "#16324F" },
  };
  const body = sheet.getRangeByIndexes(1, 0, rows.length, headers.length);
  body.format = {
    font: { color: "#243447", size: 9 },
    rowHeight: 20,
    verticalAlignment: "center",
    borders: { insideHorizontal: { style: "hair", color: "#E5EAF0" } },
  };
  sheet.tables.add(`A1:${String.fromCharCode(64 + headers.length)}${rows.length + 1}`, true, tableName).style = "TableStyleMedium2";
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
  for (let c = 0; c < headers.length; c += 1) {
    sheet.getRangeByIndexes(0, c, rows.length + 1, 1).format.columnWidth = widths[c] ?? 15;
  }
}

writeDataSheet(orderSheet, orderHeader, orders, "OrdersTable", {0: 18, 1: 13, 2: 16, 3: 12, 4: 31, 5: 10, 6: 13, 7: 12, 8: 26, 9: 9, 10: 13, 11: 10, 12: 15, 13: 10, 14: 11, 15: 16, 16: 24});
writeDataSheet(paymentSheet, paymentHeader, payments, "PaymentsTable", {0: 18, 1: 18, 2: 13, 3: 31, 4: 15, 5: 15, 6: 20, 7: 16});
writeDataSheet(customerSheet, customerHeader, customers, "CustomersTable", {0: 12, 1: 31, 2: 10, 3: 12, 4: 13, 5: 16, 6: 12, 7: 14, 8: 12});
writeDataSheet(productSheet, productHeader, products, "ProductsTable", {0: 12, 1: 27, 2: 15, 3: 18, 4: 15, 5: 14, 6: 12});
writeDataSheet(targetSheet, targetHeader, targets, "TargetsTable", {0: 16, 1: 13, 2: 10, 3: 15, 4: 14, 5: 14});
writeDataSheet(checkSheet, issueHeader, issueRows, "AcceptanceTable", {0: 25, 1: 16, 2: 30, 3: 38});

orderSheet.getRange(`B2:B${orders.length + 1}`).format.numberFormat = "yyyy-mm-dd";
orderSheet.getRange(`K2:K${orders.length + 1}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
orderSheet.getRange(`L2:L${orders.length + 1}`).format.numberFormat = "0.0%";
orderSheet.getRange(`M2:M${orders.length + 1}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
paymentSheet.getRange(`C2:C${payments.length + 1}`).format.numberFormat = "yyyy-mm-dd";
paymentSheet.getRange(`E2:E${payments.length + 1}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
customerSheet.getRange(`H2:H${customers.length + 1}`).format.numberFormat = "¥#,##0;[Red](¥#,##0);-";
productSheet.getRange(`E2:F${products.length + 1}`).format.numberFormat = "¥#,##0.00;[Red](¥#,##0.00);-";
targetSheet.getRange(`B2:B${targets.length + 1}`).format.numberFormat = "yyyy-mm";
targetSheet.getRange(`D2:D${targets.length + 1}`).format.numberFormat = "¥#,##0;[Red](¥#,##0);-";
targetSheet.getRange(`E2:F${targets.length + 1}`).format.numberFormat = "0.0%";

cover.showGridLines = false;
cover.getRange("A1:J2").merge();
cover.getRange("A1").values = [["华辰商贸集团｜2026 上半年订单—回款经营诊断案例"]];
cover.getRange("A1:J2").format = { fill: "#102A43", font: { bold: true, color: "#FFFFFF", size: 20 }, verticalAlignment: "center", horizontalAlignment: "left" };
cover.getRange("A4:J4").merge();
cover.getRange("A4").values = [["用于验证：自然语言理解 → 安全计划 → 本地自动执行 → 结果验收（数据全部为虚构）"]];
cover.getRange("A4:J4").format = { fill: "#D9EAF7", font: { color: "#16324F", italic: true, size: 11 }, verticalAlignment: "center" };
cover.getRange("A6:B9").values = [
  ["数据域", "规模"], ["订单明细", `${orders.length} 行`], ["回款流水", `${payments.length} 行`], ["主数据与目标", `${customers.length + products.length + targets.length} 行`],
];
cover.getRange("D6:J9").values = [
  ["一段话任务", "执行边界", "验收重点", "隐私", "版本", "期间", "状态"],
  ["清洗 + 验收 + 模糊匹配 + 对账", "不执行代码/宏", "金额与日期容差", "仅本地数据处理", "V4 案例", "2026H1", "待运行"],
  ["趋势 + 透视 + 贡献 + 异常", "不猜缺失参数", "待认领与重复流水", "模型仅看结构", "难度：高", "6 个月", "待运行"],
  ["RFM + 目标差异", "高风险项人工确认", "结果可追溯", "字段名可能敏感", "可复现", "6 区域", "待运行"],
];
cover.getRange("A6:J6").format = { fill: "#2F855A", font: { bold: true, color: "#FFFFFF" } };
cover.getRange("A6:J9").format.borders = { preset: "outside", style: "thin", color: "#A0AEC0" };
cover.getRange("A11:J11").merge();
cover.getRange("A11").values = [["建议给 AI 的高难度指令"]];
cover.getRange("A11:J11").format = { fill: "#F6AD55", font: { bold: true, color: "#5F370E", size: 12 } };
cover.getRange("A12:J15").merge();
cover.getRange("A12").values = [["把订单明细去掉完全重复行并清理文本空格；检查订单号唯一、客户名称和金额非空、数量 1–100、折扣率 0–0.3；按客户名称原值与客户主数据做 88% 模糊匹配；订单和回款按订单号对账，金额容差 0.05 元、日期窗口 7 天，拆分回款不要强行自动通过；然后生成区域月份销售汇总、区域×渠道销售透视、订单金额 IQR 异常明细、客户 RFM 分群，并把月度区域汇总与目标表按区域月份匹配。任何歧义、重复流水、待认领款和容差外差异都送人工核验，不要改原表。"]];
cover.getRange("A12:J15").format = { fill: "#FFF8E8", font: { color: "#4A3B20", size: 11 }, wrapText: true, verticalAlignment: "top", borders: { preset: "outside", style: "thin", color: "#F6AD55" } };
cover.getRange("A17:J17").merge();
cover.getRange("A17").values = [["注意：真实业务文件先脱敏并取得客户许可；本案例所有公司、账号、联系方式均为程序生成。"]];
cover.getRange("A17:J17").format = { fill: "#FEE2E2", font: { bold: true, color: "#9B2C2C" }, wrapText: true };
cover.getRange("A1:J18").format.columnWidth = 16;
cover.getRange("A1:A18").format.columnWidth = 22;
cover.getRange("B1:B18").format.columnWidth = 18;
cover.getRange("D1:D18").format.columnWidth = 24;
cover.getRange("A12:J15").format.rowHeight = 28;
cover.freezePanes.freezeRows(4);

// Conditional formatting makes risk cells easy to scan when the workbook is opened.
orderSheet.getRange(`M2:M${orders.length + 1}`).conditionalFormats.add("cellIs", {
  operator: "lessThan", formula: 0, format: { font: { color: "#C53030", bold: true }, fill: "#FFF5F5" },
});
paymentSheet.getRange(`E2:E${payments.length + 1}`).conditionalFormats.add("cellIs", {
  operator: "lessThan", formula: 0, format: { font: { color: "#C53030", bold: true }, fill: "#FFF5F5" },
});

const previewSheets = ["案例说明", "订单明细", "回款流水", "客户主数据", "商品主数据", "区域月度目标", "验收口径"];
for (const sheetName of previewSheets) {
  const rendered = await workbook.render({ sheetName, autoCrop: "all", scale: sheetName === "案例说明" ? 1 : 0.75, format: "png" });
  const safeName = sheetName.replaceAll("/", "_");
  await fs.writeFile(new URL(`preview_${safeName}.png`, outputDir), new Uint8Array(await rendered.arrayBuffer()));
}

const inspection = await workbook.inspect({ kind: "workbook,sheet,table,formula", maxChars: 12000, tableMaxRows: 3, tableMaxCols: 8 });
await fs.writeFile(new URL("inspection.ndjson", outputDir), inspection.ndjson ?? String(inspection), "utf8");

const output = await SpreadsheetFile.exportXlsx(workbook);
const outputUrl = new URL("华辰商贸_2026H1订单回款经营诊断_高难度案例.xlsx", outputDir);
await output.save(fileURLToPath(outputUrl));

console.log(JSON.stringify({
  output: fileURLToPath(outputUrl),
  orders: orders.length,
  payments: payments.length,
  customers: customers.length,
  products: products.length,
  targets: targets.length,
  previews: previewSheets.length,
}));
