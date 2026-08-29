const $ = (id) => document.getElementById(id);
const state = {
  data: null, active: null, busy: false, analysis: null, chart: null,
  taskId: "",
  dashboardCharts: [],
  recipes: [], reviews: [], recipeLoadError: "", reviewLoadError: "",
  aiCapabilities: null, aiPlan: null, aiPlanToken: "", aiPlanStatus: "neutral", aiChartSpec: null
};

const endpoints = {
  state: "/api/state", upload: "/api/upload", select: "/api/select", clean: "/api/clean", columns: "/api/columns", replace: "/api/replace",
  concat: "/api/concat", join: "/api/join", compare: "/api/compare", summary: "/api/summary",
  split: "/api/split", mask: "/api/mask", export: "/api/export", reset: "/api/reset", undo: "/api/undo", redo: "/api/redo",
  demo: "/api/demo", analysis: "/api/analysis", chart: "/api/chart", anomalies: "/api/anomalies",
  pivot: "/api/pivot", rfm: "/api/rfm", analysisExport: "/api/analysis-export",
  fuzzyCluster: "/api/fuzzy-cluster", fuzzyLookup: "/api/fuzzy-lookup",
  recipes: "/api/recipes", recipeSave: "/api/recipe-save", recipeRun: "/api/recipe-run",
  validate: "/api/validate", reconcileAdvanced: "/api/reconcile-advanced",
  reviews: "/api/reviews", reviewDecision: "/api/review-decision",
  aiCapabilities: "/api/ai-capabilities", aiDiagnose: "/api/ai-diagnose", aiPlan: "/api/ai-plan", aiChartPlan: "/api/ai-chart-plan", aiUnified: "/api/ai-unified", configStatus: "/api/config-status", aiExecute: "/api/ai-execute", aiEngineering: "/api/ai-engineering"
};

const recipePresets = {
  standard_clean: {
    name: "标准清洗准备",
    description: "清理文本、空行空列和重复记录，并推断字段类型。",
    steps: [
      { operation: "clean", params: { trim_whitespace: true, normalize_blank_strings: true, drop_empty_rows: true, drop_empty_columns: true, drop_duplicates: true, infer_types: true } }
    ]
  },
  monthly_summary: {
    name: "月度汇总复跑",
    description: "按已确认字段规则完成清洗，再生成分组汇总；请根据客户表头修改字段。",
    steps: [
      { operation: "clean", params: { trim_whitespace: true, normalize_blank_strings: true, drop_empty_rows: true, drop_empty_columns: true, drop_duplicates: false, infer_types: true } },
      { operation: "summary", params: { by: ["月份", "类别"], aggregations: { "金额": "sum" } } }
    ]
  },
  reconcile_delivery: {
    name: "对账交付准备",
    description: "统一空白与字段类型，适合作为高级对账前置配方。",
    steps: [
      { operation: "clean", params: { trim_whitespace: true, normalize_blank_strings: true, drop_empty_rows: true, drop_empty_columns: true, drop_duplicates: false, infer_types: true } }
    ]
  }
};

function toast(message, type = "") {
  const el = $("toast");
  el.textContent = message;
  el.className = `toast show ${type}`;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => { el.className = "toast"; }, 3600);
}

function busy(on, message = "正在处理，请稍候…") {
  state.busy = on;
  document.querySelectorAll("button").forEach(btn => {
    if (btn.id === "modalCancel") return;
    if (btn.id === "undoBtn") btn.disabled = on || !(state.data?.can_undo);
    else if (btn.id === "redoBtn") btn.disabled = on || !(state.data?.can_redo);
    else if (btn.id === "aiExecuteBtn") btn.disabled = on || !canExecuteAiPlan();
    else btn.disabled = on;
  });
  if (on) toast(message);
}

async function api(path, options = {}) {
  const requestOptions = { ...options };
  const headers = new Headers(options.headers || {});
  if (state.taskId) headers.set("X-Task-ID", state.taskId);
  requestOptions.headers = headers;
  const response = await fetch(path, requestOptions);
  const contentType = response.headers.get("content-type") || "";
  const payload = contentType.includes("application/json") ? await response.json() : await response.text();
  if (!response.ok) throw new Error(payload.error || payload.message || payload || `请求失败 (${response.status})`);
  const returnedTaskId = payload?.task_id || payload?.state?.task_id;
  if (returnedTaskId) state.taskId = String(returnedTaskId);
  return payload;
}

async function post(path, payload) {
  return api(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"})[c]);
}

function selectedValues(select) { return [...select.selectedOptions].map(o => o.value); }
function activeTable() { return state.data?.tables?.find(t => t.id === state.data.active_table) || null; }
function columnsFor(tableId) { return state.data?.tables?.find(t => t.id === tableId)?.columns || []; }

function fillSelect(el, items, selected = null, placeholder = "请选择") {
  if (!el) return;
  const isMulti = el.multiple;
  const old = selected ?? (isMulti ? selectedValues(el) : el.value);
  const values = Array.isArray(old) ? old : [old];
  el.innerHTML = isMulti ? "" : `<option value="">${placeholder}</option>`;
  items.forEach(item => {
    const value = typeof item === "string" ? item : item.value;
    const label = typeof item === "string" ? item : item.label;
    const option = document.createElement("option");
    option.value = value; option.textContent = label;
    option.selected = values.includes(value);
    el.appendChild(option);
  });
}

function tableOptions() { return (state.data?.tables || []).map(t => ({ value: t.id, label: t.name })); }

function renderTables() {
  const tables = state.data?.tables || [];
  $("fileCount").textContent = state.data?.file_count || 0;
  $("tableCount").textContent = tables.length;
  const list = $("tableList");
  if (!tables.length) {
    list.className = "table-list empty-state-small";
    list.textContent = "尚未导入数据";
  } else {
    list.className = "table-list";
    list.innerHTML = tables.map(t => `<button class="table-item ${t.id === state.data.active_table ? "active" : ""}" data-table="${escapeHtml(t.id)}"><span class="sheet-icon">表</span><span><strong title="${escapeHtml(t.name)}">${escapeHtml(t.name)}</strong><small>${Number(t.rows).toLocaleString()} 行 · ${t.columns.length} 列</small></span></button>`).join("");
    list.querySelectorAll("[data-table]").forEach(btn => btn.addEventListener("click", () => selectTable(btn.dataset.table)));
  }
  fillSelect($("activeTable"), tableOptions(), state.data?.active_table, "暂无数据表");
  ["concatTables", "joinLeft", "joinRight", "compareBase", "compareTarget", "exportTables", "fuzzySource", "fuzzyLookup",
    "recipeTable", "validationTable", "reconcileLeft", "reconcileRight", "aiTables"].forEach(id => fillSelect($(id), tableOptions()));
  if (tables.length) {
    if (!$("joinLeft").value) $("joinLeft").value = tables[0].id;
    if (!$("joinRight").value) $("joinRight").value = tables[Math.min(1, tables.length - 1)].id;
    if (!$("compareBase").value) $("compareBase").value = tables[0].id;
    if (!$("compareTarget").value) $("compareTarget").value = tables[Math.min(1, tables.length - 1)].id;
    if (!$("fuzzySource").value) $("fuzzySource").value = tables[0].id;
    if (!$("fuzzyLookup").value) $("fuzzyLookup").value = tables[Math.min(1, tables.length - 1)].id;
    if (!$("recipeTable").value) $("recipeTable").value = state.data.active_table || tables[0].id;
    if (!$("validationTable").value) $("validationTable").value = state.data.active_table || tables[0].id;
    if (!$("reconcileLeft").value) $("reconcileLeft").value = tables[0].id;
    if (!$("reconcileRight").value) $("reconcileRight").value = tables[Math.min(1, tables.length - 1)].id;
    if (!selectedValues($("aiTables")).length) {
      const preferredTable = state.data.active_table || tables[0].id;
      [...$("aiTables").options].forEach(option => { option.selected = option.value === preferredTable; });
    }
    if (!selectedValues($("exportTables")).length) [...$("exportTables").options].forEach(o => o.selected = o.value === state.data.active_table);
  }
  updateDependentColumns();
}

function renderProfile() {
  const profile = state.data?.profile || {};
  $("metricRows").textContent = profile.rows == null ? "—" : Number(profile.rows).toLocaleString();
  $("metricColumns").textContent = profile.columns == null ? "—" : Number(profile.columns).toLocaleString();
  $("metricMissing").textContent = profile.missing_cells == null ? "—" : Number(profile.missing_cells).toLocaleString();
  $("metricDuplicates").textContent = profile.duplicate_rows == null ? "—" : Number(profile.duplicate_rows).toLocaleString();
  const banner = $("qualityBanner");
  const warnings = state.data?.warnings || [];
  banner.classList.toggle("hidden", !warnings.length);
  banner.textContent = warnings.join("；");
  const cp = $("columnProfile");
  cp.innerHTML = (profile.column_profiles || []).slice(0, 30).map(c => `<span class="profile-chip"><b>${escapeHtml(c.name)}</b> · ${escapeHtml(c.dtype)} · 空值 ${Number(c.missing || 0).toLocaleString()}</span>`).join("");
}

function renderPreview() {
  const preview = state.data?.preview;
  const table = $("previewTable");
  if (!preview?.columns?.length) {
    table.innerHTML = '<tbody><tr><td class="blank-table">导入文件后在这里预览数据</td></tr></tbody>';
    return;
  }
  const head = `<thead><tr><th class="row-index">#</th>${preview.columns.map(c => `<th title="${escapeHtml(c)}">${escapeHtml(c)}</th>`).join("")}</tr></thead>`;
  const body = `<tbody>${preview.rows.map((row, i) => `<tr><td class="row-index">${i + 1}</td>${preview.columns.map(c => `<td title="${escapeHtml(row[c])}">${escapeHtml(row[c])}</td>`).join("")}</tr>`).join("")}</tbody>`;
  table.innerHTML = head + body;
}

function renderOperations() {
  const ops = state.data?.operations || [];
  $("operationCount").textContent = `${ops.length} 步`;
  $("undoBtn").disabled = state.busy || !state.data?.can_undo;
  $("redoBtn").disabled = state.busy || !state.data?.can_redo;
  $("operationList").innerHTML = ops.length ? ops.slice().reverse().map((op, i) => `<div class="operation-item"><span class="step">${ops.length - i}</span><span><strong>${escapeHtml(op.name)}</strong><small>${escapeHtml(op.detail || "")} ${op.before_rows != null ? `· ${Number(op.before_rows).toLocaleString()} → ${Number(op.after_rows).toLocaleString()} 行` : ""}</small></span><time>${escapeHtml(op.time || "")}</time></div>`).join("") : '<div class="blank-log">处理操作将在这里显示</div>';
}

function formatMetric(value, digits = 1) {
  if (value === null || value === undefined || value === "") return "—";
  const number = Number(value);
  if (!Number.isFinite(number)) return "—";
  return number.toLocaleString("zh-CN", { maximumFractionDigits: digits });
}

function renderAnalysis(payload) {
  state.analysis = payload;
  $("analysisEmpty").classList.add("hidden");
  $("analysisResult").classList.remove("hidden");
  const quality = payload.quality || {};
  const overview = payload.overview || {};
  const score = Math.max(0, Math.min(100, Number(quality.score || 0)));
  $("qualityScore").textContent = Math.round(score);
  $("qualityRing").style.setProperty("--score", score);
  $("qualityGrade").textContent = quality.grade || "已完成体检";
  $("qualitySummary").textContent = quality.summary || "已完成数据结构与质量扫描";
  const kpis = [
    ["缺失率", `${formatMetric((overview.missing_rate || 0) * 100)}%`, `${formatMetric(overview.missing_cells || 0, 0)} 个空值`],
    ["重复率", `${formatMetric((overview.duplicate_rate || 0) * 100)}%`, `${formatMetric(overview.duplicate_rows || 0, 0)} 行重复`],
    ["数值字段", formatMetric(overview.numeric_columns || 0, 0), `日期字段 ${formatMetric(overview.date_columns || 0, 0)}`],
    ["数据规模", `${formatMetric(overview.rows || 0, 0)} 行`, `${formatMetric(overview.columns || 0, 0)} 列 · ${formatMetric(overview.memory_mb || 0, 2)} MB`]
  ];
  $("analysisKpis").innerHTML = kpis.map(item => `<article class="analysis-kpi"><span>${item[0]}</span><strong>${item[1]}</strong><small>${item[2]}</small></article>`).join("");

  const insights = payload.insights || [];
  $("insightCount").textContent = `${insights.length} 条`;
  $("insightList").innerHTML = insights.length ? insights.map((item, index) => `<div class="insight-item"><span>${index + 1}</span><div><strong>${escapeHtml(item.title || "分析发现")}</strong><p>${escapeHtml(item.detail || item.message || "")}</p></div></div>`).join("") : '<div class="empty-analysis-list">当前没有足够字段生成自动洞察，可使用下方自助图表。</div>';

  const issues = payload.issues || [];
  $("issueCount").textContent = `${issues.length} 项`;
  $("issueList").innerHTML = issues.length ? issues.map(item => {
    const severity = ["danger", "warning"].includes(item.severity) ? item.severity : "info";
    const icon = severity === "danger" ? "!" : severity === "warning" ? "△" : "i";
    const detail = [item.detail || item.message, item.recommendation].filter(Boolean).join(" · ");
    return `<div class="issue-item ${severity}"><span>${icon}</span><div><strong>${escapeHtml(item.title || item.column || "质量提示")}</strong><p>${escapeHtml(detail)}</p></div></div>`;
  }).join("") : '<div class="empty-analysis-list">未发现明显的数据质量问题。</div>';
  renderCorrelation(payload.correlations || {});
}

function renderCorrelation(data) {
  const columns = data.columns || [];
  const matrix = data.matrix || [];
  const target = $("correlationMatrix");
  if (columns.length < 2 || !matrix.length) {
    target.innerHTML = '<div class="empty-analysis-list">至少需要两个可用数值字段才能计算相关性。</div>';
    return;
  }
  const color = value => {
    if (value === null || value === undefined) return "#f1f4f7";
    const number = Number(value);
    if (!Number.isFinite(number)) return "#f1f4f7";
    const alpha = .12 + Math.abs(number) * .65;
    return number >= 0 ? `rgba(18,148,103,${alpha})` : `rgba(212,85,76,${alpha})`;
  };
  const head = `<tr><th></th>${columns.map(column => `<th title="${escapeHtml(column)}">${escapeHtml(String(column).slice(0, 9))}</th>`).join("")}</tr>`;
  const body = columns.map((column, row) => `<tr><th title="${escapeHtml(column)}">${escapeHtml(String(column).slice(0, 9))}</th>${columns.map((_, col) => {
    const value = matrix[row]?.[col];
    return `<td class="correlation-cell" style="background:${color(value)}" title="${escapeHtml(column)} × ${escapeHtml(columns[col])}: ${formatMetric(value, 3)}">${formatMetric(value, 2)}</td>`;
  }).join("")}</tr>`).join("");
  target.innerHTML = `<table class="correlation-table">${head}${body}</table>`;
}

function shortLabel(value, length = 10) {
  const text = String(value ?? "");
  return text.length > length ? `${text.slice(0, length)}…` : text;
}

function chartColors() {
  const palettes = {
    business_dark: ["#33d69f", "#52a8ff", "#f5c451", "#ff7c83", "#a98bff", "#67d5df"],
    economist: ["#e3120b", "#006ba2", "#3eBCD2", "#f5b14c", "#6b6ecf", "#888888"],
    swiss: ["#e4002b", "#111111", "#757575", "#d4d4d4", "#0067a0", "#f0a500"],
    finance: ["#0b5fa5", "#00a6a6", "#70ad47", "#ffc000", "#ed7d31", "#5b9bd5"],
    warm: ["#bc5f45", "#d89045", "#e7b75d", "#6f8b74", "#8d6e63", "#a45b73"],
    minimal: ["#20242a", "#737b84", "#a8afb6", "#d2d6da", "#496f8a", "#8d6a9f"],
    default: ["#15865c", "#0f8d86", "#3d78c5", "#7c69c7", "#cf6e7d", "#d98a39", "#94a447", "#4f9fba", "#9a6f52", "#65758b"]
  };
  const theme = palettes[state.chart?.theme] || palettes.default;
  const custom = Array.isArray(state.chart?.series_colors)
    ? state.chart.series_colors.filter(color => /^#[0-9a-f]{6}$/i.test(String(color || "")))
    : [];
  return custom.length ? [...custom, ...theme.filter(color => !custom.includes(color))] : theme;
}

function chartSeriesColor(item, index, payload = state.chart) {
  if (/^#[0-9a-f]{6}$/i.test(String(item?.color || ""))) return item.color;
  const custom = payload?.series_colors?.[index];
  if (/^#[0-9a-f]{6}$/i.test(String(custom || ""))) return custom;
  const colors = chartColors();
  return colors[index % colors.length];
}

function chartColorWithAlpha(color, alpha) {
  const match = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(color || ""));
  if (!match) return color;
  return `rgba(${parseInt(match[1],16)},${parseInt(match[2],16)},${parseInt(match[3],16)},${Math.max(0,Math.min(1,alpha))})`;
}

function chartVisualOptions(payload = state.chart) {
  const clamp = (value, min, max, fallback) => {
    const number = Number(value);
    return Number.isFinite(number) ? Math.max(min, Math.min(max, number)) : fallback;
  };
  return {
    background: /^#[0-9a-f]{6}$/i.test(String(payload?.background_color || "")) ? payload.background_color : "#FFFFFF",
    text: /^#[0-9a-f]{6}$/i.test(String(payload?.text_color || "")) ? payload.text_color : "#243831",
    fontSize: clamp(payload?.font_size, 10, 24, 12),
    rotation: clamp(payload?.label_rotation, -90, 90, 0),
    opacity: clamp(payload?.opacity, .2, 1, .92),
    gap: clamp(payload?.bar_gap, 0, .8, .22),
    height: clamp(payload?.chart_height, 240, 600, 340),
    showGrid: payload?.show_grid !== false,
    legendPosition: ["top", "bottom", "left", "right"].includes(payload?.legend_position) ? payload.legend_position : "bottom"
  };
}

function optionalChartNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function chartLegend(series, payload = state.chart) {
  if (payload?.show_legend === false || !series?.length) return "";
  const options = chartVisualOptions(payload);
  const vertical = ["left", "right"].includes(options.legendPosition);
  const items = series.map((item, index) => `<span><i style="background:${chartSeriesColor(item,index,payload)}"></i>${escapeHtml(shortLabel(item.name,16))}</span>`).join("");
  return `<div class="chart-legend" style="${vertical ? "flex-direction:column;align-items:flex-start;min-width:110px" : ""}">${items}</div>`;
}

function chartWithLegend(svg, series, payload = state.chart) {
  const legend = chartLegend(series, payload);
  if (!legend) return svg;
  const position = chartVisualOptions(payload).legendPosition;
  if (position === "top") return `<div>${legend}${svg}</div>`;
  if (position === "left") return `<div style="display:flex;align-items:center;gap:12px">${legend}<div style="min-width:0;flex:1">${svg}</div></div>`;
  if (position === "right") return `<div style="display:flex;align-items:center;gap:12px"><div style="min-width:0;flex:1">${svg}</div>${legend}</div>`;
  return `<div>${svg}${legend}</div>`;
}

function formatChartMetric(value, digits = 1, payload = state.chart) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "—";
  const mode = payload?.number_format || "auto";
  if (mode === "currency") return `¥${formatMetric(number, digits)}`;
  if (mode === "percent") return `${formatMetric(number * 100, digits)}%`;
  if (mode === "wan") return `${formatMetric(number / 10000, digits)}万`;
  if (mode === "yi") return `${formatMetric(number / 100000000, digits)}亿`;
  return formatMetric(number, digits);
}

function renderChart(payload) {
  state.chart = payload;
  document.querySelector(".chart-stage")?.setAttribute("data-theme", payload.theme || "default");
  const canvas = $("chartCanvas");
  const options = chartVisualOptions(payload);
  canvas.style.background = options.background;
  canvas.style.color = options.text;
  canvas.style.fontSize = `${options.fontSize}px`;
  canvas.style.minHeight = `${options.height}px`;
  canvas.style.setProperty("--chart-font-size", `${options.fontSize}px`);
  canvas.style.setProperty("--chart-height", `${options.height}px`);
  $("chartTitle").textContent = payload.title || "分析图表";
  $("chartBadge").textContent = payload.badge || `${(payload.labels || payload.points || []).length} 个数据点`;
  $("chartInsight").textContent = payload.summary || "";
  $("saveChartBtn").disabled = false;
  $("downloadChartBtn").disabled = false;
  const type = payload.chart_type || "bar";
  if (type === "pie") return renderPieChart(payload);
  if (type === "scatter") return renderScatterChart(payload);
  if (type === "horizontal_bar") return renderHorizontalBarChart(payload);
  if (type === "grouped_bar") return renderGroupedBarChart(payload);
  if (type === "stacked_bar") return renderStackedBarChart(payload);
  if (type === "radar") return renderRadarChart(payload);
  if (type === "funnel") return renderFunnelChart(payload);
  if (type === "waterfall") return renderWaterfallChart(payload);
  if (type === "treemap") return renderTreemapChart(payload);
  if (type === "heatmap") return renderHeatmapChart(payload);
  if (type === "box") return renderBoxChart(payload);
  if (type === "gantt") return renderGanttChart(payload);
  return renderCartesianChart(payload, type);
}

function renderCartesianChart(payload, type) {
  const labels = payload.labels || [];
  const values = (payload.values || []).map(Number);
  if (!labels.length || !values.some(Number.isFinite)) {
    $("chartCanvas").innerHTML = '<div class="chart-placeholder">没有可用于当前图表的有效数据，请调整字段或统计方式。</div>';
    return;
  }
  const options = chartVisualOptions(payload);
  const width = 760, height = options.height, left = 76, right = 18, top = 22, bottom = 78;
  const plotW = width - left - right, plotH = height - top - bottom;
  const finite = values.filter(Number.isFinite);
  const referenceValues = (payload.reference_lines || []).map(item => Number(item.value)).filter(Number.isFinite);
  let min = optionalChartNumber(payload.y_min) ?? (["line", "area"].includes(type) ? Math.min(...finite) : Math.min(0, ...finite));
  let max = optionalChartNumber(payload.y_max) ?? Math.max(0, ...finite);
  if (referenceValues.length) { min = Math.min(min, ...referenceValues); max = Math.max(max, ...referenceValues); }
  if (min === max) { min -= 1; max += 1; }
  const y = value => top + (max - value) / (max - min) * plotH;
  const baseY = y(Math.max(min, Math.min(max, 0)));
  const step = plotW / Math.max(labels.length, 1);
  const primaryColor = chartSeriesColor(null, 0, payload);
  let svg = `<svg class="chart-svg${payload.style_3d ? " is-3d" : ""}" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(payload.title || "分析图表")}" style="font-size:${options.fontSize}px;color:${options.text}"><rect width="${width}" height="${height}" fill="${options.background}"/>`;
  for (let tick = 0; tick <= 4; tick++) {
    const value = max - (max - min) * tick / 4;
    const yy = top + plotH * tick / 4;
    if (options.showGrid) svg += `<line x1="${left}" y1="${yy}" x2="${width-right}" y2="${yy}" stroke="#e4e9ed" stroke-width="1"/>`;
    svg += `<text x="${left-7}" y="${yy+3}" text-anchor="end" fill="${options.text}">${escapeHtml(formatChartMetric(value, 1, payload))}</text>`;
  }
  (payload.reference_lines || []).forEach(item => { const yy=y(Number(item.value)); svg += `<line x1="${left}" y1="${yy}" x2="${width-right}" y2="${yy}" stroke="${item.color}" stroke-width="2" stroke-dasharray="7 5"/><text x="${width-right-3}" y="${yy-5}" text-anchor="end" style="fill:${item.color};font-weight:700">${escapeHtml(item.label)} · ${escapeHtml(formatChartMetric(item.value,1,payload))}</text>`; });
  if (["line", "area"].includes(type)) {
    const points = values.map((value, index) => `${left + step * (index + .5)},${y(Number.isFinite(value) ? value : 0)}`);
    if (type === "area") {
      svg += `<path d="M${left + step * .5},${baseY} L${points.join(" L")} L${left + step * (labels.length - .5)},${baseY} Z" fill="url(#areaFill)" opacity=".72"/><defs><linearGradient id="areaFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#16865d"/><stop offset="1" stop-color="#dff5eb"/></linearGradient></defs>`;
    }
    svg += `<path d="M${points.join(" L")}" fill="none" stroke="${primaryColor}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" opacity="${options.opacity}"/>`;
    values.forEach((value, index) => {
      const x = left + step * (index + .5), yy = y(Number.isFinite(value) ? value : 0);
      const highlighted=String(labels[index])===String(payload.highlight?.value), color=highlighted?payload.highlight.color:primaryColor;
      svg += `<circle cx="${x}" cy="${yy}" r="${highlighted?6:4}" fill="#fff" stroke="${color}" stroke-width="${highlighted?3:2}"><title>${escapeHtml(labels[index])}: ${escapeHtml(formatChartMetric(value, 2, payload))}</title></circle>`;
      if(payload.show_labels) svg += `<text x="${x}" y="${yy-9}" text-anchor="middle" style="font-weight:700">${escapeHtml(formatChartMetric(value,1,payload))}</text>`;
    });
  } else {
    const barW = Math.max(3, Math.min(44, step * (1-options.gap)));
    values.forEach((value, index) => {
      if (!Number.isFinite(value)) return;
      const x = left + step * index + (step - barW) / 2, yy = y(value);
      const rectY = Math.min(yy, baseY), rectH = Math.max(1, Math.abs(baseY - yy));
      const highlighted=String(labels[index])===String(payload.highlight?.value), fill=highlighted?payload.highlight.color:primaryColor;
      svg += `<rect x="${x}" y="${rectY}" width="${barW}" height="${rectH}" rx="4" fill="${fill}" opacity="${options.opacity}"><title>${escapeHtml(labels[index])}: ${escapeHtml(formatChartMetric(value, 2, payload))}</title></rect>`;
      if(payload.show_labels) svg += `<text x="${x+barW/2}" y="${value>=0?rectY-5:rectY+rectH+12}" text-anchor="middle" style="font-weight:700">${escapeHtml(formatChartMetric(value,1,payload))}</text>`;
    });
  }
  const skip = Math.max(1, Math.ceil(labels.length / 12));
  labels.forEach((label, index) => {
    if (index % skip) return;
    const x = left + step * (index + .5);
    const tickY=height-bottom+18;
    svg += `<text x="${x}" y="${tickY}" text-anchor="middle" fill="${options.text}" transform="rotate(${options.rotation} ${x} ${tickY})">${escapeHtml(shortLabel(label, 8))}</text>`;
  });
  svg += `<text x="${left+plotW/2}" y="${height-12}" text-anchor="middle" fill="${options.text}" style="font-weight:700">${escapeHtml(payload.x_axis_label || "类别")}</text>`;
  svg += `<text x="18" y="${top+plotH/2}" text-anchor="middle" fill="${options.text}" style="font-weight:700" transform="rotate(-90 18 ${top+plotH/2})">${escapeHtml(payload.y_axis_label || "数值")}</text>`;
  svg += `</svg>`;
  $("chartCanvas").innerHTML = svg;
}

function renderHorizontalBarChart(payload) {
  const rows = (payload.labels || []).map((label, index) => ({ label, value: Number(payload.values?.[index]) })).filter(item => Number.isFinite(item.value));
  if (!rows.length) return void ($("chartCanvas").innerHTML = '<div class="chart-placeholder">横向条形图没有有效汇总数据。</div>');
  const width=760,left=150,right=65,top=18,rowH=28,height=Math.max(230,top+rows.length*rowH+28),plotW=width-left-right;
  const min=Math.min(0,...rows.map(r=>r.value)),max0=Math.max(0,...rows.map(r=>r.value)),max=max0===min?min+1:max0;
  const x=value=>left+(value-min)/(max-min)*plotW,zero=x(0);
  let svg=`<svg class="chart-svg${payload.style_3d ? " is-3d" : ""}" viewBox="0 0 ${width} ${height}">`;
  const positive=chartSeriesColor(null,0,payload),negative=chartSeriesColor(null,1,payload);
  rows.forEach((row,i)=>{const y=top+i*rowH+3,end=x(row.value),bx=Math.min(zero,end),bw=Math.max(1,Math.abs(end-zero));svg+=`<text x="${left-8}" y="${y+14}" text-anchor="end">${escapeHtml(shortLabel(row.label,18))}</text><rect x="${bx}" y="${y}" width="${bw}" height="18" rx="5" fill="${row.value<0?negative:positive}"><title>${escapeHtml(row.label)}: ${escapeHtml(formatMetric(row.value,2))}</title></rect><text x="${row.value<0?bx-5:bx+bw+5}" y="${y+14}" text-anchor="${row.value<0?'end':'start'}">${escapeHtml(formatMetric(row.value,1))}</text>`;});
  svg+=`</svg>`; $("chartCanvas").innerHTML=svg;
}

function renderPieChart(payload) {
  const pairs = (payload.labels || []).map((label, index) => ({ label, value: Number(payload.values?.[index]) })).filter(item => Number.isFinite(item.value) && item.value > 0);
  const total = pairs.reduce((sum, item) => sum + item.value, 0);
  if (!pairs.length || total <= 0) {
    $("chartCanvas").innerHTML = '<div class="chart-placeholder">占比图需要大于 0 的数值。</div>';
    return;
  }
  const colors = chartColors(), cx = 250, cy = 154, radius = 105;
  let angle = -Math.PI / 2, paths = "";
  pairs.forEach((item, index) => {
    const next = angle + item.value / total * Math.PI * 2;
    const x1 = cx + radius * Math.cos(angle), y1 = cy + radius * Math.sin(angle);
    const x2 = cx + radius * Math.cos(next), y2 = cy + radius * Math.sin(next);
    const large = next - angle > Math.PI ? 1 : 0;
    paths += `<path d="M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2} Z" fill="${colors[index % colors.length]}" stroke="#fff" stroke-width="2"><title>${escapeHtml(item.label)}: ${escapeHtml(formatMetric(item.value, 2))} (${formatMetric(item.value / total * 100)}%)</title></path>`;
    angle = next;
  });
  const options=chartVisualOptions(payload);
  const svg = `<svg class="chart-svg${payload.style_3d ? " is-3d" : ""}" viewBox="0 0 520 315">${paths}<circle cx="${cx}" cy="${cy}" r="56" fill="${options.background}"/><text x="${cx}" y="${cy-4}" text-anchor="middle">合计</text><text x="${cx}" y="${cy+18}" text-anchor="middle" style="font-size:17px;font-weight:800;fill:${options.text}">${escapeHtml(formatChartMetric(total, 1, payload))}</text></svg>`;
  const legendSeries = pairs.slice(0,10).map(item=>({name:`${shortLabel(item.label,12)} ${formatMetric(item.value/total*100)}%`}));
  $("chartCanvas").innerHTML = chartWithLegend(svg,legendSeries,payload);
}

function renderScatterChart(payload) {
  const points = (payload.points || []).map(item => ({ x: Number(item.x), y: Number(item.y), label: item.label || "" })).filter(item => Number.isFinite(item.x) && Number.isFinite(item.y));
  if (!points.length) {
    $("chartCanvas").innerHTML = '<div class="chart-placeholder">散点图需要两个可转换为数值的字段。</div>';
    return;
  }
  const options=chartVisualOptions(payload),width = 760, height = options.height, left = 76, right = 18, top = 22, bottom = 72;
  let minX = Math.min(...points.map(p => p.x)), maxX = Math.max(...points.map(p => p.x));
  let minY = Math.min(...points.map(p => p.y)), maxY = Math.max(...points.map(p => p.y));
  if (minX === maxX) { minX -= 1; maxX += 1; } if (minY === maxY) { minY -= 1; maxY += 1; }
  const x = value => left + (value - minX) / (maxX - minX) * (width-left-right);
  const y = value => top + (maxY - value) / (maxY - minY) * (height-top-bottom);
  let svg = `<svg class="chart-svg${payload.style_3d ? " is-3d" : ""}" viewBox="0 0 ${width} ${height}">`;
  for (let tick = 0; tick <= 4; tick++) {
    const xx = left + (width-left-right) * tick / 4, yy = top + (height-top-bottom) * tick / 4;
    if(options.showGrid) svg += `<line x1="${xx}" y1="${top}" x2="${xx}" y2="${height-bottom}" stroke="#e5e9ed"/><line x1="${left}" y1="${yy}" x2="${width-right}" y2="${yy}" stroke="#e5e9ed"/>`;
    svg += `<text x="${xx}" y="${height-bottom+18}" text-anchor="middle">${escapeHtml(formatMetric(minX + (maxX-minX)*tick/4, 1))}</text><text x="${left-7}" y="${yy+3}" text-anchor="end">${escapeHtml(formatMetric(maxY - (maxY-minY)*tick/4, 1))}</text>`;
  }
  points.slice(0, 1000).forEach(point => { svg += `<circle cx="${x(point.x)}" cy="${y(point.y)}" r="4" fill="${chartSeriesColor(null,0,payload)}" opacity="${options.opacity}"><title>${escapeHtml(point.label)} X=${escapeHtml(formatMetric(point.x, 2))}, Y=${escapeHtml(formatMetric(point.y, 2))}</title></circle>`; });
  const trend = payload.trendline;
  if (trend?.points?.length === 2) {
    const first = trend.points[0], last = trend.points[1];
    svg += `<line x1="${x(Number(first.x))}" y1="${y(Number(first.y))}" x2="${x(Number(last.x))}" y2="${y(Number(last.y))}" stroke="#d15f54" stroke-width="3" stroke-dasharray="7 5"><title>${escapeHtml(trend.equation || "线性回归")} · R²=${escapeHtml(formatMetric(trend.r_squared, 3))}</title></line>`;
  }
  svg += `<text x="${left+(width-left-right)/2}" y="${height-10}" text-anchor="middle" style="font-weight:700">${escapeHtml(payload.x_axis_label||"X")}</text><text x="18" y="${top+(height-top-bottom)/2}" text-anchor="middle" style="font-weight:700" transform="rotate(-90 18 ${top+(height-top-bottom)/2})">${escapeHtml(payload.y_axis_label||"Y")}</text></svg>`;
  $("chartCanvas").innerHTML = svg;
}

function renderStackedBarChart(payload) {
  const labels = payload.labels || [], series = payload.series || [];
  if (!labels.length || !series.length) return void ($("chartCanvas").innerHTML = '<div class="chart-placeholder">堆叠柱图需要横轴、系列和数值字段。</div>');
  const width = 760, height = 315, left = 62, right = 18, top = 20, bottom = 58;
  const totals = labels.map((_, i) => series.reduce((sum, item) => sum + Math.max(0, Number(item.values?.[i]) || 0), 0));
  const max = Math.max(1, ...totals), plotW = width-left-right, plotH = height-top-bottom, step = plotW / labels.length;
  const colors = series.map((item,index)=>chartSeriesColor(item,index,payload));
  let svg = `<svg class="chart-svg${payload.style_3d ? " is-3d" : ""}" viewBox="0 0 ${width} ${height}">`;
  for (let tick=0; tick<=4; tick++) { const yy=top+plotH*tick/4, value=max*(1-tick/4); svg += `<line x1="${left}" y1="${yy}" x2="${width-right}" y2="${yy}" stroke="#e4e9ed"/><text x="${left-7}" y="${yy+3}" text-anchor="end">${escapeHtml(formatMetric(value,1))}</text>`; }
  labels.forEach((label, i) => {
    let cumulative = 0; const barW = Math.max(5, Math.min(42, step*.65)), bx=left+step*i+(step-barW)/2;
    series.forEach((item, si) => { const value=Math.max(0, Number(item.values?.[i])||0), h=value/max*plotH, yy=top+plotH-(cumulative+value)/max*plotH; cumulative += value; svg += `<rect x="${bx}" y="${yy}" width="${barW}" height="${Math.max(0,h)}" fill="${colors[si%colors.length]}"><title>${escapeHtml(label)} · ${escapeHtml(item.name)}: ${escapeHtml(formatMetric(value,2))}</title></rect>`; });
    svg += `<text x="${left+step*(i+.5)}" y="${height-bottom+18}" text-anchor="middle">${escapeHtml(shortLabel(label,7))}</text>`;
  });
  svg += `</svg>`;
  $("chartCanvas").innerHTML = chartWithLegend(svg, series, payload);
}

function renderGroupedBarChart(payload) {
  const labels=payload.labels||[],series=payload.series||[];
  if(!labels.length||!series.length) return void ($("chartCanvas").innerHTML='<div class="chart-placeholder">分组柱状图需要横轴、系列和数值字段。</div>');
  const all=series.flatMap(s=>s.values||[]).filter(value=>value!==null&&value!=="").map(Number).filter(Number.isFinite);
  if(!all.length) return void ($("chartCanvas").innerHTML='<div class="chart-placeholder">分组柱状图没有有效数值。</div>');
  const options=chartVisualOptions(payload);
  let min=optionalChartNumber(payload.y_min)??Math.min(0,...all),max=optionalChartNumber(payload.y_max)??Math.max(0,...all);
  if(max===min){min-=1;max+=1;}
  const width=760,height=options.height,left=82,right=18,top=24,bottom=82,plotW=width-left-right,plotH=height-top-bottom,step=plotW/labels.length,groupW=step*(1-options.gap),barW=Math.max(2,groupW/series.length),y=v=>top+(max-v)/(max-min)*plotH,zero=y(Math.max(min,Math.min(max,0))),colors=series.map((item,index)=>chartSeriesColor(item,index,payload));
  let svg=`<svg class="chart-svg${payload.style_3d ? " is-3d" : ""}" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(payload.title||"分组柱状图")}" style="font-size:${options.fontSize}px;color:${options.text}"><rect width="${width}" height="${height}" fill="${options.background}"/>`;
  for(let tick=0;tick<=5;tick++){const yy=top+plotH*tick/5,value=max-(max-min)*tick/5;if(options.showGrid)svg+=`<line x1="${left}" y1="${yy}" x2="${width-right}" y2="${yy}" stroke="#dfe6e3" stroke-width="1"/>`;svg+=`<text x="${left-8}" y="${yy+4}" text-anchor="end" fill="${options.text}">${escapeHtml(formatChartMetric(value,1,payload))}</text>`;}
  labels.forEach((label,i)=>{series.forEach((item,si)=>{const raw=item.values?.[i];if(raw===null||raw==="")return;const value=Number(raw);if(!Number.isFinite(value))return;const yy=y(value),x=left+step*i+(step-groupW)/2+si*barW;svg+=`<rect x="${x}" y="${Math.min(yy,zero)}" width="${Math.max(1,barW-1)}" height="${Math.max(1,Math.abs(zero-yy))}" rx="2" fill="${colors[si]}" opacity="${options.opacity}"><title>${escapeHtml(label)} · ${escapeHtml(item.name)}: ${escapeHtml(formatMetric(value,2))}</title></rect>`;});const x=left+step*(i+.5),tickY=height-bottom+18;svg+=`<text x="${x}" y="${tickY}" text-anchor="middle" fill="${options.text}" transform="rotate(${options.rotation} ${x} ${tickY})">${escapeHtml(shortLabel(label,7))}</text>`;});
  svg+=`<text x="${left+plotW/2}" y="${height-12}" text-anchor="middle" fill="${options.text}" style="font-weight:700">${escapeHtml(payload.x_axis_label||"类别")}</text><text x="18" y="${top+plotH/2}" text-anchor="middle" fill="${options.text}" style="font-weight:700" transform="rotate(-90 18 ${top+plotH/2})">${escapeHtml(payload.y_axis_label||"数值")}</text></svg>`;
  $("chartCanvas").innerHTML=chartWithLegend(svg,series,payload);
}

function renderRadarChart(payload) {
  const labels=(payload.labels||[]).slice(0,10),series=(payload.series||[]).slice(0,6);
  if(labels.length<3||!series.length) return void ($("chartCanvas").innerHTML='<div class="chart-placeholder">雷达图至少需要 3 个横轴类别和 1 个系列。</div>');
  const width=760,height=340,cx=360,cy=166,r=112,colors=series.map((item,index)=>chartSeriesColor(item,index,payload)),axisMax=labels.map((_,i)=>Math.max(1,...series.map(s=>Math.abs(Number(s.values?.[i])||0))));
  const point=(i,ratio)=>{const a=-Math.PI/2+i*Math.PI*2/labels.length;return [cx+Math.cos(a)*r*ratio,cy+Math.sin(a)*r*ratio];};
  let svg=`<svg class="chart-svg" viewBox="0 0 ${width} ${height}">`;
  [0.25,.5,.75,1].forEach(level=>{svg+=`<polygon points="${labels.map((_,i)=>point(i,level).join(',')).join(' ')}" fill="none" stroke="#dbe3ea"/>`;});
  labels.forEach((label,i)=>{const [x,y]=point(i,1.15),[ax,ay]=point(i,1);svg+=`<line x1="${cx}" y1="${cy}" x2="${ax}" y2="${ay}" stroke="#dbe3ea"/><text x="${x}" y="${y}" text-anchor="middle">${escapeHtml(shortLabel(label,8))}</text>`;});
  series.forEach((item,si)=>{const pts=labels.map((_,i)=>point(i,Math.max(0,Number(item.values?.[i])||0)/axisMax[i]));svg+=`<polygon points="${pts.join(' ')}" fill="${colors[si%colors.length]}" fill-opacity=".12" stroke="${colors[si%colors.length]}" stroke-width="2"><title>${escapeHtml(item.name)}</title></polygon>`;});svg+=`</svg>`;
  $("chartCanvas").innerHTML=chartWithLegend(svg,series,payload);
}

function renderFunnelChart(payload) {
  const rows=(payload.labels||[]).map((label,i)=>({label,value:Number(payload.values?.[i])})).filter(r=>Number.isFinite(r.value)&&r.value>=0).sort((a,b)=>b.value-a.value);
  if(!rows.length) return void ($("chartCanvas").innerHTML='<div class="chart-placeholder">漏斗图需要非负的阶段数值。</div>');
  const width=760,height=Math.max(260,rows.length*38+32),cx=330,max=Math.max(1,rows[0].value),colors=chartColors();let svg=`<svg class="chart-svg" viewBox="0 0 ${width} ${height}">`;
  rows.forEach((row,i)=>{const next=rows[i+1]?.value??row.value,w1=470*row.value/max,w2=470*next/max,y=16+i*38;svg+=`<polygon points="${cx-w1/2},${y} ${cx+w1/2},${y} ${cx+w2/2},${y+31} ${cx-w2/2},${y+31}" fill="${colors[i%colors.length]}" opacity=".9"><title>${escapeHtml(row.label)}: ${escapeHtml(formatMetric(row.value,2))}</title></polygon><text x="${cx+250}" y="${y+20}">${escapeHtml(shortLabel(row.label,14))} · ${escapeHtml(formatMetric(row.value,1))}</text>`;});svg+=`</svg>`;$("chartCanvas").innerHTML=svg;
}

function renderWaterfallChart(payload) {
  const rows=(payload.labels||[]).map((label,i)=>({label,value:Number(payload.values?.[i])})).filter(r=>Number.isFinite(r.value));
  if(!rows.length) return void ($("chartCanvas").innerHTML='<div class="chart-placeholder">瀑布图没有有效增减数据。</div>');
  let running=0;const items=rows.map(r=>{const start=running;running+=r.value;return {...r,start,end:running};}),all=[0,...items.flatMap(i=>[i.start,i.end])],min=Math.min(...all),max0=Math.max(...all),max=max0===min?min+1:max0;
  const width=760,height=315,left=62,right=18,top=18,bottom=58,plotW=width-left-right,plotH=height-top-bottom,step=plotW/items.length,y=v=>top+(max-v)/(max-min)*plotH;let svg=`<svg class="chart-svg" viewBox="0 0 ${width} ${height}">`;
  items.forEach((item,i)=>{const x=left+step*i+step*.18,bw=step*.64,yy1=y(item.start),yy2=y(item.end),fill=item.value>=0?chartSeriesColor(null,0,payload):chartSeriesColor(null,1,payload);svg+=`<rect x="${x}" y="${Math.min(yy1,yy2)}" width="${bw}" height="${Math.max(2,Math.abs(yy2-yy1))}" rx="3" fill="${fill}"><title>${escapeHtml(item.label)}: ${item.value>=0?'+':''}${escapeHtml(formatMetric(item.value,2))} · 累计 ${escapeHtml(formatMetric(item.end,2))}</title></rect>`;if(i<items.length-1)svg+=`<line x1="${x+bw}" y1="${yy2}" x2="${left+step*(i+1)+step*.18}" y2="${yy2}" stroke="#8996a5" stroke-dasharray="3 3"/>`;svg+=`<text x="${left+step*(i+.5)}" y="${height-bottom+18}" text-anchor="middle">${escapeHtml(shortLabel(item.label,7))}</text>`;});svg+=`</svg>`;$("chartCanvas").innerHTML=svg;
}

function renderTreemapChart(payload) {
  const rows=(payload.labels||[]).map((label,i)=>({label,value:Number(payload.values?.[i])})).filter(r=>Number.isFinite(r.value)&&r.value>0),total=rows.reduce((s,r)=>s+r.value,0);
  if(!rows.length||total<=0) return void ($("chartCanvas").innerHTML='<div class="chart-placeholder">矩形树图需要大于 0 的汇总值。</div>');
  const width=760,height=315,colors=chartColors();let cursor=0,svg=`<svg class="chart-svg" viewBox="0 0 ${width} ${height}">`;
  rows.forEach((row,i)=>{const w=width*row.value/total,x=cursor;cursor+=w;svg+=`<rect x="${x+1}" y="1" width="${Math.max(1,w-2)}" height="313" rx="4" fill="${colors[i%colors.length]}"><title>${escapeHtml(row.label)}: ${escapeHtml(formatMetric(row.value,2))} (${escapeHtml(formatMetric(row.value/total*100,1))}%)</title></rect>`;if(w>58)svg+=`<text x="${x+w/2}" y="150" text-anchor="middle" style="font-weight:800;fill:#fff">${escapeHtml(shortLabel(row.label,8))}</text><text x="${x+w/2}" y="169" text-anchor="middle" style="fill:#fff">${escapeHtml(formatMetric(row.value/total*100,1))}%</text>`;});svg+=`</svg>`;$("chartCanvas").innerHTML=svg;
}

function renderHeatmapChart(payload) {
  const labels = payload.labels || [], series = payload.series || [];
  if (!labels.length || !series.length) return void ($("chartCanvas").innerHTML = '<div class="chart-placeholder">热力图需要两个分类维度和一个数值字段。</div>');
  const values = series.flatMap(item => item.values || []).map(Number).filter(Number.isFinite), max = Math.max(1, ...values.map(Math.abs));
  const width=760, height=315, left=120, right=18, top=45, bottom=34, cellW=(width-left-right)/labels.length, cellH=(height-top-bottom)/series.length;
  let svg=`<svg class="chart-svg${payload.style_3d ? " is-3d" : ""}" viewBox="0 0 ${width} ${height}">`;
  labels.forEach((label,i)=>{ svg += `<text x="${left+cellW*(i+.5)}" y="${top-10}" text-anchor="middle">${escapeHtml(shortLabel(label,7))}</text>`; });
  series.forEach((item,row)=>{ svg += `<text x="${left-8}" y="${top+cellH*(row+.55)}" text-anchor="end">${escapeHtml(shortLabel(item.name,12))}</text>`; labels.forEach((label,col)=>{ const value=Number(item.values?.[col])||0, alpha=.1+.8*Math.abs(value)/max, fill=chartColorWithAlpha(chartSeriesColor(null,value<0?1:0,payload),alpha), textFill=alpha>.55?'#fff':'#29413b'; svg += `<rect x="${left+cellW*col+1}" y="${top+cellH*row+1}" width="${Math.max(1,cellW-2)}" height="${Math.max(1,cellH-2)}" rx="3" fill="${fill}"><title>${escapeHtml(label)} × ${escapeHtml(item.name)}: ${escapeHtml(formatMetric(value,2))}</title></rect>`; if(cellW>45&&cellH>24) svg += `<text x="${left+cellW*(col+.5)}" y="${top+cellH*(row+.58)}" text-anchor="middle" style="fill:${textFill}">${escapeHtml(formatMetric(value,1))}</text>`; }); });
  svg += `</svg>`; $("chartCanvas").innerHTML=svg;
}

function renderBoxChart(payload) {
  const boxes=payload.boxes||[]; if(!boxes.length) return void ($("chartCanvas").innerHTML='<div class="chart-placeholder">箱线图没有有效数值。</div>');
  const all=boxes.flatMap(b=>[b.min,b.q1,b.median,b.q3,b.max,...(b.outliers||[])]).map(Number).filter(Number.isFinite), min=Math.min(...all), max0=Math.max(...all), max=max0===min?min+1:max0;
  const width=760,height=315,left=62,right=18,top=18,bottom=58,plotW=width-left-right,plotH=height-top-bottom,step=plotW/boxes.length,y=v=>top+(max-v)/(max-min)*plotH;
  let svg=`<svg class="chart-svg${payload.style_3d ? " is-3d" : ""}" viewBox="0 0 ${width} ${height}">`;
  for(let tick=0;tick<=4;tick++){const yy=top+plotH*tick/4,v=max-(max-min)*tick/4;svg+=`<line x1="${left}" y1="${yy}" x2="${width-right}" y2="${yy}" stroke="#e4e9ed"/><text x="${left-7}" y="${yy+3}" text-anchor="end">${escapeHtml(formatMetric(v,1))}</text>`;}
  const boxColor=chartSeriesColor(null,0,payload),accentColor=chartSeriesColor(null,1,payload);
  boxes.forEach((b,i)=>{const x=left+step*(i+.5),bw=Math.max(10,Math.min(44,step*.55));svg+=`<line x1="${x}" y1="${y(b.max)}" x2="${x}" y2="${y(b.min)}" stroke="${boxColor}" stroke-width="2"/><line x1="${x-bw*.3}" y1="${y(b.max)}" x2="${x+bw*.3}" y2="${y(b.max)}" stroke="${boxColor}"/><line x1="${x-bw*.3}" y1="${y(b.min)}" x2="${x+bw*.3}" y2="${y(b.min)}" stroke="${boxColor}"/><rect x="${x-bw/2}" y="${y(b.q3)}" width="${bw}" height="${Math.max(2,y(b.q1)-y(b.q3))}" fill="${chartColorWithAlpha(boxColor,.24)}" stroke="${boxColor}"/><line x1="${x-bw/2}" y1="${y(b.median)}" x2="${x+bw/2}" y2="${y(b.median)}" stroke="${accentColor}" stroke-width="3"><title>中位数 ${escapeHtml(formatMetric(b.median,2))} · 样本 ${b.count}</title></line>`;(b.outliers||[]).forEach(v=>{svg+=`<circle cx="${x}" cy="${y(Number(v))}" r="3" fill="${accentColor}"><title>异常值 ${escapeHtml(formatMetric(v,2))}</title></circle>`;});svg+=`<text x="${x}" y="${height-bottom+18}" text-anchor="middle">${escapeHtml(shortLabel(b.label,8))}</text>`;});
  svg+=`</svg>`;$("chartCanvas").innerHTML=svg;
}

function renderGanttChart(payload) {
  const items=payload.items||[]; if(!items.length) return void ($("chartCanvas").innerHTML='<div class="chart-placeholder">甘特图没有有效任务日期。</div>');
  const starts=items.map(i=>Date.parse(i.start)),ends=items.map(i=>Date.parse(i.end)+86400000),min=Math.min(...starts),max=Math.max(...ends),span=Math.max(86400000,max-min),width=760,left=145,right=22,top=25,rowH=Math.max(24,Math.min(36,250/items.length)),height=Math.max(180,top+items.length*rowH+35),plotW=width-left-right;
  const x=t=>left+(t-min)/span*plotW;let svg=`<svg class="chart-svg${payload.style_3d ? " is-3d" : ""}" viewBox="0 0 ${width} ${height}">`;
  const ganttColor=chartSeriesColor(null,0,payload);
  items.forEach((item,i)=>{const y=top+i*rowH,bx=x(Date.parse(item.start)),bw=Math.max(4,x(Date.parse(item.end)+86400000)-bx),progress=Math.max(0,Math.min(100,Number(item.progress)||0));svg+=`<text x="${left-8}" y="${y+16}" text-anchor="end">${escapeHtml(shortLabel(item.task,16))}</text><rect x="${bx}" y="${y+3}" width="${bw}" height="18" rx="6" fill="${chartColorWithAlpha(ganttColor,.18)}"><title>${escapeHtml(item.start)} 至 ${escapeHtml(item.end)} · ${progress}%</title></rect><rect x="${bx}" y="${y+3}" width="${bw*progress/100}" height="18" rx="6" fill="${ganttColor}"/><text x="${Math.min(width-right-4,bx+bw+5)}" y="${y+16}">${escapeHtml(formatMetric(progress,0))}%</text>`;});
  svg+=`</svg>`;$("chartCanvas").innerHTML=svg;
}

function renderDashboard() {
  const target = $("dashboardGrid");
  if (!state.dashboardCharts.length) {
    target.innerHTML = '<div class="dashboard-empty">先生成图表，再点击“固定到看板”；最多保留 6 张。</div>';
    return;
  }
  target.innerHTML = state.dashboardCharts.map((item, index) => `<article class="dashboard-tile"><div class="dashboard-tile-head"><h4>${escapeHtml(item.title)}</h4><button type="button" data-remove-chart="${index}" aria-label="移除此图">×</button></div><div class="chart-canvas">${item.html}</div><p>${escapeHtml(item.summary || "")}</p></article>`).join("");
  target.querySelectorAll("[data-remove-chart]").forEach(button => button.addEventListener("click", () => {
    state.dashboardCharts.splice(Number(button.dataset.removeChart), 1);
    renderDashboard();
  }));
}

function saveCurrentChart() {
  if (!state.chart) return toast("请先生成一张图表", "error");
  if (state.dashboardCharts.length >= 6) return toast("看板最多保留 6 张图，请先移除一张", "error");
  state.dashboardCharts.push({
    title: state.chart.title || "分析图表",
    summary: state.chart.summary || "",
    html: $("chartCanvas").innerHTML
  });
  renderDashboard();
  toast("已固定到经营看板", "success");
}

function downloadCurrentChart() {
  const svg = $("chartCanvas").querySelector("svg");
  if (!svg || !state.chart) return toast("当前图表无法导出，请先重新生成", "error");
  const clone = svg.cloneNode(true);
  clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  const viewBox = clone.viewBox.baseVal;
  const width = Math.max(520, viewBox?.width || 760), height = Math.max(260, viewBox?.height || 315);
  const source = new XMLSerializer().serializeToString(clone);
  const image = new Image();
  image.onload = () => {
    const canvas = document.createElement("canvas");
    canvas.width = width * 2; canvas.height = height * 2;
    const context = canvas.getContext("2d");
    context.fillStyle = "#ffffff"; context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    URL.revokeObjectURL(image.src);
    canvas.toBlob(blob => {
      if (!blob) return toast("PNG 生成失败", "error");
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `${String(state.chart.title || "分析图表").replace(/[\\/:*?"<>|]/g, "_").slice(0, 60)}.png`;
      link.click();
      setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    }, "image/png");
  };
  image.onerror = () => toast("PNG 导出失败，请使用浏览器打印功能", "error");
  image.src = URL.createObjectURL(new Blob([source], { type: "image/svg+xml;charset=utf-8" }));
}

function renderAll() {
  $("taskId").textContent = state.data?.task_id || "—";
  if (state.data?.task_name && document.activeElement !== $("taskName")) $("taskName").value = state.data.task_name;
  renderTables(); renderProfile(); renderPreview(); renderOperations();
}

async function refresh() {
  state.data = await api(endpoints.state);
  renderAll();
}

async function selectTable(id) {
  if (!id) return;
  try {
    if (id !== state.data?.active_table) resetAnalysisView();
    await post(endpoints.select, { table: id }); await refresh();
  }
  catch (e) { toast(e.message, "error"); }
}

function resetAnalysisView() {
  state.analysis = null; state.chart = null; state.aiChartSpec = null;
  $("analysisEmpty")?.classList.remove("hidden");
  $("analysisResult")?.classList.add("hidden");
  if ($("chartCanvas")) $("chartCanvas").innerHTML = '<div class="chart-placeholder">选择左侧字段后生成柱状图、趋势图、占比图、直方图或散点图</div>';
  if ($("chartTitle")) $("chartTitle").textContent = "选择字段生成图表";
  if ($("chartInsight")) $("chartInsight").textContent = "";
  if ($("saveChartBtn")) $("saveChartBtn").disabled = true;
  if ($("downloadChartBtn")) $("downloadChartBtn").disabled = true;
  document.querySelector(".chart-stage")?.setAttribute("data-theme", "default");
}

function updateDependentColumns() {
  const cols = activeTable()?.columns || [];
  ["dedupeColumns", "groupColumns", "aggColumn", "splitColumn", "maskColumns", "columnKeep", "renameColumn", "sortColumn", "replaceColumn",
    "chartDimension", "chartMeasure", "chartSeries", "chartStart", "chartEnd", "chartProgress", "anomalyColumn", "pivotIndex", "pivotColumns", "pivotValue", "rfmCustomer", "rfmDate", "rfmAmount", "fuzzyClusterColumn"].forEach(id => fillSelect($(id), cols));
  const leftCols = columnsFor($("joinLeft").value);
  const rightCols = columnsFor($("joinRight").value);
  fillSelect($("joinLeftKey"), leftCols);
  fillSelect($("joinRightKey"), rightCols);
  const baseCols = columnsFor($("compareBase").value);
  const targetCols = columnsFor($("compareTarget").value);
  const common = baseCols.filter(c => targetCols.includes(c));
  fillSelect($("compareKey"), common);
  fillSelect($("compareColumns"), common);
  fillSelect($("fuzzySourceKey"), columnsFor($("fuzzySource").value));
  fillSelect($("fuzzyLookupKey"), columnsFor($("fuzzyLookup").value));
  const preferredName = values => values.find(value => /客户|公司|企业|供应商|商品|产品|名称|姓名|name/i.test(String(value))) || values[0];
  if (!$("fuzzySourceKey").value) $("fuzzySourceKey").value = preferredName(columnsFor($("fuzzySource").value)) || "";
  if (!$("fuzzyLookupKey").value) $("fuzzyLookupKey").value = preferredName(columnsFor($("fuzzyLookup").value)) || "";
  chooseAnalysisDefaults(cols);
  updateV3Columns();
}

function chooseAnalysisDefaults(cols) {
  if (!cols.length) return;
  const profile = state.data?.profile?.column_profiles || [];
  const semantic = Object.fromEntries(profile.map(item => [item.name, String(item.dtype || "").toLowerCase()]));
  const find = (patterns, typePattern = null) => cols.find(column => {
    const name = String(column);
    return patterns.some(pattern => pattern.test(name)) || (typePattern && typePattern.test(semantic[column] || ""));
  });
  const numeric = find([/金额|收入|销售额|实付|价格|成本|利润|数量|件数|评分|指标/i], /number|numeric|integer|float|数值|整数|小数/);
  const date = find([/日期|时间|年月|月份|季度|date|time/i], /date|time|日期/);
  const startDate = find([/开始|起始|开工|计划开始|start/i]) || date;
  const endDate = find([/结束|截止|完工|计划结束|end|finish|deadline/i]) || date;
  const progress = find([/进度|完成率|progress|percent/i], /number|numeric|integer|float|数值/);
  const customer = find([/客户|会员|用户|买家|顾客|customer|user/i]);
  const dimension = find([/地区|区域|渠道|类别|分类|产品|商品|部门|门店|省|市/i]) || date || cols[0];
  const series = cols.find(column => column !== dimension && /地区|区域|渠道|类别|分类|产品|商品|部门|状态|省|市/i.test(String(column))) || cols.find(column => column !== dimension) || cols[0];
  const setIfBlank = (id, value) => { if ($(id) && !$(id).value && value && cols.includes(value)) $(id).value = value; };
  setIfBlank("chartDimension", dimension);
  setIfBlank("chartMeasure", numeric || cols[0]);
  setIfBlank("chartSeries", series);
  setIfBlank("chartStart", startDate);
  setIfBlank("chartEnd", endDate);
  setIfBlank("chartProgress", progress);
  setIfBlank("anomalyColumn", numeric);
  setIfBlank("pivotIndex", dimension);
  setIfBlank("pivotValue", numeric || cols[0]);
  setIfBlank("rfmCustomer", customer || cols[0]);
  setIfBlank("rfmDate", date);
  setIfBlank("rfmAmount", numeric);
  setIfBlank("fuzzyClusterColumn", find([/客户|公司|企业|供应商|商品|产品|名称|姓名|name/i]) || cols[0]);
}

async function uploadFiles(files) {
  if (!files.length) return;
  const allowed = [".xlsx", ".csv"];
  for (const file of files) {
    const ext = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();
    if (!allowed.includes(ext)) return toast(`不支持文件“${file.name}”，请使用 .xlsx 或 .csv`, "error");
    if (file.size > 50 * 1024 * 1024) return toast(`文件“${file.name}”超过 50 MB 限制`, "error");
  }
  const form = new FormData();
  [...files].forEach(file => form.append("files", file));
  form.append("task_name", $("taskName").value.trim());
  $("uploadProgress").classList.remove("hidden"); busy(true, "正在读取文件与工作表…");
  try {
    await api(endpoints.upload, { method: "POST", body: form });
    await refresh(); toast(`已导入 ${files.length} 个文件`, "success");
  } catch (e) { toast(e.message, "error"); }
  finally { $("uploadProgress").classList.add("hidden"); busy(false); }
}

function ensureTable() { if (!state.data?.active_table) { toast("请先导入并选择数据表", "error"); return false; } return true; }

async function runAction(path, payload, message) {
  busy(true, message);
  try {
    const result = await post(path, payload);
    await refresh(); toast(result.message || "处理完成", "success");
    if (result.download_url) window.location.assign(result.download_url);
    return result;
  } catch (e) { toast(e.message, "error"); }
  finally { busy(false); }
}

async function runAnalysis() {
  if (!ensureTable()) return;
  busy(true, "正在扫描字段、异常、分布和相关关系…");
  try {
    const result = await post(endpoints.analysis, { table: state.data.active_table });
    renderAnalysis(result); toast("智能分析已生成", "success");
  } catch (e) { toast(e.message, "error"); }
  finally { busy(false); }
}

function updateChartControlVisibility() {
  const type = $("chartType").value;
  $("chartSeriesField").classList.toggle("hidden", !["grouped_bar", "stacked_bar", "radar", "heatmap"].includes(type));
  $("chartGanttFields").classList.toggle("hidden", type !== "gantt");
  const measureField = $("chartMeasure").closest(".field");
  measureField.classList.toggle("hidden", type === "gantt");
  $("chartAggregation").closest(".field").classList.toggle("hidden", type === "gantt");
  $("chartDateGrain").closest(".field").classList.toggle("hidden", ["gantt", "histogram", "scatter", "box"].includes(type));
  document.querySelectorAll("[data-chart-type]").forEach(card => card.classList.toggle("active", card.dataset.chartType === type));
}

async function runChart() {
  if (!ensureTable()) return;
  const payload = {
    table: state.data.active_table, chart_type: $("chartType").value,
    dimension: $("chartDimension").value, measure: $("chartMeasure").value,
    series: $("chartSeries").value, start: $("chartStart").value,
    end: $("chartEnd").value, progress: $("chartProgress").value,
    aggregation: $("chartAggregation").value, top_n: Number($("chartTopN").value),
    date_grain: $("chartDateGrain").value, style_3d: $("chartStyle3d").checked,
    theme: "default", number_format: "auto", sort: "auto", reference_lines: [], highlight: null, show_labels: true, show_legend: true
  };
  if (!payload.dimension && payload.chart_type !== "histogram") return toast("请选择维度或横轴字段", "error");
  if (payload.chart_type !== "gantt" && !payload.measure) return toast("请选择指标或纵轴字段", "error");
  if (["grouped_bar", "stacked_bar", "radar", "heatmap"].includes(payload.chart_type) && !payload.series) return toast("请选择系列字段", "error");
  if (payload.chart_type === "gantt" && (!payload.start || !payload.end)) return toast("甘特图必须选择开始日期和结束日期", "error");
  busy(true, "正在计算并生成可视化…");
  try { renderChart(await post(endpoints.chart, payload)); toast("可视化已生成", "success"); }
  catch (e) { toast(e.message, "error"); }
  finally { busy(false); }
}

function applyAiChartSpecToControls(spec) {
  const chart = spec?.chart;
  if (!chart) return;
  const assignments = {
    chartType: chart.chart_type, chartDimension: chart.dimension, chartMeasure: chart.measure,
    chartSeries: chart.series, chartStart: chart.start, chartEnd: chart.end,
    chartProgress: chart.progress, chartAggregation: chart.aggregation,
    chartTopN: String(chart.top_n), chartDateGrain: chart.date_grain
  };
  Object.entries(assignments).forEach(([id,value]) => { if ($(id) && value != null && [...$(id).options].some(option => option.value === value)) $(id).value=value; });
  $("chartStyle3d").checked=chart.style_3d===true;
  updateChartControlVisibility();
}

function renderAiChartResult(payload) {
  const target=$("aiChartResult"), status=payload?.status||"unsupported";
  target.className=`ai-chart-result ${status==="ready"?"ready":"needs-input"}`;
  const questions=Array.isArray(payload?.clarification_questions)?payload.clarification_questions:[];
  const detail=status==="ready"?(payload.normalized_request||payload.message||"图表已生成"):questions.join("；")||(payload.message||"当前需求暂时无法完成");
  target.innerHTML=`<strong>${status==="ready"?"已生成，可继续修改":status==="clarification"?"需要补充":"超出能力"}</strong><span>${escapeHtml(detail)}</span>`;
}

async function runAiChart() {
  if(!ensureTable()) return;
  const prompt=$("aiChartPrompt").value.trim(), apiKey=$("aiChartApiKey").value.trim();
  if(prompt.length<8) return toast("请更具体地说明想看的图表或修改要求", "error");
  if(!apiKey) return toast("请填写 DeepSeek API Key", "error");
  busy(true,state.aiChartSpec?"AI 正在理解修改要求并重新绘图…":"AI 正在选择图表、字段和呈现方式…");
  try {
    const payload=await post(endpoints.aiChartPlan,{prompt,api_key:apiKey,model:$("aiChartModel").value,table_id:state.data.active_table,current_spec:state.aiChartSpec});
    renderAiChartResult(payload);
    if(payload.status==="ready"&&payload.spec&&payload.chart){state.aiChartSpec=payload.spec;applyAiChartSpecToControls(payload.spec);renderChart(payload.chart);toast("AI 可视化已生成，可继续输入一句话修改", "success");}
    else toast(payload.message||"AI 需要更多信息");
  } catch(error) {
    const safe=redactAiSecrets(String(error.message||"AI 可视化生成失败").split(apiKey).join("[API Key 已隐藏]"));
    $("aiChartResult").className="ai-chart-result needs-input";
    $("aiChartResult").innerHTML=`<strong>生成失败</strong><span>${escapeHtml(safe)}</span>`;
    toast(safe,"error");
  } finally { $("aiChartApiKey").value=""; busy(false); }
}

function setupAiChartUi(){
  document.querySelectorAll("[data-ai-chart-example]").forEach(button=>button.addEventListener("click",()=>{$("aiChartPrompt").value=button.dataset.aiChartExample;$("aiChartPrompt").focus();}));
  $("aiChartRunBtn").addEventListener("click",runAiChart);
  $("aiChartClearBtn").addEventListener("click",()=>{state.aiChartSpec=null;$("aiChartPrompt").value="";$("aiChartResult").className="ai-chart-result neutral";$("aiChartResult").innerHTML="<strong>新图上下文</strong><span>下一条需求会从空白图表开始。</span>";});
  window.addEventListener("beforeunload",()=>{$("aiChartApiKey").value="";});
}

function isPlainObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function responseCandidates(payload, extraKeys = []) {
  if (!isPlainObject(payload)) return [];
  const keys = ["result", "summary", "validation", "reconciliation", "report", "data", ...extraKeys];
  const candidates = [payload];
  keys.forEach(key => { if (isPlainObject(payload[key])) candidates.push(payload[key]); });
  candidates.slice(1).forEach(candidate => {
    if (isPlainObject(candidate.summary)) candidates.push(candidate.summary);
  });
  return candidates;
}

function pickResponseValue(payload, keys, fallback = null) {
  for (const candidate of responseCandidates(payload)) {
    for (const key of keys) {
      if (candidate[key] !== undefined && candidate[key] !== null && candidate[key] !== "") return candidate[key];
    }
  }
  return fallback;
}

function extractCollection(payload, keys) {
  if (Array.isArray(payload)) return payload;
  const candidates = responseCandidates(payload, ["state"]);
  if (isPlainObject(payload?.state)) candidates.push(payload.state);
  for (const candidate of candidates) {
    for (const key of keys) {
      if (Array.isArray(candidate[key])) return candidate[key];
    }
  }
  return [];
}

function displayValue(value, fallback = "—") {
  if (value === null || value === undefined || value === "") return fallback;
  if (typeof value === "boolean") return value ? "是" : "否";
  if (typeof value === "number") return Number.isFinite(value) ? value.toLocaleString("zh-CN", { maximumFractionDigits: 4 }) : fallback;
  if (typeof value === "string") return value;
  try { return JSON.stringify(value); }
  catch (_) { return String(value); }
}

function formatRate(value) {
  if (value === null || value === undefined || value === "") return "—";
  if (typeof value === "string" && value.includes("%")) return value;
  const number = Number(value);
  if (!Number.isFinite(number)) return displayValue(value);
  const percentage = Math.abs(number) <= 1 ? number * 100 : number;
  return `${percentage.toLocaleString("zh-CN", { maximumFractionDigits: 2 })}%`;
}

function adoptResponseState(payload) {
  const incoming = isPlainObject(payload?.state) ? payload.state : null;
  if (!incoming) return false;
  const normalized = isPlainObject(incoming.data) && (incoming.data.tables || incoming.data.task_id) ? incoming.data : incoming;
  state.data = { ...(state.data || {}), ...normalized };
  renderAll();
  return true;
}

async function finishV3Response(payload) {
  if (!adoptResponseState(payload)) await refresh();
}

function responseDetailList(payload, keys = ["details", "rule_results", "step_results", "items", "issues"]) {
  return extractCollection(payload, keys).slice(0, 20);
}

function renderV3Result(targetId, { title, message = "", metrics = [], details = [], error = false }) {
  const target = $(targetId);
  if (!target) return;
  const usableMetrics = metrics.filter(item => item && item.value !== undefined && item.value !== null && item.value !== "");
  const detailRows = details.map((item, index) => {
    const object = isPlainObject(item) ? item : { detail: item };
    const itemTitle = object.name || object.title || object.rule || object.rule_type || object.type || object.operation || `明细 ${index + 1}`;
    const itemStatus = object.status || object.result || object.message || object.detail || object.reason || "";
    const failed = object.failed_count ?? object.failure_count ?? object.failed ?? object.count;
    const suffix = failed !== undefined && failed !== null && failed !== "" ? ` · 失败 ${displayValue(failed)}` : "";
    return `<div class="v3-result-item"><strong>${escapeHtml(displayValue(itemTitle))}</strong><small>${escapeHtml(`${displayValue(itemStatus, "")} ${suffix}`.trim() || "已返回结果")}</small></div>`;
  }).join("");
  target.className = `v3-result${error ? " error" : ""}`;
  target.innerHTML = `
    <div class="v3-result-heading"><strong>${escapeHtml(title)}</strong><small>${escapeHtml(message)}</small></div>
    ${usableMetrics.length ? `<div class="v3-result-metrics">${usableMetrics.map(item => `<div class="v3-result-metric"><span>${escapeHtml(item.label)}</span><strong>${escapeHtml(displayValue(item.value))}</strong></div>`).join("")}</div>` : ""}
    ${detailRows ? `<div class="v3-result-details">${detailRows}</div>` : ""}
    ${!usableMetrics.length && !detailRows ? `<div class="v3-inline-note ${error ? "warning" : "success"}">${escapeHtml(message || (error ? "请求未完成" : "服务已返回最新状态"))}</div>` : ""}`;
}

function aiResponseCandidates(payload) {
  const candidates = [];
  const add = value => { if (isPlainObject(value) && !candidates.includes(value)) candidates.push(value); };
  add(payload);
  ["plan", "result", "data", "execution", "report"].forEach(key => add(payload?.[key]));
  [...candidates].forEach(candidate => {
    ["plan", "result", "execution"].forEach(key => add(candidate?.[key]));
  });
  return candidates;
}

function aiResponseValue(payload, keys, fallback = null) {
  for (const candidate of aiResponseCandidates(payload)) {
    for (const key of keys) {
      if (candidate[key] !== undefined && candidate[key] !== null && candidate[key] !== "") return candidate[key];
    }
  }
  return fallback;
}

function aiResponseCollection(payload, keys) {
  for (const candidate of aiResponseCandidates(payload)) {
    for (const key of keys) if (Array.isArray(candidate[key])) return candidate[key];
  }
  return [];
}

function redactAiSecrets(value) {
  return String(value ?? "").replace(/\bsk-[A-Za-z0-9_\-]{6,}\b/g, "[API Key 已隐藏]");
}

function safeAiParams(value, depth = 0) {
  if (depth > 4) return "…";
  if (Array.isArray(value)) return value.slice(0, 30).map(item => safeAiParams(item, depth + 1));
  if (!isPlainObject(value)) return value;
  return Object.fromEntries(Object.entries(value)
    .filter(([key]) => !/api.?key|password|secret|credential|access.?token/i.test(key))
    .slice(0, 30)
    .map(([key, item]) => [key, safeAiParams(item, depth + 1)]));
}

function normalizeAiPlanStatus(payload) {
  const source = isPlainObject(payload?.plan) ? payload.plan : payload;
  const raw = String(aiResponseValue(source, ["status", "plan_status", "decision", "classification"], "")).trim().toLowerCase().replace(/[\s_]+/g, "-");
  const needsInput = aiResponseValue(source, ["needs_clarification", "requires_clarification", "need_more_info"], false);
  const supported = aiResponseValue(source, ["supported", "is_supported", "executable"], null);
  const questions = aiResponseCollection(source, ["questions", "clarifications", "clarification_questions", "missing_fields", "required_information"]);
  if (needsInput === true || questions.length || ["clarification", "needs-clarification", "clarification-required", "needs-input", "need-info", "incomplete", "需补充", "需要补充"].includes(raw)) return "needs-input";
  if (supported === false || ["unsupported", "rejected", "refused", "blocked", "out-of-scope", "not-supported", "impossible", "拒绝", "不支持", "无法完成"].includes(raw)) return "rejected";
  if (supported === true || ["supported", "ready", "executable", "accepted", "ok", "ready-for-confirmation", "待确认", "可执行", "支持"].includes(raw)) return "supported";
  if (aiResponseValue(payload, ["plan_token", "execution_token"], "")) return "supported";
  return "needs-input";
}

function aiStatusMeta(status) {
  if (status === "supported") return { label: "支持，可执行", className: "supported", title: "计划已生成" };
  if (status === "needs-input") return { label: "需要补充", className: "needs-input", title: "需要补充信息" };
  if (status === "rejected") return { label: "超出能力，已拒绝", className: "rejected", title: "当前无法完成" };
  return { label: "尚未规划", className: "neutral", title: "等待生成计划" };
}

function aiPlanToken(payload) {
  return String(aiResponseValue(payload, ["plan_token", "execution_token", "confirmation_token"], "") || "");
}

function aiPlanSteps(payload) {
  return aiResponseCollection(payload, ["steps", "operations", "actions", "plan_steps"]);
}

function formatAiStep(step, index) {
  if (!isPlainObject(step)) return { title: `步骤 ${index + 1}`, description: redactAiSecrets(displayValue(step, "待执行操作")), params: "" };
  const operation = step.operation || step.op || step.action || step.type || "";
  const capability = aiResponseCollection(state.aiCapabilities, ["operations", "capabilities", "supported_operations", "actions"])
    .find(item => isPlainObject(item) && String(item.id || item.operation || item.name) === String(operation));
  const title = step.title || step.name || capability?.label || operation || `步骤 ${index + 1}`;
  const description = step.description || step.summary || step.reason || step.message || "";
  const rawParams = step.params ?? step.parameters ?? step.arguments ?? step.config;
  const friendlyInputIds = Array.isArray(step.input_ids) ? step.input_ids.map(inputId => {
    if (String(inputId).startsWith("$")) return inputId;
    const table = state.data?.tables?.find(item => String(item.id) === String(inputId));
    return table ? `${table.name} (${table.id})` : inputId;
  }) : step.input_ids;
  const params = {
    ...(step.input_ids !== undefined ? { input_tables: friendlyInputIds } : {}),
    ...(step.output_name !== undefined ? { output_name: step.output_name } : {}),
    ...(isPlainObject(rawParams) ? rawParams : rawParams !== undefined ? { value: rawParams } : {})
  };
  let paramsText = "";
  if (Object.keys(params).length) {
    try { paramsText = JSON.stringify(safeAiParams(params)); }
    catch (_) { paramsText = displayValue(params, ""); }
    if (paramsText.length > 700) paramsText = `${paramsText.slice(0, 697)}…`;
  }
  return { title: redactAiSecrets(title), description: redactAiSecrets(description), params: redactAiSecrets(paramsText) };
}

function aiClarificationText(payload) {
  const source = isPlainObject(payload?.plan) ? payload.plan : payload;
  const questions = aiResponseCollection(source, ["questions", "clarifications", "clarification_questions", "missing_fields", "required_information"])
    .map(item => isPlainObject(item) ? (item.question || item.message || item.field || item.name || displayValue(item)) : item)
    .filter(Boolean);
  const direct = aiResponseValue(source, ["clarification", "required_input", "next_question"], "");
  if (direct) questions.unshift(direct);
  return questions.length ? questions.map((item, index) => `${index + 1}. ${redactAiSecrets(displayValue(item))}`).join("\n") : "";
}

function canExecuteAiPlan() {
  return Boolean(!state.busy && state.aiPlanStatus === "supported" && state.aiPlanToken && $("aiConfirmCheck")?.checked);
}

function updateAiActionState() {
  const confirmable = Boolean(!state.busy && state.aiPlanStatus === "supported" && state.aiPlanToken);
  if ($("aiConfirmCheck")) {
    $("aiConfirmCheck").disabled = !confirmable;
    if (!confirmable) $("aiConfirmCheck").checked = false;
  }
  if ($("aiExecuteBtn")) $("aiExecuteBtn").disabled = !canExecuteAiPlan();
}

function setAiInputsLocked(locked) {
  ["aiApiKey", "aiModel", "aiTables", "aiPrompt"].forEach(id => { if ($(id)) $(id).disabled = locked; });
  document.querySelectorAll("[data-ai-example]").forEach(button => { button.disabled = locked; });
  if ($("aiClearBtn")) $("aiClearBtn").disabled = locked;
}

function resetAiPlan(message = "AI 会先判断需求是否在程序能力范围内，再列出具体步骤、参数和风险。") {
  state.aiPlan = null;
  state.aiPlanToken = "";
  state.aiPlanStatus = "neutral";
  const panel = $("aiPlanPanel");
  if (!panel) return;
  panel.className = "ai-plan-panel empty";
  panel.querySelector(".ai-plan-head h3").textContent = "等待生成计划";
  $("aiPlanStatus").className = "ai-status neutral";
  $("aiPlanStatus").textContent = "尚未规划";
  $("aiPlanSummary").textContent = message;
  $("aiNormalizedRequestText").textContent = "";
  $("aiNormalizedRequest").classList.add("hidden");
  $("aiClarification").classList.add("hidden");
  $("aiClarification").textContent = "";
  $("aiPlanSteps").innerHTML = '<li class="empty">生成计划后在这里逐步核对。</li>';
  $("aiPlanRisks").innerHTML = '<li class="empty">暂无风险信息。</li>';
  $("aiExecutionResult").classList.add("hidden");
  updateAiActionState();
}

function renderAiPlan(payload) {
  let status = normalizeAiPlanStatus(payload);
  const token = aiPlanToken(payload);
  if (status === "supported" && !token) status = "needs-input";
  state.aiPlan = payload;
  state.aiPlanToken = token;
  state.aiPlanStatus = status;
  const meta = aiStatusMeta(status);
  const panel = $("aiPlanPanel");
  panel.className = `ai-plan-panel ${meta.className}`;
  panel.querySelector(".ai-plan-head h3").textContent = meta.title;
  $("aiPlanStatus").className = `ai-status ${meta.className}`;
  $("aiPlanStatus").textContent = meta.label;

  const planSource = isPlainObject(payload?.plan) ? payload.plan : payload;
  const normalizedRequest = payload?.normalized_request ?? planSource?.summary ?? "";
  $("aiNormalizedRequestText").textContent = redactAiSecrets(displayValue(normalizedRequest, ""));
  $("aiNormalizedRequest").classList.toggle("hidden", !normalizedRequest);

  let summary = aiResponseValue(planSource, ["message", "reason", "explanation", "summary"], "");
  if (!summary) {
    if (status === "supported") summary = "需求已转换为程序可执行步骤。请逐项核对后再确认执行。";
    else if (status === "rejected") summary = "该需求包含程序当前不支持或无法安全自动完成的操作，系统不会执行。";
    else summary = token ? "计划信息不完整，请补充后重新生成。" : "服务未返回可执行凭证，请根据提示补充需求后重新生成计划。";
  }
  $("aiPlanSummary").textContent = redactAiSecrets(displayValue(summary));

  const clarification = aiClarificationText(payload);
  $("aiClarification").textContent = clarification;
  $("aiClarification").classList.toggle("hidden", !clarification);

  const steps = aiPlanSteps(payload).slice(0, 50);
  $("aiPlanSteps").innerHTML = steps.length ? steps.map((step, index) => {
    const item = formatAiStep(step, index);
    return `<li><strong>${escapeHtml(item.title)}</strong>${item.description ? `<span>${escapeHtml(item.description)}</span>` : ""}${item.params ? `<code>${escapeHtml(item.params)}</code>` : ""}</li>`;
  }).join("") : '<li class="empty">服务未返回结构化步骤。</li>';

  const expirySeconds = Number(aiResponseValue(payload, ["expires_in_seconds", "ttl_seconds"], 0));
  const risks = [
    ...aiResponseCollection(payload, ["risks", "warnings", "safety_notes", "confirmations"]),
    ...aiResponseCollection(planSource, ["assumptions"]),
    ...(Number.isFinite(expirySeconds) && expirySeconds > 0 ? [`此计划凭证约 ${Math.ceil(expirySeconds / 60)} 分钟内有效，且只能执行一次。`] : [])
  ].filter((item, index, values) => values.findIndex(other => displayValue(other) === displayValue(item)) === index).slice(0, 30);
  $("aiPlanRisks").innerHTML = risks.length ? risks.map(item => {
    const text = isPlainObject(item) ? (item.message || item.description || item.risk || item.title || displayValue(item)) : item;
    return `<li>${escapeHtml(redactAiSecrets(displayValue(text)))}</li>`;
  }).join("") : `<li>${status === "supported" ? "执行将生成新结果表；请确认字段、容差、去重和导出规则符合订单约定。" : "当前计划不可执行，不会修改任务数据。"}</li>`;

  $("aiExecutionResult").classList.add("hidden");
  updateAiActionState();
}

function renderAiCapabilities(payload, error = "") {
  const target = $("aiCapabilities");
  if (!target) return;
  if (error) {
    target.className = "ai-capabilities warning";
    target.innerHTML = `<span class="ai-capabilities-dot"></span><span>能力说明暂不可读取：${escapeHtml(redactAiSecrets(error))}。仍可尝试生成计划。</span>`;
    return;
  }
  let operations = aiResponseCollection(payload, ["operations", "capabilities", "supported_operations", "actions"]);
  if (!operations.length) {
    const object = aiResponseValue(payload, ["capabilities", "supported_operations"], null);
    if (isPlainObject(object)) operations = Object.keys(object).filter(key => object[key] !== false);
  }
  const names = operations.map(item => isPlainObject(item) ? (item.label || item.name || item.operation || item.id) : item).filter(Boolean);
  const available = aiResponseValue(payload, ["available", "enabled", "ready"], true) !== false;
  target.className = `ai-capabilities ${available ? "ready" : "warning"}`;
  const detail = names.length ? `支持 ${names.length} 类操作：${names.slice(0, 7).map(redactAiSecrets).join("、")}${names.length > 7 ? "等" : ""}` : "能力接口已连接，生成计划时会逐项校验可执行范围";
  target.innerHTML = `<span class="ai-capabilities-dot"></span><span>${escapeHtml(available ? detail : "AI 规划服务当前未启用，请检查本地服务设置")}</span>`;
}

async function loadAiCapabilities() {
  try {
    const payload = await api(endpoints.aiCapabilities);
    state.aiCapabilities = payload;
    renderAiCapabilities(payload);
  } catch (error) {
    state.aiCapabilities = null;
    renderAiCapabilities(null, error.message);
  }
}

async function diagnoseAiConnection() {
  const apiKey = $("aiApiKey").value.trim();
  const status = $("aiConnectionStatus");
  if (!apiKey) return toast("请先在密钥框填写新的 DeepSeek API Key", "error");
  status.className = "ai-connection-status testing";
  status.textContent = "正在检测 DNS / HTTPS / 密钥 / 余额 / 模型…";
  busy(true, "正在执行最小化 API 连接测试…");
  setAiInputsLocked(true);
  try {
    const payload = await post(endpoints.aiDiagnose, { api_key: apiKey, model: $("aiModel").value });
    status.className = "ai-connection-status success";
    status.textContent = `${payload.message || "连接成功"} · ${payload.model || $("aiModel").value}`;
    toast("DeepSeek API 连接可用", "success");
  } catch (error) {
    const safeMessage = redactAiSecrets(String(error.message || "连接检测失败").split(apiKey).join("[API Key 已隐藏]"));
    status.className = "ai-connection-status error";
    status.textContent = safeMessage;
    toast(safeMessage, "error");
  } finally {
    busy(false);
    setAiInputsLocked(false);
  }
}

async function generateAiPlan() {
  const prompt = $("aiPrompt").value.trim();
  const apiKey = $("aiApiKey").value.trim();
  const tableIds = selectedValues($("aiTables"));
  if (prompt.length < 12) return toast("请更具体地描述处理需求，至少填写 12 个字符", "error");
  if (!apiKey) return toast("请填写 DeepSeek API Key；密钥只用于本次规划请求", "error");
  if (!tableIds.length) return toast("请先导入数据，并选择允许 AI 使用的数据表", "error");

  resetAiPlan("正在让 DeepSeek 理解需求并校验程序能力，请稍候…");
  $("aiPlanPanel").className = "ai-plan-panel";
  $("aiPlanStatus").className = "ai-status executing";
  $("aiPlanStatus").textContent = "规划中";
  busy(true, "AI 正在生成可审查计划，不会自动执行…");
  setAiInputsLocked(true);
  try {
    const payload = await post(endpoints.aiPlan, {
      prompt, api_key: apiKey, model: $("aiModel").value, table_ids: tableIds
    });
    renderAiPlan(payload);
    if (state.aiPlanStatus === "supported") toast("计划已生成，请核对后勾选确认执行", "success");
    else if (state.aiPlanStatus === "rejected") toast("该需求超出当前程序能力，已安全拒绝", "error");
    else toast("计划需要补充信息，请根据提示修改需求");
  } catch (error) {
    const safeMessage = redactAiSecrets(String(error.message || "规划请求失败").split(apiKey).join("[API Key 已隐藏]"));
    resetAiPlan(`计划生成失败：${safeMessage}`);
    $("aiPlanPanel").className = "ai-plan-panel rejected";
    $("aiPlanStatus").className = "ai-status error";
    $("aiPlanStatus").textContent = "规划失败";
    toast(safeMessage, "error");
  } finally {
    busy(false);
    setAiInputsLocked(false);
    updateAiActionState();
  }
}

function renderAiExecution(payload, error = "") {
  const target = $("aiExecutionResult");
  target.classList.remove("hidden");
  const status = $("aiExecutionStatus");
  if (error) {
    status.className = "ai-status error";
    status.textContent = "执行失败";
    $("aiExecutionSummary").textContent = `执行未完成：${redactAiSecrets(error)}`;
    $("aiExecutionMetrics").innerHTML = "";
    $("aiExecutionDetails").innerHTML = "";
    return;
  }
  const rawStatus = String(aiResponseValue(payload, ["status", "execution_status"], "completed")).toLowerCase();
  const failed = /failed|error|rejected|unsupported|blocked|失败|拒绝|不支持/.test(rawStatus);
  const partial = /partial|warning|review|部分|待核验/.test(rawStatus);
  status.className = `ai-status ${failed ? "error" : partial ? "needs-input" : "success"}`;
  status.textContent = failed ? "执行失败" : partial ? "部分完成 / 待核验" : "执行完成";
  const summary = aiResponseValue(payload, ["message", "summary", "result_message"], "计划已执行，结果表已加入当前任务。");
  $("aiExecutionSummary").textContent = redactAiSecrets(displayValue(summary));

  const outputs = aiResponseCollection(payload, ["output_tables", "created_tables", "tables", "outputs"]);
  const metrics = [
    { label: "完成步骤", value: aiResponseValue(payload, ["steps_completed", "executed_steps", "step_count"], null) },
    { label: "新增结果表", value: outputs.length || aiResponseValue(payload, ["tables_created", "output_count"], null) },
    { label: "处理行数", value: aiResponseValue(payload, ["rows_processed", "processed_rows", "row_count"], null) },
    { label: "待人工核验", value: aiResponseValue(payload, ["pending_reviews", "review_count", "needs_review"], null) }
  ].filter(item => item.value !== null && item.value !== undefined && item.value !== "");
  $("aiExecutionMetrics").innerHTML = metrics.map(item => `<div><span>${escapeHtml(item.label)}</span><strong title="${escapeHtml(displayValue(item.value))}">${escapeHtml(displayValue(item.value))}</strong></div>`).join("");

  const details = aiResponseCollection(payload, ["step_results", "results", "details", "operations", "outputs"]).slice(0, 50);
  $("aiExecutionDetails").innerHTML = details.map((item, index) => {
    const object = isPlainObject(item) ? item : { message: item };
    const title = object.title || object.name || object.operation || object.action || `执行项 ${index + 1}`;
    const message = object.message || object.summary || object.result || object.status || displayValue(item);
    return `<li><strong>${escapeHtml(redactAiSecrets(displayValue(title)))}</strong> · ${escapeHtml(redactAiSecrets(displayValue(message)))}</li>`;
  }).join("");
}

async function executeAiPlan() {
  if (!canExecuteAiPlan()) return toast("请先生成可执行计划，并勾选人工确认", "error");
  const token = state.aiPlanToken;
  $("aiExecutionResult").classList.remove("hidden");
  $("aiExecutionStatus").className = "ai-status executing";
  $("aiExecutionStatus").textContent = "执行中";
  $("aiExecutionSummary").textContent = "正在按已确认计划执行，不会扩大数据表范围…";
  $("aiExecutionMetrics").innerHTML = "";
  $("aiExecutionDetails").innerHTML = "";
  busy(true, "正在执行你已确认的 AI 计划…");
  setAiInputsLocked(true);
  try {
    const payload = await post(endpoints.aiExecute, { plan_token: token, confirmed: true });
    const rawStatus = String(aiResponseValue(payload, ["status", "execution_status"], "completed")).toLowerCase();
    if (/failed|error|rejected|unsupported|blocked|失败|拒绝|不支持/.test(rawStatus)) {
      throw new Error(aiResponseValue(payload, ["message", "summary", "reason"], "服务返回执行失败"));
    }
    renderAiExecution(payload);
    if (!adoptResponseState(payload)) await refresh();
    await loadReviews({ silent: true });
    state.aiPlanToken = "";
    state.aiPlanStatus = "completed";
    $("aiConfirmCheck").checked = false;
    $("aiPlanStatus").className = "ai-status success";
    $("aiPlanStatus").textContent = "已执行";
    toast("AI 计划执行完成，请检查新增结果表和核验中心", "success");
  } catch (error) {
    const safeMessage = redactAiSecrets(error.message || "执行请求失败");
    renderAiExecution(null, safeMessage);
    state.aiPlanToken = "";
    state.aiPlanStatus = "needs-input";
    $("aiConfirmCheck").checked = false;
    $("aiPlanPanel").className = "ai-plan-panel needs-input";
    $("aiPlanStatus").className = "ai-status needs-input";
    $("aiPlanStatus").textContent = "需重新规划";
    $("aiPlanSummary").textContent = "执行凭证按一次性安全机制处理。请检查错误原因后重新生成并确认计划。";
    toast(safeMessage, "error");
  } finally {
    busy(false);
    setAiInputsLocked(false);
    updateAiActionState();
  }
}

function setupAiUi() {
  $("aiApiKey").value = "";
  $("aiToggleKeyBtn").addEventListener("click", () => {
    const show = $("aiApiKey").type === "password";
    $("aiApiKey").type = show ? "text" : "password";
    $("aiToggleKeyBtn").textContent = show ? "隐藏" : "显示";
    $("aiToggleKeyBtn").setAttribute("aria-pressed", String(show));
    $("aiToggleKeyBtn").setAttribute("aria-label", show ? "隐藏 API Key" : "显示 API Key");
  });
  $("aiPrompt").addEventListener("input", () => {
    $("aiPromptCount").textContent = `${$("aiPrompt").value.length} / 8000`;
    if (state.aiPlan || state.aiPlanToken) resetAiPlan("需求内容已变化，请重新生成计划。");
  });
  ["aiModel", "aiTables"].forEach(id => $(id).addEventListener("change", () => {
    if (state.aiPlan || state.aiPlanToken) resetAiPlan("模型或数据范围已变化，请重新生成计划。");
  }));
  document.querySelectorAll("[data-ai-example]").forEach(button => button.addEventListener("click", () => {
    $("aiPrompt").value = button.dataset.aiExample;
    $("aiPrompt").dispatchEvent(new Event("input"));
    $("aiPrompt").focus();
  }));
  $("aiClearBtn").addEventListener("click", () => {
    $("aiPrompt").value = "";
    $("aiPromptCount").textContent = "0 / 8000";
    resetAiPlan();
    $("aiPrompt").focus();
  });
  $("aiPlanBtn").addEventListener("click", generateAiPlan);
  $("aiDiagnoseBtn").addEventListener("click", diagnoseAiConnection);
  $("aiConfirmCheck").addEventListener("change", updateAiActionState);
  $("aiExecuteBtn").addEventListener("click", executeAiPlan);
  window.addEventListener("beforeunload", () => { $("aiApiKey").value = ""; });
  updateAiActionState();
}

function renderEngineeringBrief(payload) {
  const target=$("engineeringResult"),brief=payload?.brief||payload;
  if(!brief||typeof brief!=="object") { target.className="engineering-result empty"; target.textContent="服务未返回有效工程方案。"; return; }
  const status=String(brief.status||"clarification"),list=(title,items,ordered=false)=>Array.isArray(items)&&items.length?`<h5>${escapeHtml(title)}</h5><${ordered?'ol':'ul'}>${items.map(item=>`<li>${escapeHtml(redactAiSecrets(displayValue(item)))}</li>`).join("")}</${ordered?'ol':'ul'}>`:"";
  const artifacts=Array.isArray(brief.artifacts)?brief.artifacts:[];
  target.className="engineering-result";
  target.innerHTML=`<h4>${escapeHtml(redactAiSecrets(brief.normalized_request||"高级工程订单方案"))}</h4><p>${escapeHtml(redactAiSecrets(brief.scope||""))}</p>${list("需要客户补充",brief.clarification_questions)}${list("约定交付物",brief.deliverables)}${list("实施步骤",brief.implementation_steps,true)}${artifacts.map(item=>`<article class="engineering-artifact"><header><span>${escapeHtml(item.name||"代码交付物")}</span><em>${escapeHtml(item.language||"text")}</em></header><pre>${escapeHtml(redactAiSecrets(item.content||""))}</pre><p>${escapeHtml(redactAiSecrets(item.usage_note||"仅供人工审查"))}</p></article>`).join("")}${list("测试与验收",brief.test_checklist)}${list("风险提示",brief.risks)}${list("人工审批点",brief.human_approval_points)}`;
  toast(status==="ready"?"高级工程交付包已生成":"还需要补充客户信息",status==="ready"?"success":"");
}

async function generateEngineeringBrief() {
  const prompt=$("engineeringPrompt").value.trim(),apiKey=$("engineeringApiKey").value.trim();
  if(prompt.length<20) return toast("请更具体地说明客户需求、环境和交付物", "error");
  if(!apiKey) return toast("请填写 DeepSeek API Key", "error");
  const tableIds=selectedValues($("aiTables"));
  if(!tableIds.length&&state.data?.active_table) tableIds.push(state.data.active_table);
  busy(true,"AI 正在生成高级工程方案并执行本地安全校验…");
  $("engineeringResult").className="engineering-result empty";
  $("engineeringResult").textContent="正在生成标准话术、交付物、代码/公式、测试清单和人工审批点…";
  try {
    const payload=await post(endpoints.aiEngineering,{category:$("engineeringCategory").value,prompt,api_key:apiKey,model:$("engineeringModel").value,table_ids:tableIds});
    renderEngineeringBrief(payload);
  } catch(error) {
    const safe=redactAiSecrets(String(error.message||"生成失败").split(apiKey).join("[API Key 已隐藏]"));
    $("engineeringResult").className="engineering-result";
    $("engineeringResult").innerHTML=`<h4>方案生成失败</h4><p>${escapeHtml(safe)}</p>`;
    toast(safe,"error");
  } finally { busy(false); }
}

function setupEngineeringUi() {
  const examples={
    vba:"Windows 11 + Office 365。客户每月收到 20 个同结构工作簿，需要 VBA 批量合并指定工作表、校验表头、记录跳过文件和错误日志，生成汇总表；不能覆盖源文件，要求提供模块代码、安装说明、测试案例和回滚方法。",
    power_bi:"客户有订单、客户、商品、日期和预算表，需要设计星型模型，编写销售额、毛利率、同比、环比、预算达成率 DAX，给出 Power Query 清洗步骤、关系、日期表、页面结构、RLS 建议和验收清单。",
    database:"客户要从 MySQL 只读提取近 24 个月订单、退款和客户数据到 Excel；需要字段映射、参数化 SELECT/CTE 查询、增量条件、索引建议、脱敏规则、超时与行数上限，禁止写回数据库。",
    business_decision:"客户要审核供应商付款：合同号、发票号、验收状态一致且金额容差 0.01 元可自动通过；重复票、拆分付款、超预算和缺验收记录进入人工核验。需要决策矩阵、例外清单、审批点和交付验收标准。"
  };
  document.querySelectorAll("[data-engineering-type]").forEach(button=>button.addEventListener("click",()=>{document.querySelectorAll("[data-engineering-type]").forEach(item=>item.classList.toggle("active",item===button));$("engineeringCategory").value=button.dataset.engineeringType;if(!$("engineeringPrompt").value.trim())$("engineeringPrompt").value=examples[button.dataset.engineeringType];}));
  $("generateEngineeringBtn").addEventListener("click",generateEngineeringBrief);
  window.addEventListener("beforeunload",()=>{$("engineeringApiKey").value="";});
}

function recipeIdentifier(recipe) {
  return recipe?.id ?? recipe?.recipe_id ?? recipe?.key ?? recipe?.name ?? "";
}

function recipeSteps(recipe) {
  const value = recipe?.steps ?? recipe?.operations ?? [];
  if (Array.isArray(value)) return value;
  if (typeof value === "string") {
    try { const parsed = JSON.parse(value); return Array.isArray(parsed) ? parsed : []; }
    catch (_) { return []; }
  }
  return [];
}

function selectedRecipe() {
  const selected = String($("recipeSelect")?.value || "");
  return state.recipes.find(recipe => String(recipeIdentifier(recipe)) === selected) || null;
}

function renderSelectedRecipe() {
  const recipe = selectedRecipe();
  const target = $("recipeDetail");
  if (!recipe) {
    target.className = `v3-inline-note${state.recipeLoadError ? " warning" : ""}`;
    target.textContent = state.recipeLoadError ? `配方接口暂不可用：${state.recipeLoadError}` : "选择配方后显示步骤数量和说明。";
    return;
  }
  const steps = recipeSteps(recipe);
  target.className = "v3-inline-note success";
  target.textContent = `${recipe.description || recipe.summary || "未填写说明"} · ${steps.length} 个步骤`;
}

function renderRecipes() {
  const items = state.recipes.map(recipe => ({
    value: String(recipeIdentifier(recipe)),
    label: String(recipe.name || recipe.title || recipeIdentifier(recipe) || "未命名配方")
  })).filter(item => item.value);
  fillSelect($("recipeSelect"), items, null, state.recipeLoadError ? "配方接口暂不可用" : "请选择配方");
  $("recipeCountBadge").textContent = `${items.length} 个配方`;
  renderSelectedRecipe();
}

async function loadRecipes({ silent = false } = {}) {
  try {
    const payload = await api(endpoints.recipes);
    adoptResponseState(payload);
    state.recipes = extractCollection(payload, ["recipes", "items"]);
    state.recipeLoadError = "";
    renderRecipes();
    if (!silent) toast(`已刷新 ${state.recipes.length} 个配方`, "success");
  } catch (error) {
    state.recipeLoadError = error.message;
    renderRecipes();
    if (!silent) toast(`无法读取配方：${error.message}`, "error");
  }
}

function loadRecipePreset(presetKey, { overwriteName = true } = {}) {
  const preset = recipePresets[presetKey] || recipePresets.standard_clean;
  if (overwriteName || !$("recipeName").value.trim()) $("recipeName").value = preset.name;
  if (overwriteName || !$("recipeDescription").value.trim()) $("recipeDescription").value = preset.description;
  $("recipeSteps").value = JSON.stringify(preset.steps, null, 2);
}

function parseRecipeEditorSteps() {
  let steps;
  try { steps = JSON.parse($("recipeSteps").value); }
  catch (error) { throw new Error(`JSON 步骤格式错误：${error.message}`); }
  if (!Array.isArray(steps) || !steps.length) throw new Error("JSON 步骤必须是至少包含一项的数组");
  if (steps.some(step => !isPlainObject(step))) throw new Error("每个配方步骤必须是 JSON 对象");
  return steps;
}

function loadSelectedRecipeToEditor() {
  const recipe = selectedRecipe();
  if (!recipe) return toast("请先选择需要载入的配方", "error");
  $("recipeName").value = recipe.name || recipe.title || "";
  $("recipeDescription").value = recipe.description || recipe.summary || "";
  $("recipeSteps").value = JSON.stringify(recipeSteps(recipe), null, 2);
  toast("配方已载入编辑器", "success");
}

async function saveRecipe() {
  const name = $("recipeName").value.trim();
  if (!name) return toast("请填写配方名称", "error");
  let steps;
  try { steps = parseRecipeEditorSteps(); }
  catch (error) { return toast(error.message, "error"); }
  busy(true, "正在保存可复用配方…");
  try {
    const payload = await post(endpoints.recipeSave, { name, description: $("recipeDescription").value.trim(), steps });
    await finishV3Response(payload);
    renderV3Result("recipeResult", {
      title: "配方已保存",
      message: pickResponseValue(payload, ["message"], "可用于后续订单复跑"),
      metrics: [{ label: "配方名称", value: name }, { label: "步骤数量", value: steps.length }],
      details: responseDetailList(payload, ["steps", "step_results"])
    });
    await loadRecipes({ silent: true });
    toast("配方保存成功", "success");
  } catch (error) {
    renderV3Result("recipeResult", { title: "保存失败", message: error.message, error: true });
    toast(error.message, "error");
  } finally { busy(false); }
}

async function runRecipe(dryRun) {
  const recipe = selectedRecipe();
  const tableId = $("recipeTable").value;
  if (!recipe) return toast("请选择需要运行的配方", "error");
  if (!tableId) return toast("请选择输入数据表", "error");
  busy(true, dryRun ? "正在试运行配方并检查影响范围…" : "正在正式运行配方…");
  try {
    const payload = await post(endpoints.recipeRun, {
      recipe_id: recipeIdentifier(recipe), table_id: tableId,
      output_name: $("recipeOutputName").value.trim() || "配方运行结果", dry_run: dryRun
    });
    await finishV3Response(payload);
    renderV3Result("recipeResult", {
      title: dryRun ? "试运行结果" : "正式运行结果",
      message: pickResponseValue(payload, ["message", "summary"], dryRun ? "未写入结果表" : "已更新任务状态"),
      metrics: [
        { label: "运行模式", value: dryRun ? "试运行" : "正式运行" },
        { label: "执行步骤", value: pickResponseValue(payload, ["steps_count", "step_count", "executed_steps"], recipeSteps(recipe).length) },
        { label: "输入行数", value: pickResponseValue(payload, ["before_rows", "input_rows", "rows_before"]) },
        { label: "输出行数", value: pickResponseValue(payload, ["after_rows", "output_rows", "rows_after"]) }
      ],
      details: responseDetailList(payload, ["step_results", "steps", "details"])
    });
    if (!dryRun) await loadReviews({ silent: true });
    toast(dryRun ? "配方试运行完成" : "配方正式运行完成", "success");
  } catch (error) {
    renderV3Result("recipeResult", { title: "配方运行失败", message: error.message, error: true });
    toast(error.message, "error");
  } finally { busy(false); }
}

let qualityRuleSequence = 0;
const qualityRuleLabels = {
  not_null: "非空", unique: "唯一", range: "数值范围", regex: "正则格式", allowed_values: "允许值"
};

function updateQualityRuleIndexes() {
  document.querySelectorAll(".quality-rule-row").forEach((row, index) => {
    const badge = row.querySelector(".quality-rule-index");
    if (badge) badge.textContent = String(index + 1).padStart(2, "0");
  });
}

function renderQualityRuleParams(row, initial = {}) {
  const type = row.querySelector(".quality-rule-type").value;
  const target = row.querySelector(".quality-rule-params");
  if (type === "range") {
    target.innerHTML = '<div class="quality-param-grid"><input class="input rule-min" type="number" step="any" placeholder="最小值（可空）"><input class="input rule-max" type="number" step="any" placeholder="最大值（可空）"></div>';
    target.querySelector(".rule-min").value = initial.min ?? initial.minimum ?? "";
    target.querySelector(".rule-max").value = initial.max ?? initial.maximum ?? "";
  } else if (type === "regex") {
    target.innerHTML = '<input class="input rule-pattern" placeholder="例如：^1\\d{10}$">';
    target.querySelector(".rule-pattern").value = initial.pattern ?? initial.regex ?? "";
  } else if (type === "allowed_values") {
    target.innerHTML = '<input class="input rule-values" placeholder="逗号或换行分隔，例如：已支付,已退款">';
    const values = initial.values ?? initial.allowed_values ?? [];
    target.querySelector(".rule-values").value = Array.isArray(values) ? values.join(", ") : values;
  } else {
    target.innerHTML = `<span class="rule-helper">${type === "unique" ? "检查重复值，空值是否参与由后端规则决定" : "检查空字符串、空白和缺失值"}</span>`;
  }
}

function addQualityRule(initial = {}) {
  const row = document.createElement("div");
  row.className = "quality-rule-row";
  row.dataset.ruleId = String(++qualityRuleSequence);
  row.innerHTML = `
    <span class="quality-rule-index"></span>
    <select class="select quality-rule-type" aria-label="规则类型">
      <option value="not_null">非空</option><option value="unique">唯一</option><option value="range">数值范围</option><option value="regex">正则格式</option><option value="allowed_values">允许值</option>
    </select>
    <select class="select quality-rule-column" aria-label="规则字段"></select>
    <div class="quality-rule-params"></div>
    <button class="rule-remove" type="button" title="移除规则" aria-label="移除规则">×</button>`;
  const type = qualityRuleLabels[initial.type] ? initial.type : "not_null";
  row.querySelector(".quality-rule-type").value = type;
  fillSelect(row.querySelector(".quality-rule-column"), columnsFor($("validationTable")?.value), initial.column || "", "请选择字段");
  renderQualityRuleParams(row, initial);
  row.querySelector(".quality-rule-type").addEventListener("change", () => renderQualityRuleParams(row));
  row.querySelector(".rule-remove").addEventListener("click", () => { row.remove(); updateQualityRuleIndexes(); });
  $("qualityRules").appendChild(row);
  updateQualityRuleIndexes();
}

function updateV3Columns() {
  const validationColumns = columnsFor($("validationTable")?.value);
  document.querySelectorAll(".quality-rule-column").forEach(select => fillSelect(select, validationColumns));

  const leftColumns = columnsFor($("reconcileLeft")?.value);
  const rightColumns = columnsFor($("reconcileRight")?.value);
  ["reconcileLeftKeys", "reconcileLeftAmount", "reconcileLeftDate"].forEach(id => fillSelect($(id), leftColumns));
  ["reconcileRightKeys", "reconcileRightAmount", "reconcileRightDate"].forEach(id => fillSelect($(id), rightColumns));
}

function collectQualityRules() {
  const rows = [...document.querySelectorAll(".quality-rule-row")];
  if (!rows.length) throw new Error("请至少添加一条验收规则");
  return rows.map((row, index) => {
    const type = row.querySelector(".quality-rule-type").value;
    const column = row.querySelector(".quality-rule-column").value;
    if (!column) throw new Error(`第 ${index + 1} 条规则尚未选择字段`);
    const rule = { type, column };
    if (type === "range") {
      const minText = row.querySelector(".rule-min").value.trim();
      const maxText = row.querySelector(".rule-max").value.trim();
      if (!minText && !maxText) throw new Error(`第 ${index + 1} 条范围规则至少填写一个边界`);
      if (minText) rule.min = Number(minText);
      if (maxText) rule.max = Number(maxText);
      if ((rule.min !== undefined && !Number.isFinite(rule.min)) || (rule.max !== undefined && !Number.isFinite(rule.max))) throw new Error(`第 ${index + 1} 条范围规则边界必须是数字`);
      if (rule.min !== undefined && rule.max !== undefined && rule.min > rule.max) throw new Error(`第 ${index + 1} 条范围规则最小值不能大于最大值`);
    } else if (type === "regex") {
      rule.pattern = row.querySelector(".rule-pattern").value;
      if (!rule.pattern) throw new Error(`第 ${index + 1} 条正则规则不能为空`);
    } else if (type === "allowed_values") {
      rule.values = row.querySelector(".rule-values").value.split(/[,，\n]/).map(value => value.trim()).filter(Boolean);
      if (!rule.values.length) throw new Error(`第 ${index + 1} 条允许值规则不能为空`);
    }
    return rule;
  });
}

async function runValidation() {
  const tableId = $("validationTable").value;
  if (!tableId) return toast("请选择待验收数据表", "error");
  let rules;
  try { rules = collectQualityRules(); }
  catch (error) { return toast(error.message, "error"); }
  busy(true, "正在逐条运行质量验收规则…");
  try {
    const payload = await post(endpoints.validate, {
      table_id: tableId, rules,
      output_name: $("validationOutputName").value.trim() || "质量验收明细"
    });
    await finishV3Response(payload);
    const details = responseDetailList(payload, ["rule_results", "rules", "details", "failures"]);
    renderV3Result("qualityResult", {
      title: "质量验收完成",
      message: pickResponseValue(payload, ["message", "summary"], "验收结果已写入任务状态"),
      metrics: [
        { label: "通过率", value: formatRate(pickResponseValue(payload, ["pass_rate", "passing_rate", "success_rate"])) },
        { label: "失败记录", value: pickResponseValue(payload, ["failed_count", "failure_count", "failed_rows", "failures"]) },
        { label: "通过规则", value: pickResponseValue(payload, ["passed_rules", "passed_count", "success_count"]) },
        { label: "规则总数", value: pickResponseValue(payload, ["total_rules", "rule_count", "total"], rules.length) }
      ],
      details
    });
    await loadReviews({ silent: true });
    toast("质量验收完成", "success");
  } catch (error) {
    renderV3Result("qualityResult", { title: "验收失败", message: error.message, error: true });
    toast(error.message, "error");
  } finally { busy(false); }
}

function readNonNegativeNumber(id, label) {
  const text = $(id).value.trim();
  if (text === "") return 0;
  const value = Number(text);
  if (!Number.isFinite(value) || value < 0) throw new Error(`${label}必须是非负数`);
  return value;
}

async function runAdvancedReconcile() {
  const leftId = $("reconcileLeft").value;
  const rightId = $("reconcileRight").value;
  const leftKeys = selectedValues($("reconcileLeftKeys"));
  const rightKeys = selectedValues($("reconcileRightKeys"));
  if (!leftId || !rightId) return toast("请选择左右数据表", "error");
  if (!leftKeys.length || leftKeys.length !== rightKeys.length) return toast("左右匹配键必须至少一项且数量一致", "error");
  const leftAmount = $("reconcileLeftAmount").value;
  const rightAmount = $("reconcileRightAmount").value;
  const leftDate = $("reconcileLeftDate").value;
  const rightDate = $("reconcileRightDate").value;
  if (!leftAmount || !rightAmount) return toast("请选择左右金额字段", "error");
  if (Boolean(leftDate) !== Boolean(rightDate)) return toast("日期字段需左右同时选择", "error");
  let amountTolerance, dateToleranceDays;
  try {
    amountTolerance = readNonNegativeNumber("reconcileAmountTolerance", "金额容差");
    dateToleranceDays = readNonNegativeNumber("reconcileDateTolerance", "日期容差");
  } catch (error) { return toast(error.message, "error"); }

  const config = {
    left_keys: leftKeys, right_keys: rightKeys,
    amount: { left_column: leftAmount, right_column: rightAmount, tolerance: amountTolerance },
    date: leftDate ? { left_column: leftDate, right_column: rightDate, tolerance_days: dateToleranceDays } : null
  };
  busy(true, "正在按多键与容差执行高级对账…");
  try {
    const payload = await post(endpoints.reconcileAdvanced, {
      left_id: leftId, right_id: rightId, config,
      output_name: $("reconcileOutputName").value.trim() || "高级对账结果"
    });
    await finishV3Response(payload);
    renderV3Result("reconcileResult", {
      title: "高级对账完成",
      message: pickResponseValue(payload, ["message", "summary"], "对账结果已生成"),
      metrics: [
        { label: "匹配记录", value: pickResponseValue(payload, ["matched_count", "matches", "matched", "exact_count"]) },
        { label: "左侧独有", value: pickResponseValue(payload, ["left_only_count", "only_left", "left_unmatched_count"]) },
        { label: "右侧独有", value: pickResponseValue(payload, ["right_only_count", "only_right", "right_unmatched_count"]) },
        { label: "容差差异", value: pickResponseValue(payload, ["difference_count", "mismatch_count", "amount_mismatch_count", "tolerance_mismatch_count"]) }
      ],
      details: responseDetailList(payload, ["details", "differences", "items", "samples"])
    });
    await loadReviews({ silent: true });
    toast("高级对账完成", "success");
  } catch (error) {
    renderV3Result("reconcileResult", { title: "对账失败", message: error.message, error: true });
    toast(error.message, "error");
  } finally { busy(false); }
}

function reviewValue(review, keys, fallback = "") {
  for (const key of keys) {
    if (review?.[key] !== undefined && review[key] !== null) return review[key];
  }
  return fallback;
}

function reviewIdentifier(review) {
  return reviewValue(review, ["id", "review_id", "item_id", "key"], "");
}

function normalizeReviewStatus(value) {
  return String(value ?? "pending").trim().toLowerCase();
}

function reviewStatusMeta(value) {
  const normalized = normalizeReviewStatus(value);
  if (["accepted", "accept", "approved", "confirmed", "已接受", "已确认", "通过"].includes(normalized)) return { label: "已接受", className: "accepted", pending: false };
  if (["rejected", "reject", "denied", "已拒绝", "拒绝"].includes(normalized)) return { label: "已拒绝", className: "rejected", pending: false };
  if (["pending", "open", "new", "unreviewed", "待确认", "待审核", "未处理", ""].includes(normalized)) return { label: "待确认", className: "pending", pending: true };
  return { label: String(value ?? "未知"), className: "", pending: false };
}

function formatReviewScore(value) {
  if (value === null || value === undefined || value === "") return "—";
  const number = Number(value);
  if (!Number.isFinite(number)) return displayValue(value);
  return Math.abs(number) <= 1 ? `${(number * 100).toLocaleString("zh-CN", { maximumFractionDigits: 1 })}%` : number.toLocaleString("zh-CN", { maximumFractionDigits: 3 });
}

function reviewFilterValues(key, fallbackKeys = []) {
  return [...new Set(state.reviews.map(item => String(reviewValue(item, [key, ...fallbackKeys], ""))).filter(Boolean))].sort((a, b) => a.localeCompare(b, "zh-CN"));
}

function renderReviewFilterOptions() {
  fillSelect($("reviewTypeFilter"), reviewFilterValues("type", ["review_type", "category"]).map(value => ({ value, label: value })), null, "全部类型");
  fillSelect($("reviewStatusFilter"), reviewFilterValues("status", ["review_status"]).map(value => ({ value, label: reviewStatusMeta(value).label })), null, "全部状态");
}

function filteredReviews() {
  const type = $("reviewTypeFilter")?.value || "";
  const status = $("reviewStatusFilter")?.value || "";
  return state.reviews.filter(review => {
    const itemType = String(reviewValue(review, ["type", "review_type", "category"], ""));
    const itemStatus = String(reviewValue(review, ["status", "review_status"], "pending"));
    return (!type || itemType === type) && (!status || itemStatus === status);
  });
}

function updateReviewSelectionState() {
  const boxes = [...document.querySelectorAll("#reviewList .review-checkbox:not(:disabled)")];
  const checked = boxes.filter(box => box.checked).length;
  $("reviewSelectAll").checked = boxes.length > 0 && checked === boxes.length;
  $("reviewSelectAll").indeterminate = checked > 0 && checked < boxes.length;
}

function renderReviews() {
  renderReviewFilterOptions();
  const reviews = filteredReviews();
  const pending = state.reviews.filter(review => reviewStatusMeta(reviewValue(review, ["status", "review_status"], "pending")).pending).length;
  $("reviewPendingBadge").textContent = String(pending);
  $("reviewSummary").textContent = `${reviews.length} / ${state.reviews.length} 条 · 待确认 ${pending}`;
  $("reviewSelectAll").checked = false;
  $("reviewSelectAll").indeterminate = false;

  if (!reviews.length) {
    const text = state.reviewLoadError ? `核验接口暂不可用：${state.reviewLoadError}` : "当前筛选条件下没有待核验记录";
    $("reviewList").innerHTML = `<tr><td colspan="7" class="blank-table">${escapeHtml(text)}</td></tr>`;
    return;
  }

  $("reviewList").innerHTML = reviews.map(review => {
    const id = reviewIdentifier(review);
    const type = reviewValue(review, ["type", "review_type", "category"], "未分类");
    const source = reviewValue(review, ["source", "source_table", "table_name", "origin"], "");
    const original = reviewValue(review, ["original", "original_value", "source_value", "left_value", "value"], "—");
    const candidate = reviewValue(review, ["candidate", "candidate_value", "suggested_value", "target_value", "right_value", "difference"], "—");
    const score = reviewValue(review, ["score", "similarity", "confidence", "difference_score"], null);
    const rawStatus = reviewValue(review, ["status", "review_status"], "pending");
    const status = reviewStatusMeta(rawStatus);
    const reason = reviewValue(review, ["reason", "message", "detail", "note", "description"], "—");
    return `<tr>
      <td class="review-check"><input class="review-checkbox" type="checkbox" value="${escapeHtml(displayValue(id, ""))}" ${id === "" ? "disabled" : ""} aria-label="选择核验记录"></td>
      <td><span class="review-primary">${escapeHtml(displayValue(type))}</span><span class="review-secondary">${escapeHtml(displayValue(source, "未标注来源"))}</span></td>
      <td>${escapeHtml(displayValue(original))}</td>
      <td>${escapeHtml(displayValue(candidate))}</td>
      <td>${escapeHtml(formatReviewScore(score))}</td>
      <td><span class="status-chip ${status.className}">${escapeHtml(status.label)}</span></td>
      <td>${escapeHtml(displayValue(reason))}</td>
    </tr>`;
  }).join("");
  document.querySelectorAll("#reviewList .review-checkbox").forEach(box => box.addEventListener("change", updateReviewSelectionState));
}

async function loadReviews({ silent = false } = {}) {
  try {
    const payload = await api(endpoints.reviews);
    adoptResponseState(payload);
    state.reviews = extractCollection(payload, ["reviews", "items", "review_items"]);
    state.reviewLoadError = "";
    renderReviews();
    if (!silent) toast(`已刷新 ${state.reviews.length} 条核验记录`, "success");
  } catch (error) {
    state.reviewLoadError = error.message;
    renderReviews();
    if (!silent) toast(`无法读取核验清单：${error.message}`, "error");
  }
}

function selectedReviewIds() {
  return [...document.querySelectorAll("#reviewList .review-checkbox:checked")].map(box => box.value).filter(Boolean);
}

async function submitReviewDecision(decision) {
  const ids = selectedReviewIds();
  if (!ids.length) return toast("请至少勾选一条核验记录", "error");
  busy(true, decision === "accept" ? "正在批量接受核验记录…" : "正在批量拒绝核验记录…");
  try {
    const payload = await post(endpoints.reviewDecision, {
      ids, decision, note: $("reviewDecisionNote").value.trim()
    });
    await finishV3Response(payload);
    await loadReviews({ silent: true });
    $("reviewFeedback").className = "v3-inline-note success";
    $("reviewFeedback").textContent = pickResponseValue(payload, ["message"], `已处理 ${ids.length} 条核验记录`);
    $("reviewDecisionNote").value = "";
    toast(decision === "accept" ? "已批量接受" : "已批量拒绝", "success");
  } catch (error) {
    $("reviewFeedback").className = "v3-inline-note warning";
    $("reviewFeedback").textContent = `提交失败：${error.message}`;
    toast(error.message, "error");
  } finally { busy(false); }
}

function setupActions() {
  $("loadDemoBtn").addEventListener("click", async () => {
    resetAnalysisView();
    await runAction(endpoints.demo, {}, "正在生成本地销售演示数据…");
  });
  $("runAnalysisBtn").addEventListener("click", runAnalysis);
  $("runChartBtn").addEventListener("click", runChart);
  $("chartType").addEventListener("change", updateChartControlVisibility);
  $("saveChartBtn").addEventListener("click", saveCurrentChart);
  $("downloadChartBtn").addEventListener("click", downloadCurrentChart);
  $("clearDashboardBtn").addEventListener("click", () => { state.dashboardCharts = []; renderDashboard(); });
  $("printDashboardBtn").addEventListener("click", () => window.print());
  updateChartControlVisibility();
  renderDashboard();
  $("exportAnalysisBtn").addEventListener("click", () => {
    if (!ensureTable()) return;
    runAction(endpoints.analysisExport, { table: state.data.active_table }, "正在生成分析交付包…");
  });
  $("runAnomalyBtn").addEventListener("click", async () => {
    if (!ensureTable() || !$("anomalyColumn").value) return toast("请选择需要检测的数值字段", "error");
    resetAnalysisView();
    await runAction(endpoints.anomalies, {
      table: state.data.active_table, column: $("anomalyColumn").value,
      method: $("anomalyMethod").value, output_name: $("anomalyOutputName").value
    }, "正在检测异常值并生成明细…");
  });
  $("runPivotBtn").addEventListener("click", async () => {
    if (!ensureTable() || !$("pivotIndex").value || !$("pivotValue").value) return toast("请选择行维度和数值字段", "error");
    resetAnalysisView();
    await runAction(endpoints.pivot, {
      table: state.data.active_table, index: $("pivotIndex").value,
      columns: $("pivotColumns").value, value: $("pivotValue").value,
      aggregation: $("pivotAggregation").value, output_name: "交叉透视结果"
    }, "正在生成交叉透视结果…");
  });
  $("runRfmBtn").addEventListener("click", async () => {
    if (!ensureTable() || !$("rfmCustomer").value || !$("rfmDate").value || !$("rfmAmount").value) return toast("请选择客户、交易日期和交易金额字段", "error");
    resetAnalysisView();
    await runAction(endpoints.rfm, {
      table: state.data.active_table, customer: $("rfmCustomer").value,
      date: $("rfmDate").value, amount: $("rfmAmount").value,
      output_name: "RFM客户分群"
    }, "正在计算客户价值分群…");
  });
  $("runFuzzyClusterBtn").addEventListener("click", async () => {
    if (!ensureTable() || !$("fuzzyClusterColumn").value) return toast("请选择需要扫描的名称字段", "error");
    resetAnalysisView();
    await runAction(endpoints.fuzzyCluster, {
      table: state.data.active_table, column: $("fuzzyClusterColumn").value,
      threshold: Number($("fuzzyClusterThreshold").value), output_name: $("fuzzyClusterOutput").value
    }, "正在扫描相似名称并生成待确认候选…");
  });
  $("runFuzzyLookupBtn").addEventListener("click", async () => {
    const payload = {
      source: $("fuzzySource").value, lookup: $("fuzzyLookup").value,
      source_key: $("fuzzySourceKey").value, lookup_key: $("fuzzyLookupKey").value,
      threshold: Number($("fuzzyLookupThreshold").value), output_name: $("fuzzyLookupOutput").value
    };
    if (!payload.source || !payload.lookup || !payload.source_key || !payload.lookup_key) return toast("请选择来源表、标准表及名称字段", "error");
    resetAnalysisView();
    await runAction(endpoints.fuzzyLookup, payload, "正在计算最佳候选与待确认记录…");
  });
  $("runCleanBtn").addEventListener("click", () => {
    if (!ensureTable()) return;
    runAction(endpoints.clean, {
      table: state.data.active_table, output_name: $("cleanOutputName").value,
      trim_text: $("cleanTrim").checked, drop_empty: $("cleanEmpty").checked,
      drop_duplicates: $("cleanDuplicates").checked, dedupe_columns: selectedValues($("dedupeColumns")),
      dedupe_keep: $("dedupeKeep").value, fill_missing: $("cleanMissing").checked,
      fill_method: $("fillMethod").value, fill_value: $("fillValue").value,
      infer_types: $("cleanTypes").checked, normalize_columns: $("cleanNormalize").checked
    }, "正在执行完整数据清洗…");
  });
  $("runColumnsBtn").addEventListener("click", () => {
    if (!ensureTable()) return;
    const renameColumn = $("renameColumn").value;
    const renameValue = $("renameValue").value.trim();
    runAction(endpoints.columns, {
      table: state.data.active_table, columns: selectedValues($("columnKeep")),
      rename_column: renameColumn, rename_value: renameValue,
      sort_column: $("sortColumn").value, ascending: $("sortAscending").value === "true",
      output_name: $("columnsOutputName").value
    }, "正在整理字段与排序…");
  });
  $("runReplaceBtn").addEventListener("click", () => {
    if (!ensureTable()) return;
    const payload = {
      table: state.data.active_table, column: $("replaceColumn").value,
      find: $("replaceFind").value, replace: $("replaceWith").value,
      mode: $("replaceMode").value, case_sensitive: $("replaceCase").checked,
      output_name: $("replaceOutputName").value
    };
    if (!payload.column || payload.find === "") return toast("请选择字段并填写查找内容", "error");
    runAction(endpoints.replace, payload, "正在查找并替换…");
  });
  $("runConcatBtn").addEventListener("click", () => {
    const tables = selectedValues($("concatTables"));
    if (tables.length < 2) return toast("请至少选择两个需要追加的数据表", "error");
    runAction(endpoints.concat, { tables, strategy: $("concatStrategy").value, add_source: $("concatSource").checked, output_name: $("concatOutputName").value }, "正在追加合并…");
  });
  $("runJoinBtn").addEventListener("click", () => {
    const payload = { left: $("joinLeft").value, right: $("joinRight").value, left_key: $("joinLeftKey").value, right_key: $("joinRightKey").value, how: $("joinHow").value, output_name: $("joinOutputName").value };
    if (!payload.left || !payload.right || !payload.left_key || !payload.right_key) return toast("请选择左右数据表和匹配字段", "error");
    runAction(endpoints.join, payload, "正在检查键值并匹配…");
  });
  $("runCompareBtn").addEventListener("click", () => {
    const payload = { base: $("compareBase").value, target: $("compareTarget").value, key: $("compareKey").value, columns: selectedValues($("compareColumns")), output_name: $("compareOutputName").value };
    if (!payload.base || !payload.target || !payload.key) return toast("请选择两张表和唯一标识字段", "error");
    runAction(endpoints.compare, payload, "正在比对新旧数据…");
  });
  $("runSummaryBtn").addEventListener("click", () => {
    if (!ensureTable()) return;
    const groups = selectedValues($("groupColumns"));
    if (!groups.length || !$("aggColumn").value) return toast("请选择分组字段和统计字段", "error");
    runAction(endpoints.summary, { table: state.data.active_table, group_by: groups, column: $("aggColumn").value, method: $("aggMethod").value, output_name: $("summaryOutputName").value }, "正在生成分组汇总…");
  });
  $("runSplitBtn").addEventListener("click", () => {
    if (!ensureTable() || !$("splitColumn").value) return toast("请选择需要拆分的字段", "error");
    runAction(endpoints.split, { table: state.data.active_table, column: $("splitColumn").value, mode: $("splitMode").value, output_name: $("splitOutputName").value }, "正在拆分并打包…");
  });
  $("runMaskBtn").addEventListener("click", () => {
    if (!ensureTable()) return;
    const columns = selectedValues($("maskColumns"));
    if (!columns.length) return toast("请选择需要脱敏的字段", "error");
    runAction(endpoints.mask, { table: state.data.active_table, columns, mode: $("maskMode").value, output_name: $("maskOutputName").value }, "正在生成脱敏副本…");
  });
  $("runExportBtn").addEventListener("click", () => {
    const tables = selectedValues($("exportTables"));
    if (!tables.length) return toast("请至少选择一个需要导出的数据表", "error");
    runAction(endpoints.export, {
      tables, format: $("exportFormat").value, filename: $("exportFilename").value,
      include_summary: $("includeSummary").checked, safe_csv: $("safeCsv").checked,
      professional: $("professionalExport").checked
    }, $("professionalExport").checked ? "正在生成专业 V3 交付包…" : "正在生成交付文件…");
  });
}

function activateV3Stage(stageName) {
  document.querySelectorAll("[data-v3-stage]").forEach(button => button.classList.toggle("active", button.dataset.v3Stage === stageName));
  document.querySelectorAll("[data-v3-panel]").forEach(panel => panel.classList.toggle("active", panel.dataset.v3Panel === stageName));
}

function setupV3Ui() {
  document.querySelectorAll("[data-v3-stage]").forEach(button => button.addEventListener("click", () => activateV3Stage(button.dataset.v3Stage)));

  loadRecipePreset($("recipePreset").value);
  $("loadRecipePresetBtn").addEventListener("click", () => loadRecipePreset($("recipePreset").value));
  $("recipePreset").addEventListener("change", () => loadRecipePreset($("recipePreset").value));
  $("refreshRecipesBtn").addEventListener("click", () => loadRecipes());
  $("recipeSelect").addEventListener("change", renderSelectedRecipe);
  $("loadRecipeToEditorBtn").addEventListener("click", loadSelectedRecipeToEditor);
  $("saveRecipeBtn").addEventListener("click", saveRecipe);
  $("dryRunRecipeBtn").addEventListener("click", () => runRecipe(true));
  $("runRecipeBtn").addEventListener("click", () => runRecipe(false));

  addQualityRule({ type: "not_null" });
  $("addQualityRuleBtn").addEventListener("click", () => addQualityRule());
  $("validationTable").addEventListener("change", updateV3Columns);
  $("runValidationBtn").addEventListener("click", runValidation);

  ["reconcileLeft", "reconcileRight"].forEach(id => $(id).addEventListener("change", updateV3Columns));
  $("runAdvancedReconcileBtn").addEventListener("click", runAdvancedReconcile);

  $("refreshReviewsBtn").addEventListener("click", () => loadReviews());
  ["reviewTypeFilter", "reviewStatusFilter"].forEach(id => $(id).addEventListener("change", renderReviews));
  $("reviewSelectAll").addEventListener("change", event => {
    document.querySelectorAll("#reviewList .review-checkbox:not(:disabled)").forEach(box => { box.checked = event.target.checked; });
    updateReviewSelectionState();
  });
  $("acceptReviewsBtn").addEventListener("click", () => submitReviewDecision("accept"));
  $("rejectReviewsBtn").addEventListener("click", () => submitReviewDecision("reject"));
}

function activateWorkspaceMode(mode) {
  const resolved = ["ai", "data", "visual", "workflow"].includes(mode) ? mode : "ai";
  document.body.dataset.workspaceMode = resolved;
  document.querySelectorAll("[data-workspace-mode]").forEach(button => button.classList.toggle("active", button.dataset.workspaceMode === resolved));
  const toggle = (id, show) => { const element=$(id); if(element){element.classList.toggle("hidden", !show); if(show){element.classList.remove("workspace-view-enter"); void element.offsetWidth; element.classList.add("workspace-view-enter");}} };
  toggle("aiCommandCenter", resolved === "ai");
  toggle("v3Hub", resolved === "workflow");
  toggle("metricGrid", resolved === "data");
  const tabs=document.querySelector(".tabs"); if(tabs) tabs.classList.toggle("hidden", resolved !== "data");
  toggle("operationLog", ["data", "workflow"].includes(resolved));
  document.querySelectorAll(".tab-panel").forEach(panel => { panel.classList.add("hidden"); panel.classList.remove("active"); });
  if(resolved === "visual") { $("panel-analysis").classList.remove("hidden"); $("panel-analysis").classList.add("active"); }
  else if(resolved === "workflow") { $("panel-export").classList.remove("hidden"); $("panel-export").classList.add("active"); }
  else if(resolved === "data") {
    const activeTab=document.querySelector(".tab.active");
    const safeTab=activeTab && !["analysis","export"].includes(activeTab.dataset.tab) ? activeTab.dataset.tab : "preview";
    document.querySelectorAll(".tab").forEach(tab=>tab.classList.toggle("active",tab.dataset.tab===safeTab));
    $("panel-"+safeTab).classList.remove("hidden"); $("panel-"+safeTab).classList.add("active");
  }
  window.scrollTo({top:0,behavior:"smooth"});
}

let unifiedPlanToken="", unifiedFollowupChart="", unifiedChartSpec=null;

function unifiedSelectedTables(){return [...document.querySelectorAll("#tableChips input:checked")].map(input=>input.value);}
function unifiedSetBusy(on,message="AI 正在规划任务…"){$("runCommandBtn").disabled=on;$("confirmExecuteBtn").disabled=on;$("chartFollowupBtn").disabled=on;if(on)toast(message);}

function unifiedRenderState(){
  const tables=state.data?.tables||[],active=state.data?.active_table;
  $("dataContext").classList.toggle("empty",!tables.length);
  $("dataSummary").textContent=tables.length?`${tables.length} 张表 · ${tables.reduce((sum,item)=>sum+Number(item.rows||0),0).toLocaleString()} 行 · 默认全部授权给 AI` : "尚未上传文件";
  $("tableChips").innerHTML=tables.map(item=>`<label class="table-chip"><input type="checkbox" value="${escapeHtml(item.id)}" checked>${escapeHtml(item.name)}</label>`).join("");
  $("tableList").innerHTML=tables.length?tables.map(item=>`<button type="button" class="table-row ${item.id===active?"active":""}" data-table-id="${escapeHtml(item.id)}"><span><b>${escapeHtml(item.name)}</b><small>${escapeHtml(item.source||"本机结果")}</small></span><em>${Number(item.rows||0).toLocaleString()} 行 · ${(item.columns||[]).length} 列</em></button>`).join(""):'<div class="empty-state">上传文件后，工作表会显示在这里。</div>';
  document.querySelectorAll("[data-table-id]").forEach(button=>button.addEventListener("click",async()=>{await post(endpoints.select,{table:button.dataset.tableId});await unifiedRefresh();}));
  const current=tables.find(item=>item.id===active),preview=state.data?.preview||{};
  $("previewTitle").textContent=current?.name||"数据预览";
  $("previewMeta").textContent=current?`${Number(current.rows||0).toLocaleString()} 行 · ${(current.columns||[]).length} 列 · 显示前 ${Math.min(30,Number(current.rows||0))} 行`:"—";
  const columns=preview.columns||[],rows=(preview.rows||[]).slice(0,30);
  $("previewTable").innerHTML=columns.length?`<thead><tr>${columns.map(column=>`<th>${escapeHtml(column)}</th>`).join("")}</tr></thead><tbody>${rows.map(row=>`<tr>${columns.map(column=>`<td title="${escapeHtml(row[column])}">${escapeHtml(row[column])}</td>`).join("")}</tr>`).join("")}</tbody>`:'<tbody><tr><td class="empty-state">暂无数据</td></tr></tbody>';
  $("exportBtn").disabled=!tables.length;
}

async function unifiedRefresh(){state.data=await api(endpoints.state);unifiedRenderState();}

async function unifiedUpload(files){
  if(!files?.length)return;const form=new FormData();
  for(const file of files){const ext=file.name.slice(file.name.lastIndexOf(".")).toLowerCase();if(![".xlsx",".xlsm",".csv",".pdf",".png",".jpg",".jpeg",".tif",".tiff",".bmp",".db",".sqlite",".sqlite3",".parquet"].includes(ext))return toast(`不支持“${file.name}”的文件类型`,"error");if(file.size>50*1024*1024)return toast(`“${file.name}”超过 50 MB`,"error");form.append("files",file);}
  // The server derives a clean task name from the first uploaded filename.
  // Omitting a fixed Chinese multipart field also avoids legacy clients
  // displaying replacement characters when no per-part charset is supplied.
  unifiedSetBusy(true,"正在读取文件与工作表…");
  try{const result=await api(endpoints.upload,{method:"POST",body:form});unifiedPlanToken="";unifiedChartSpec=null;unifiedFollowupChart="";$("resultWorkspace").classList.add("hidden");$("chartStage").classList.add("hidden");$("commandInput").value="";$("commandInput").dispatchEvent(new Event("input"));await unifiedRefresh();toast(result.message||"已建立独立任务并导入文件","success");}
  catch(error){toast(error.message,"error");}finally{unifiedSetBusy(false);}
}

function unifiedShowResult(mode,title,message){
  $("resultWorkspace").classList.remove("hidden");$("resultMode").textContent=mode;$("resultTitle").textContent=title;$("assistantMessage").textContent=message||"AI 已完成任务理解与本地安全检查。";$("resultWorkspace").scrollIntoView({behavior:"smooth",block:"start"});
}

function unifiedRenderPlan(payload){
  const steps=payload?.plan?.steps||[];
  $("planPanel").classList.toggle("hidden",!steps.length);
  $("planPanel").innerHTML=steps.map((step,index)=>`<div class="plan-step"><span>${index+1}</span><div><b>${escapeHtml(step.operation||"处理步骤")}</b><small>输入：${escapeHtml((step.inputs||[]).join("、")||"当前数据")} · 输出：${escapeHtml(step.output||"结果表")}</small><small>${escapeHtml(JSON.stringify(step.params||{}))}</small></div></div>`).join("");
  unifiedPlanToken=payload.plan_token||"";unifiedFollowupChart=payload.follow_up_chart_request||"";
  $("approvalBar").classList.toggle("hidden",!unifiedPlanToken);
}

function unifiedRenderEngineering(payload){
  const brief=payload?.brief||{},sections=[];
  const list=(title,items,ordered=false)=>{if(Array.isArray(items)&&items.length)sections.push(`<section class="engineering-section"><h4>${escapeHtml(title)}</h4><${ordered?"ol":"ul"}>${items.map(item=>`<li>${escapeHtml(displayValue(item))}</li>`).join("")}</${ordered?"ol":"ul"}></section>`);};
  const automation=payload?.automation;
  if(automation){
    const summary=automation.model||{},validation=automation.validation?.summary||{};
    const missing=Array.isArray(automation.missing)&&automation.missing.length?`<p><small>要启用无人值守发布，请在项目 .env 配置：${escapeHtml(automation.missing.join("、"))}</small></p>`:"";
    const reportLink=automation.report_url?`<a class="secondary-button" href="${escapeHtml(automation.report_url)}" target="_blank" rel="noopener">打开已发布报表</a>`:"";
    sections.push(`<section class="engineering-section automation-result"><h4>${automation.published?"已完成全自动发布":"自动化交付包已完成"}</h4><p>${escapeHtml(automation.message||"")}</p><ul><li>事实表 ${escapeHtml(summary.fact_rows??0)} 行 · 5 张维度表</li><li>${escapeHtml(summary.measures??0)} 个 DAX · ${escapeHtml(summary.pages??0)} 页 · ${escapeHtml(summary.visuals??0)} 个视觉对象</li><li>本地验收 ${escapeHtml(validation.passed??0)}/${escapeHtml(validation.checks??0)} 通过</li></ul>${missing}<div class="chart-actions"><a class="secondary-button" href="${escapeHtml(automation.download_url||"#")}" download>下载 Power BI 自动化交付包</a>${reportLink}</div></section>`);
  }
  list("交付物",brief.deliverables);list("实施步骤",brief.implementation_steps,true);list("测试与验收",brief.test_checklist);list("风险与人工审批点",[...(brief.risks||[]),...(brief.human_approval_points||[])]);
  (brief.artifacts||[]).forEach(item=>sections.push(`<section class="engineering-section"><h4>${escapeHtml(item.name||"工程交付物")}</h4><pre>${escapeHtml(item.content||"")}</pre><small>${escapeHtml(item.usage_note||"请人工审查后使用")}</small></section>`));
  $("engineeringPanel").classList.remove("hidden");$("engineeringPanel").innerHTML=sections.join("")||'<div class="engineering-section">AI 需要更多业务信息后才能生成工程方案。</div>';
}

async function unifiedHandleResponse(payload){
  $("planPanel").classList.add("hidden");$("engineeringPanel").classList.add("hidden");$("approvalBar").classList.add("hidden");
  const labels={chart:"可视化",data:"数据处理",data_then_chart:"处理 + 可视化",engineering:"工程交付",unsupported:"暂不支持"};
  const mode=payload.mode||"unknown";unifiedShowResult(labels[mode]||"AI 结果",payload.normalized_request||payload.route?.normalized_request||"AI 任务结果",payload.message||payload.route?.reason||payload.privacy);
  if(mode==="chart"){
    if(payload.status==="ready"&&payload.chart){unifiedChartSpec=payload.spec;$("chartStage").classList.remove("hidden");renderChart(payload.chart);}else{$("chartStage").classList.add("hidden");const questions=payload.clarification_questions||[];$("assistantMessage").textContent=questions.join("；")||payload.message||"请补充图表需求。";}
  }else if(mode==="engineering")unifiedRenderEngineering(payload);
  else if(mode==="data"||mode==="data_then_chart"){
    unifiedRenderPlan(payload);
    if(payload.auto_execute&&unifiedPlanToken){
      const operations=(payload?.plan?.steps||[]).map(step=>step?.operation);
      const quarterly=operations.includes("quarterly_sales_report");
      const inventory=operations.includes("inventory_management_report");
      const hr=operations.includes("hr_management_report");
      const adaptive=operations.includes("adaptive_analysis_report");
      const enterprise=operations.includes("enterprise_diagnosis_report");
      $("assistantMessage").textContent=enterprise?"已识别为企业集团经营诊断，正在本机勾稽财务、客户、人员、成本与库存，生成十张专业工作表…":adaptive?"专用模块未完全匹配，正在本机识别主表、字段角色、表关系、指标、趋势和异常并生成九张工作表…":hr?"已识别为员工考勤绩效薪资经营报告，正在本机整合、评分、预警并生成十张工作表…":inventory?"已识别为采购销售库存联动报告，正在本机核算库存、判断补货与积压并生成九张工作表…":quarterly?"已识别为多表季度销售报告，正在本机自动清洗、去重、排除无效订单并生成八张工作表…":"已识别为标准销售经营报告，正在本机自动计算并生成五张工作表…";
      await unifiedExecutePlan();
    }
  }
}

async function unifiedRunCommand(options={}){
  const prompt=(options.prompt??$("commandInput").value).trim();if(prompt.length<8)return toast("请更具体地说明要完成的任务","error");
  const tableIds=options.tableIds||unifiedSelectedTables();unifiedSetBusy(true,options.modeHint==="chart"?"AI 正在修改图表…":"AI 正在理解需求并选择工具…");
  try{const payload=await post(endpoints.aiUnified,{prompt,table_ids:tableIds,current_chart_spec:options.currentSpec??null,mode_hint:options.modeHint??null});await unifiedHandleResponse(payload);}
  catch(error){unifiedShowResult("执行失败","AI 命令未完成",redactAiSecrets(error.message));toast(error.message,"error");}
  finally{unifiedSetBusy(false);}
}

async function unifiedExecutePlan(){
  if(!unifiedPlanToken)return;unifiedSetBusy(true,"正在本机执行已确认的处理计划…");
  try{const result=await post(endpoints.aiExecute,{plan_token:unifiedPlanToken,confirmed:true});unifiedPlanToken="";$("approvalBar").classList.add("hidden");await unifiedRefresh();$("assistantMessage").textContent=result.download_url?`${result.message}，专业 Excel 经营报告已生成并开始下载。`:result.message;toast(result.message,"success");if(result.download_url)window.location.assign(result.download_url);
    if(unifiedFollowupChart){const first=result.output_tables?.[0]?.id||state.data.active_table,prompt=unifiedFollowupChart;unifiedFollowupChart="";await unifiedRunCommand({prompt,tableIds:[first],modeHint:"chart",currentSpec:null});}
  }catch(error){toast(error.message,"error");$("assistantMessage").textContent=error.message;}finally{unifiedSetBusy(false);}
}

async function unifiedExport(){const ids=(state.data?.tables||[]).map(item=>item.id);if(!ids.length)return;unifiedSetBusy(true,"正在生成专业 Excel 交付包…");try{const result=await post(endpoints.export,{tables:ids,format:"xlsx",filename:"AI_Excel_处理交付包",include_summary:true,safe_csv:true,professional:true});window.location.assign(result.download_url);toast("交付包已生成","success");}catch(error){toast(error.message,"error");}finally{unifiedSetBusy(false);}}

async function unifiedClearTask(successMessage="当前任务已清空"){
  await post(endpoints.reset,{});unifiedPlanToken="";unifiedChartSpec=null;unifiedFollowupChart="";$('commandInput').value="";$('commandInput').dispatchEvent(new Event("input"));$('resultWorkspace').classList.add("hidden");await unifiedRefresh();toast(successMessage,"success");
}

async function unifiedDiagnoseConnection(){
  const status=$("aiConfigStatus");status.className="status-pill waiting";status.innerHTML="<i></i>正在测试 AI…";
  try{const result=await post(endpoints.aiDiagnose,{});status.className="status-pill ready";status.innerHTML=`<i></i>AI 连接正常 · ${escapeHtml(result.model||"DeepSeek")}`;toast("DeepSeek 网络、密钥、余额和模型均可用","success");}
  catch(error){status.className="status-pill error";status.innerHTML="<i></i>AI 连接失败";toast(redactAiSecrets(error.message),"error");}
}

async function setupUnifiedUi(){
  $("chooseFileBtn").addEventListener("click",()=>$("fileInput").click());$("dropZone").addEventListener("click",event=>{if(event.target.closest("button"))return;$("fileInput").click();});$("fileInput").addEventListener("change",event=>{unifiedUpload(event.target.files);event.target.value="";});
  ["dragenter","dragover"].forEach(name=>$("dropZone").addEventListener(name,event=>{event.preventDefault();$("dropZone").classList.add("dragging");}));["dragleave","drop"].forEach(name=>$("dropZone").addEventListener(name,event=>{event.preventDefault();$("dropZone").classList.remove("dragging");}));$("dropZone").addEventListener("drop",event=>unifiedUpload(event.dataTransfer.files));
  $("commandInput").addEventListener("input",()=>{$("commandCount").textContent=`${$("commandInput").value.length} / 8000`;});$("commandInput").addEventListener("keydown",event=>{if(event.ctrlKey&&event.key==="Enter"){event.preventDefault();unifiedRunCommand();}});$("runCommandBtn").addEventListener("click",()=>unifiedRunCommand());
  document.querySelectorAll("[data-command]").forEach(button=>button.addEventListener("click",()=>{$("commandInput").value=button.dataset.command;$("commandInput").dispatchEvent(new Event("input"));$("commandInput").focus();}));
  $("confirmExecuteBtn").addEventListener("click",unifiedExecutePlan);$("chartFollowupBtn").addEventListener("click",()=>{const prompt=$("chartFollowup").value.trim();if(prompt)unifiedRunCommand({prompt,tableIds:unifiedSelectedTables(),modeHint:"chart",currentSpec:unifiedChartSpec});});$("downloadChartBtn").addEventListener("click",downloadCurrentChart);
  $("demoBtn").addEventListener("click",async()=>{if((state.data?.tables||[]).length&&!window.confirm("加载演示数据会清空当前任务，再创建一份独立的虚构销售演示任务。是否继续？"))return;unifiedSetBusy(true,"正在新建本地演示任务…");try{await post(endpoints.demo,{});unifiedPlanToken="";unifiedChartSpec=null;unifiedFollowupChart="";$("resultWorkspace").classList.add("hidden");await unifiedRefresh();toast("演示任务已创建，仅保留一张虚构销售表","success");}catch(error){toast(error.message,"error");}finally{unifiedSetBusy(false);}});$("exportBtn").addEventListener("click",unifiedExport);
  $("clearTaskBtn").addEventListener("click",async()=>{if(!(state.data?.tables||[]).length)return toast("当前已经是空任务","success");if(!window.confirm("确定清空当前任务的全部表格、处理结果和临时文件吗？此操作无法恢复。"))return;unifiedSetBusy(true,"正在清空当前任务…");try{await unifiedClearTask();}catch(error){toast(error.message,"error");}finally{unifiedSetBusy(false);}});
  $("resetBtn").addEventListener("click",async()=>{if((state.data?.tables||[]).length&&!window.confirm("确定结束当前任务并新建空白任务吗？"))return;await unifiedClearTask("已新建空白任务");});
  $("aiConfigStatus").addEventListener("click",unifiedDiagnoseConnection);$("aiConfigStatus").addEventListener("keydown",event=>{if(event.key==="Enter"||event.key===" "){event.preventDefault();unifiedDiagnoseConnection();}});
  try{const config=await api(endpoints.configStatus);$("aiConfigStatus").className=`status-pill ${config.configured?"ready":"error"}`;$("aiConfigStatus").innerHTML=`<i></i>${escapeHtml(config.configured?`AI 已配置 · ${config.model}`:"AI 待配置 · 安全密钥")}`;$("modelLabel").textContent=config.configured?`${config.model} 自动判断工具与步骤`:"请运行安全配置脚本保存 DeepSeek API Key";}catch(error){$("aiConfigStatus").className="status-pill error";$("aiConfigStatus").innerHTML="<i></i>AI 配置读取失败";}
  await unifiedRefresh();
}

function setupUi() {
  document.querySelectorAll("[data-workspace-mode]").forEach(button => button.addEventListener("click", () => activateWorkspaceMode(button.dataset.workspaceMode)));
  document.querySelectorAll("[data-chart-type]").forEach(card => card.addEventListener("click", () => {
    $("chartType").value=card.dataset.chartType;
    updateChartControlVisibility();
    $("chartType").scrollIntoView({behavior:"smooth",block:"center"});
  }));
  document.querySelectorAll(".tab").forEach(tab => tab.addEventListener("click", () => {
    if (tab.dataset.tab === "analysis") return activateWorkspaceMode("visual");
    if (tab.dataset.tab === "export") return activateWorkspaceMode("workflow");
    document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active", t === tab));
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.toggle("active", p.id === `panel-${tab.dataset.tab}`));
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.toggle("hidden", p.id !== `panel-${tab.dataset.tab}`));
  }));
  $("activeTable").addEventListener("change", e => selectTable(e.target.value));
  ["joinLeft", "joinRight", "compareBase", "compareTarget", "fuzzySource", "fuzzyLookup"].forEach(id => $(id).addEventListener("change", updateDependentColumns));
  const zone = $("uploadZone");
  $("fileInput").addEventListener("change", e => { uploadFiles(e.target.files); e.target.value = ""; });
  ["dragenter", "dragover"].forEach(evt => zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.add("dragging"); }));
  ["dragleave", "drop"].forEach(evt => zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.remove("dragging"); }));
  zone.addEventListener("drop", e => uploadFiles(e.dataTransfer.files));
  $("undoBtn").addEventListener("click", () => runAction(endpoints.undo, {}, "正在撤销上一步…"));
  $("redoBtn").addEventListener("click", () => runAction(endpoints.redo, {}, "正在重做下一步…"));
  $("resetBtn").addEventListener("click", () => {
    $("modalMessage").textContent = "将删除本任务导入的临时文件、处理结果和预览缓存。此操作无法恢复。";
    $("confirmModal").classList.remove("hidden");
  });
  $("modalCancel").addEventListener("click", () => $("confirmModal").classList.add("hidden"));
  $("modalConfirm").addEventListener("click", async () => {
    $("confirmModal").classList.add("hidden");
    resetAnalysisView();
    state.dashboardCharts = [];
    renderDashboard();
    await runAction(endpoints.reset, {}, "正在清空任务数据…");
  });
  $("taskName").addEventListener("change", () => post("/api/task", { name: $("taskName").value }).catch(e => toast(e.message, "error")));
  document.querySelectorAll(".option-card input").forEach(input => input.addEventListener("change", () => input.closest(".option-card").classList.toggle("checked", input.checked)));
  setupAiUi();
  setupAiChartUi();
  setupEngineeringUi();
  setupV3Ui();
  setupActions();
  activateWorkspaceMode("ai");
}

window.addEventListener("DOMContentLoaded", async () => {
  if ($("unifiedApp")) {
    try { await setupUnifiedUi(); }
    catch (e) { toast(`无法连接本地处理服务：${e.message}`, "error"); }
    return;
  }
  setupUi();
  try {
    await refresh();
    await Promise.all([loadRecipes({ silent: true }), loadReviews({ silent: true }), loadAiCapabilities()]);
  }
  catch (e) { toast(`无法连接本地处理服务：${e.message}`, "error"); }
});
