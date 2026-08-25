import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [sourcePath, outputPath, configPath] = process.argv.slice(2);
if (!sourcePath || !outputPath) {
  console.error("Usage: node build_eval_workbook.mjs <input.csv> <output.xlsx> [config.json]");
  process.exit(2);
}

const colors = { navy: "#17365D", gray: "#F3F6F9", red: "#FCE4D6", amber: "#FFF2CC", border: "#B7C9D6", text: "#1F2937" };
const requiredHeaders = ["case_id", "评测对象", "场景分类", "测试输入", "难度", "样本类型", "预期表现", "硬性约束", "重点观察项", "数据来源（AI合成）"];
const defaultConfig = {
  title: "产品输出质量评测汇总",
  version: "eval-workbook-v1.0",
  metrics: [
    { id: "M1", name: "核心指标1（待配置）", kind: "binary", weight: 0.5, definition: "目标卡中的第一个核心结果。", method: "二值：达成=1；未达成=0。", anchors: "1=满足目标卡全部必要条件；0=任一必要条件缺失。", positive_examples: ["正例：满足全部必要条件。", "正例：目标结果在输出中可观察。"], negative_examples: ["反例：遗漏至少一个必要条件。", "反例：只有形式完整但未达成目标。"], disagreement: "将形式完整误判为业务目标达成。", hard_line: "核心目标未达成时进入复核。" },
    { id: "M2", name: "核心指标2（待配置）", kind: "scale5", weight: 0.5, definition: "目标卡中的第二个核心结果强度。", method: "1–5 绝对评分。", anchors: "1=无有效表现；3=部分满足；5=满足全部锚点且无硬伤。", positive_examples: ["正例：关键行为和结果均可观察。", "正例：输出满足全部评分锚点。"], negative_examples: ["反例：表现停留在抽象表述。", "反例：触发硬伤或关键缺口。"], disagreement: "用个人偏好替代业务结果判断。", hard_line: "触发硬性错误时最高2分。" },
  ],
  veto: { name: "硬性约束", definition: "不得违反目标卡、法规、安全或事实约束。", method: "二值：通过/不通过。", positive_examples: ["正例：全部硬性约束满足。", "正例：无安全、事实或权限冲突。"], negative_examples: ["反例：触发禁止项。", "反例：关键事实被改写。"], disagreement: "低影响措辞差异与硬性冲突混淆。", hard_line: "任一硬性约束失败即不通过。" },
};

function colLetter(oneBased) { let n = oneBased; let out = ""; while (n > 0) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26); } return out; }
function styleHeader(sheet, address) { const r = sheet.getRange(address); r.format = { fill: colors.navy, font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: colors.border } }; r.format.rowHeight = 32; }
function styleBody(sheet, address) { sheet.getRange(address).format = { font: { color: colors.text }, verticalAlignment: "top", wrapText: true, borders: { preset: "inside", style: "thin", color: colors.border } }; }
function setWidths(sheet, widths) { for (const [c, w] of Object.entries(widths)) sheet.getRange(`${c}:${c}`).format.columnWidth = w; }
function safeName(value) { return String(value).replace(/[^\p{Letter}\p{Number}_-]+/gu, "_"); }
function normalizeMetric(metric, index) { return { id: String(metric.id || `M${index + 1}`), name: String(metric.name || `指标${index + 1}`), kind: metric.kind || "binary", weight: Number(metric.weight || 0), definition: String(metric.definition || ""), method: String(metric.method || ""), anchors: String(metric.anchors || ""), positive_examples: Array.isArray(metric.positive_examples) ? metric.positive_examples : [], negative_examples: Array.isArray(metric.negative_examples) ? metric.negative_examples : [], disagreement: String(metric.disagreement || ""), hard_line: String(metric.hard_line || ""), normalization: metric.normalization || null }; }

const csvText = await fs.readFile(sourcePath, "utf8");
const parsed = await Workbook.fromCSV(csvText, { sheetName: "Source" });
const sourceValues = parsed.worksheets.getItemAt(0).getUsedRange().values;
if (!sourceValues?.length || sourceValues.length < 2) throw new Error("输入 CSV 没有数据行");
const header = sourceValues[0].map((v) => String(v ?? ""));
const missing = requiredHeaders.filter((h) => !header.includes(h));
if (missing.length) throw new Error(`输入 CSV 缺少列：${missing.join("、")}`);
const rows = sourceValues.slice(1).filter((row) => row.some((v) => String(v ?? "").trim() !== ""));
if (!rows.length) throw new Error("输入 CSV 没有有效样本");
const index = Object.fromEntries(header.map((h, i) => [h, i]));
const pick = (row, name) => row[index[name]] ?? "";

let config = defaultConfig;
if (configPath) config = { ...defaultConfig, ...JSON.parse(await fs.readFile(configPath, "utf8")) };
const metrics = (config.metrics || []).map(normalizeMetric);
if (!metrics.length || metrics.length > 5) throw new Error("metrics 必须有 1–5 个指标");
const ids = new Set();
for (const metric of metrics) {
  if (ids.has(metric.id)) throw new Error(`指标 id 重复：${metric.id}`);
  ids.add(metric.id);
  if (!["binary", "scale5", "number"].includes(metric.kind)) throw new Error(`不支持的指标类型：${metric.kind}`);
  if (metric.positive_examples.length < 2 || metric.negative_examples.length < 2) throw new Error(`${metric.id} 必须有至少 2 个正例和 2 个反例`);
}
const weightSum = metrics.reduce((sum, metric) => sum + metric.weight, 0);
const configProblems = [];
if (Math.abs(weightSum - 1) > 0.001) configProblems.push(`权重合计为 ${weightSum}，应为 1`);
for (const metric of metrics) {
  if (metric.kind === "number" && (!metric.normalization || Number(metric.normalization.max) <= Number(metric.normalization.min))) {
    configProblems.push(`${metric.id} 为 number，必须配置 normalization.min/max 才能计入加权参考分`);
  }
}
const configError = configProblems.join("；");
const veto = { ...defaultConfig.veto, ...(config.veto || {}) };
const lastRow = rows.length + 1;

const workbook = Workbook.create();
const evalSheet = workbook.worksheets.add("评测集");
const workbenchSheet = workbook.worksheets.add("评分工作台");
const standardsSheet = workbook.worksheets.add("评分标准");
const summarySheet = workbook.worksheets.add("评分汇总");

// 1. Evaluation set
const evalRows = rows.map((row) => requiredHeaders.map((h) => pick(row, h)));
evalSheet.getRange(`A1:J${lastRow}`).values = [requiredHeaders, ...evalRows];
styleHeader(evalSheet, "A1:J1"); styleBody(evalSheet, `A2:J${lastRow}`);
evalSheet.freezePanes.freezeRows(1); evalSheet.freezePanes.freezeColumns(1); evalSheet.showGridLines = false;
setWidths(evalSheet, { A: 12, B: 20, C: 20, D: 68, E: 12, F: 22, G: 44, H: 48, I: 48, J: 14 });
evalSheet.tables.add(`A1:J${lastRow}`, true, "EvalCasesTable");

// 2. Dynamic scoring workbench
const identity = ["case_id", "评测对象", "场景分类", "难度", "样本类型", "模型输出"];
const metricHeaders = metrics.map((m) => `${m.id}｜${m.name}`);
const workbenchHeader = [...identity, ...metricHeaders, "硬性否决（通过/不通过）", "否决原因", "评分理由", "是否Bad Case（公式）", "测试输入", "预期表现", "硬性约束", "重点观察项", "数据来源（AI合成）", "评分标准版本"];
const metricStart = identity.length + 1;
const metricCols = metrics.map((_, i) => colLetter(metricStart + i));
const vetoCol = colLetter(metricStart + metrics.length);
const vetoReasonCol = colLetter(metricStart + metrics.length + 1);
const scoreReasonCol = colLetter(metricStart + metrics.length + 2);
const badCol = colLetter(metricStart + metrics.length + 3);
const contextStart = metricStart + metrics.length + 4;
const workbenchRows = rows.map((row) => [
  pick(row, "case_id"), pick(row, "评测对象"), pick(row, "场景分类"), pick(row, "难度"), pick(row, "样本类型"), null,
  ...Array(metrics.length).fill(null), null, null, null, null,
  pick(row, "测试输入"), pick(row, "预期表现"), pick(row, "硬性约束"), pick(row, "重点观察项"), pick(row, "数据来源（AI合成）"), config.version || "eval-workbook-v1.0",
]);
const workbenchLastCol = colLetter(workbenchHeader.length);
workbenchSheet.getRange(`A1:${workbenchLastCol}${lastRow}`).values = [workbenchHeader, ...workbenchRows];
styleHeader(workbenchSheet, `A1:${workbenchLastCol}1`); styleBody(workbenchSheet, `A2:${workbenchLastCol}${lastRow}`);
workbenchSheet.freezePanes.freezeRows(1); workbenchSheet.freezePanes.freezeColumns(6); workbenchSheet.showGridLines = false;
const widths = { A: 12, B: 20, C: 20, D: 12, E: 22, F: 56 };
for (const col of metricCols) widths[col] = 18;
widths[vetoCol] = 20; widths[vetoReasonCol] = 34; widths[scoreReasonCol] = 40; widths[badCol] = 18;
for (let i = 0; i < 6; i++) widths[colLetter(contextStart + i)] = i === 0 ? 68 : i === 1 ? 44 : i < 5 ? 48 : 24;
setWidths(workbenchSheet, widths);
for (const [i, metric] of metrics.entries()) {
  const col = metricCols[i];
  const range = `${col}2:${col}${lastRow}`;
  if (metric.kind === "binary") workbenchSheet.getRange(range).dataValidation = { rule: { type: "list", values: [0, 1] } };
  else if (metric.kind === "scale5") workbenchSheet.getRange(range).dataValidation = { rule: { type: "whole", operator: "between", formula1: 1, formula2: 5 } };
}
workbenchSheet.getRange(`${vetoCol}2:${vetoCol}${lastRow}`).dataValidation = { rule: { type: "list", values: ["通过", "不通过"] } };
const riskyChecks = metrics.flatMap((m, i) => m.kind === "binary" ? [`${metricCols[i]}r=0`] : m.kind === "scale5" ? [`${metricCols[i]}r<=2`] : []);
workbenchSheet.getRange(`${badCol}2:${badCol}${lastRow}`).formulas = rows.map((_, i) => {
  const r = i + 2;
  const filled = [`F${r}<>""`, ...metricCols.map((c) => `${c}${r}<>""`), `${vetoCol}${r}<>""`];
  const bad = riskyChecks.length ? `OR(${riskyChecks.map((x) => x.replace("r", String(r))).join(",")})` : "FALSE";
  return [`=IF(OR(${filled.join(",")}),IF(${vetoCol}${r}="不通过","是",IF(${bad},"需复核","否")),"")`];
});
workbenchSheet.getRange(`${badCol}2:${badCol}${lastRow}`).conditionalFormats.add("containsText", { text: "是", format: { fill: colors.red, font: { bold: true, color: "#9C0006" } } });
workbenchSheet.getRange(`${badCol}2:${badCol}${lastRow}`).conditionalFormats.add("containsText", { text: "需复核", format: { fill: colors.amber, font: { bold: true, color: "#7F6000" } } });
workbenchSheet.getRange(`${vetoCol}2:${vetoCol}${lastRow}`).conditionalFormats.add("containsText", { text: "不通过", format: { fill: colors.red, font: { bold: true, color: "#9C0006" } } });
workbenchSheet.tables.add(`A1:${workbenchLastCol}${lastRow}`, true, "ScoringWorkbenchTable");

// 3. Standards from config
const standardsHeader = ["维度", "定义", "权重/门槛", "打分方式", "档位锚点", "正例", "反例", "易分歧点", "硬性线"];
const standardRows = metrics.map((m) => [m.id + "｜" + m.name, m.definition, `${m.weight * 100}%`, m.method, m.anchors, m.positive_examples.join("；"), m.negative_examples.join("；"), m.disagreement, m.hard_line]);
standardRows.push([veto.name, veto.definition, "独立硬门槛", veto.method, "通过=满足全部硬性条件；不通过=任一硬性条件失败。", (veto.positive_examples || []).join("；"), (veto.negative_examples || []).join("；"), veto.disagreement || "", veto.hard_line || ""]);
standardsSheet.getRange(`A1:I${standardRows.length + 1}`).values = [standardsHeader, ...standardRows];
styleHeader(standardsSheet, "A1:I1"); styleBody(standardsSheet, `A2:I${standardRows.length + 1}`);
standardsSheet.freezePanes.freezeRows(1); standardsSheet.showGridLines = false;
setWidths(standardsSheet, { A: 24, B: 44, C: 18, D: 32, E: 58, F: 44, G: 44, H: 48, I: 48 });
standardsSheet.tables.add(`A1:I${standardRows.length + 1}`, true, "ScoringStandardsTable");

// 4. Formula-driven summary
summarySheet.getRange("A1:F1").merge(); summarySheet.getRange("A1").values = [[config.title || "产品输出质量评测汇总"]];
summarySheet.getRange("A1:F1").format = { fill: colors.navy, font: { bold: true, color: "#FFFFFF", size: 14 }, horizontalAlignment: "center", verticalAlignment: "center" };
summarySheet.getRange("A2:F2").merge(); summarySheet.getRange("A2").values = [["填入‘评分工作台’后，以下结果由公式自动更新；硬性否决独立判断，不被平均分抵消。"]];
summarySheet.getRange("A2:F2").format = { fill: colors.gray, font: { color: colors.text, italic: true }, wrapText: true };
summarySheet.getRange("A4:F4").values = [["指标", "公式口径", "权重", "当前值", "状态", "备注"]]; styleHeader(summarySheet, "A4:F4");
let summaryRows = metrics.map((m) => [`${m.id}｜${m.name}`, m.kind === "binary" ? "通过数 / 已评分数" : "评分平均值", m.weight, null, null, m.definition]);
summaryRows.push([veto.name, "通过数 / 已裁定数", 0, null, null, "独立硬门槛"]);
summaryRows.push(["Bad Case数", "是否Bad Case=是的case数", 0, null, null, "由工作台公式标记"]);
summaryRows.push(["已填模型输出数", "模型输出非空case数", 0, null, null, "用于判断是否已开始评分"]);
summaryRows.push(["加权参考分", "按指标权重归一化后加权", null, null, null, configError || "仅作参考；硬门槛失败时仍不通过"]);
summaryRows.push(["配置状态", "指标权重合计与指标定义检查", null, null, null, configError || "配置通过"]);
summarySheet.getRange(`A5:F${summaryRows.length + 4}`).values = summaryRows;
const scoreRange = `'评分工作台'!`;
const summaryMetricRows = [];
for (let i = 0; i < metrics.length; i++) {
  const row = 5 + i; const col = metricCols[i];
  summarySheet.getRange(`D${row}`).formulas = [[`=IFERROR(AVERAGE(${scoreRange}${col}2:${col}${lastRow}),"")`]];
  summarySheet.getRange(`E${row}`).formulas = [[`=IF(D${row}="","待评分","已更新")`]];
  summarySheet.getRange(`D${row}`).format.numberFormat = metrics[i].kind === "scale5" || metrics[i].kind === "number" ? "0.0" : "0.0%";
  summaryMetricRows.push({ row, metric: metrics[i] });
}
const vetoRow = 5 + metrics.length; const badRow = vetoRow + 1; const outputRow = badRow + 1; const weightedRow = outputRow + 1; const configRow = weightedRow + 1;
summarySheet.getRange(`D${vetoRow}`).formulas = [[`=IFERROR(COUNTIF(${scoreRange}${vetoCol}2:${vetoCol}${lastRow},"通过")/(COUNTIF(${scoreRange}${vetoCol}2:${vetoCol}${lastRow},"通过")+COUNTIF(${scoreRange}${vetoCol}2:${vetoCol}${lastRow},"不通过")),"")`]];
summarySheet.getRange(`E${vetoRow}`).formulas = [[`=IF(D${vetoRow}="","待裁定",IF(D${vetoRow}<1,"存在否决失败","已更新"))`]];
summarySheet.getRange(`D${badRow}`).formulas = [[`=COUNTIF(${scoreRange}${badCol}2:${badCol}${lastRow},"是")`]]; summarySheet.getRange(`E${badRow}`).formulas = [[`=IF(D${badRow}=0,"无Bad Case","需复核")`]];
summarySheet.getRange(`D${outputRow}`).formulas = [[`=COUNTIF(${scoreRange}F2:F${lastRow},"<>")`]]; summarySheet.getRange(`E${outputRow}`).formulas = [[`=IF(D${outputRow}=0,"待填输出","已填输出")`]];
const weightedTerms = summaryMetricRows.flatMap(({ row, metric }) => { if (metric.kind === "scale5") return [`D${row}/5*${metric.weight}`]; if (metric.kind === "binary") return [`D${row}*${metric.weight}`]; if (metric.kind === "number" && metric.normalization && Number(metric.normalization.max) > Number(metric.normalization.min)) return [`(D${row}-${metric.normalization.min})/(${metric.normalization.max}-${metric.normalization.min})*${metric.weight}`]; return []; });
summarySheet.getRange(`D${weightedRow}`).formulas = [[weightedTerms.length ? `=IF(COUNT(D5:D${4 + metrics.length})<${weightedTerms.length},"",${weightedTerms.join("+")})` : `=""`]];
summarySheet.getRange(`E${weightedRow}`).formulas = [[`=IF(D${weightedRow}="","待评分","已更新")`]];
summarySheet.getRange(`D${configRow}`).values = [[configError ? "错误" : "通过"]]; summarySheet.getRange(`E${configRow}`).values = [[configError ? "需修正" : "已检查"]];
styleBody(summarySheet, `A5:F${configRow}`); summarySheet.getRange(`A5:A${configRow}`).format.font = { bold: true, color: colors.text };
summarySheet.getRange(`C5:C${configRow}`).format.numberFormat = "0%"; summarySheet.getRange(`D${vetoRow}`).format.numberFormat = "0.0%"; summarySheet.getRange(`D${badRow}:D${outputRow}`).format.numberFormat = "0"; summarySheet.getRange(`D${weightedRow}`).format.numberFormat = "0.0%";
summarySheet.freezePanes.freezeRows(4); summarySheet.showGridLines = false; setWidths(summarySheet, { A: 28, B: 44, C: 14, D: 16, E: 18, F: 58 });
summarySheet.tables.add(`A4:F${configRow}`, true, "SummaryTable");

const outputDir = outputPath.includes("/") ? outputPath.slice(0, outputPath.lastIndexOf("/")) : ".";
await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook); await output.save(outputPath);
const check = await workbook.inspect({ kind: "table", sheetId: "评分汇总", range: `A4:F${configRow}`, include: "values,formulas", tableMaxRows: configRow, tableMaxCols: 6, maxChars: 14000 }); console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" }); console.log(errors.ndjson);
for (const sheetName of ["评测集", "评分工作台", "评分标准", "评分汇总"]) { const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" }); await fs.writeFile(`${outputDir}/preview-${safeName(sheetName)}.png`, new Uint8Array(await preview.arrayBuffer())); }
console.log(`saved ${outputPath}`);
