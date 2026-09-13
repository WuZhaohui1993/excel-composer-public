import ExcelJS from "exceljs";
import type { FolderAttributeRow, IssueRow, ReviewFlowRow } from "./types.js";
import { generatePreset } from "./generator.js";
import { cleanText, displayArea, displayCodeName, normalizeDeviceName, normalizeJoinKey } from "./normalizers.js";

export interface GenericFileInput {
  fileName: string;
  mimeType?: string;
  buffer: Buffer;
}

export interface GenericSourceSummary {
  id: string;
  fileName: string;
  kind: "excel" | "pdf" | "derived";
  sheetName: string;
  sheetState: "visible" | "hidden" | "veryHidden";
  role: "data" | "instruction" | "dictionary" | "empty" | "derived" | "issue";
  defaultHeaderRow: number;
  defaultDataStartRow: number;
  rowCount: number;
  columnCount: number;
  headers: string[];
  suggestedFillDownColumns: string[];
  previewRows: Record<string, string>[];
}

export interface GenericSourceConfig {
  sourceId: string;
  headerRow?: number;
  dataStartRow?: number;
  fillDownColumns?: string[];
}

export interface GenericJoinRule {
  sourceId: string;
  leftKey: string;
  rightKey: string;
  joinType: "left" | "inner";
  prefix?: string;
  duplicateStrategy?: "first" | "expand" | "merge";
  multiValueSeparator?: string;
}

export interface GenericOutputColumn {
  label: string;
  sourceId?: string;
  column?: string;
  constant?: string;
  template?: string;
}

export interface GenericTemplateExportConfig {
  sourceId: string;
  headerRow?: number;
  dataStartRow?: number;
}

export interface GenericResultSheetConfig {
  sourceId: string;
  sheetName?: string;
  outputColumns?: GenericOutputColumn[];
}

export interface GenericCombineConfig {
  baseSourceId?: string;
  structuredExtraction?: "none" | "auto";
  /** @deprecated Kept for compatibility with earlier configs. Use structuredExtraction instead. */
  businessPreset?: "none" | "preset";
  includeAuxiliarySheets?: boolean;
  sourceConfigs?: GenericSourceConfig[];
  joins?: GenericJoinRule[];
  outputColumns?: GenericOutputColumn[];
  templateExport?: GenericTemplateExportConfig;
  resultSheets?: GenericResultSheetConfig[];
}

interface GenericTable {
  source: GenericSourceSummary;
  rows: Record<string, string>[];
}

interface GenericParseResult {
  tables: GenericTable[];
  issues: IssueRow[];
}

export interface GenericPreviewResult {
  sources: GenericSourceSummary[];
  outputPreview: Record<string, string>[];
  issues: IssueRow[];
  stats: {
    sourceCount: number;
    outputRowCount: number;
    issueCount: number;
  };
}

const PREVIEW_LIMIT = 100;
const EXCEL_PARSE_IGNORE_NODES = ["dataValidations"];
const HEADER_FIELD_PATTERN = /(序号|名称|姓名|部门|电话|客户|订单|行号|商品|项目|编号|编码|代码|主项号|子主项号|设计分工|负责人|备注|类别|分类|类型|属性|sheet|id|装置|单元|功能区|专业|日期|时间|状态|路径|单位|规模)/i;
const GROUP_HEADER_PATTERN = /(功能区|区域|单元|装置|主项|项目|类别|分类|属性|专业|阶段|设计|分工|部门|单位)/i;
const TECHNICAL_HEADER_PATTERN = /^\{.*\}$|文本\(|数据集|关联|#日期/;

function colName(index: number): string {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function normalizeKey(value: unknown): string {
  return cleanText(value).replace(/[；;，、]/g, ",").replace(/\s+/g, "").toLowerCase();
}

function uniqueHeader(raw: unknown, index: number, used: Set<string>): string {
  const base = cleanText(raw) || `列${colName(index)}`;
  let header = base;
  let suffix = 2;
  while (used.has(header)) {
    header = `${base}_${suffix}`;
    suffix += 1;
  }
  used.add(header);
  return header;
}

function issue(type: string, message: string, extra: Partial<IssueRow> = {}): IssueRow {
  return { type, severity: "warning", message, ...extra };
}

function errorIssue(type: string, message: string, extra: Partial<IssueRow> = {}): IssueRow {
  return { type, severity: "error", message, ...extra };
}

function rowValues(row: ExcelJS.Row, maxColumn: number): string[] {
  const values: string[] = [];
  for (let col = 1; col <= maxColumn; col += 1) {
    values.push(cleanText(row.getCell(col).value));
  }
  return values;
}

function distinctValues(values: string[]): string[] {
  return [...new Set(values.filter(Boolean))];
}

function keywordDistinctCount(values: string[], pattern: RegExp): number {
  return distinctValues(values).filter((value) => pattern.test(value)).length;
}

function isDataLikeCell(value: string): boolean {
  const compact = value.replace(/\s+/g, "");
  if (!compact) return false;
  if (/^[+-]?\d+(?:\.\d+)?$/.test(compact)) return true;
  if (/^\d{2,}[A-Z]?$/.test(compact)) return true;
  if (/^\d+(?:[-~～]\d+)?[A-Z]?$/.test(compact)) return true;
  if (/^[A-Z]{1,8}[-_]\d[\dA-Z._/-]*$/i.test(compact)) return true;
  return compact.length > 28 && !HEADER_FIELD_PATTERN.test(compact);
}

function shouldCombinePreviousHeaderRow(worksheet: ExcelJS.Worksheet, headerRow: number, maxColumn: number): boolean {
  if (headerRow <= 1) return false;
  const previous = rowValues(worksheet.getRow(headerRow - 1), maxColumn).filter(Boolean);
  const current = rowValues(worksheet.getRow(headerRow), maxColumn).filter(Boolean);
  if (previous.length < 2 || current.length < 2) return false;

  const previousDistinct = distinctValues(previous);
  const currentDistinct = distinctValues(current);
  const previousDuplicateCount = previous.length - previousDistinct.length;
  const previousGroupCount = keywordDistinctCount(previous, GROUP_HEADER_PATTERN);
  const currentFieldCount = keywordDistinctCount(current, HEADER_FIELD_PATTERN);
  const currentDataLikeRatio = current.filter(isDataLikeCell).length / current.length;

  return (
    previousDistinct.length >= 3 &&
    previousDuplicateCount >= 2 &&
    previousGroupCount >= 2 &&
    currentDistinct.length >= 2 &&
    currentFieldCount >= 2 &&
    currentDataLikeRatio < 0.35
  );
}

function compactHeaderPart(value: string): string {
  return cleanText(value).replace(/\s+/g, "").toLowerCase();
}

function mergeHeaderParts(parent: string, leaf: string): string {
  const parentText = cleanText(parent);
  const leafText = cleanText(leaf);
  if (!parentText) return leafText;
  if (!leafText) return parentText;
  if (compactHeaderPart(parentText) === compactHeaderPart(leafText)) return leafText;
  return `${parentText}.${leafText}`;
}

function rawHeaderValues(worksheet: ExcelJS.Worksheet, headerRow: number, maxColumn: number): string[] {
  const leaf = rowValues(worksheet.getRow(headerRow), maxColumn);
  if (!shouldCombinePreviousHeaderRow(worksheet, headerRow, maxColumn)) return leaf;
  const parent = rowValues(worksheet.getRow(headerRow - 1), maxColumn);
  return leaf.map((value, index) => mergeHeaderParts(parent[index] ?? "", value));
}

function hasMeaningfulCellValue(cell: ExcelJS.Cell): boolean {
  return cleanText(cell.value) !== "";
}

function effectiveMaxColumn(worksheet: ExcelJS.Worksheet): number {
  let maxColumn = 0;
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (hasMeaningfulCellValue(cell)) {
        maxColumn = Math.max(maxColumn, colNumber);
      }
    });
  });
  return maxColumn || worksheet.columnCount || worksheet.actualColumnCount;
}

function cloneExcelValue<T>(value: T): T {
  if (value == null || typeof value !== "object") return value;
  return JSON.parse(JSON.stringify(value));
}

function normalizeColumnNames(values?: string[]): string[] {
  return (values ?? [])
    .flatMap((value) => cleanText(value).split(/[，,；;、]/))
    .map((value) => cleanText(value))
    .filter(Boolean);
}

function suggestedFillDownColumns(headers: string[]): string[] {
  return headers
    .filter((header) => /(功能区|区域|单元|装置|装置名称|主项|类别|分类|分组|设计分工|所属单位|单位)/.test(header))
    .slice(0, 12);
}

function worksheetState(worksheet: ExcelJS.Worksheet): GenericSourceSummary["sheetState"] {
  if (worksheet.state === "hidden" || worksheet.state === "veryHidden") return worksheet.state;
  return "visible";
}

function classifyWorksheet(
  worksheet: ExcelJS.Worksheet,
  rows: Record<string, string>[],
  headers: string[]
): GenericSourceSummary["role"] {
  const name = cleanText(worksheet.name);
  if (worksheetState(worksheet) !== "visible" || /^hidden#/i.test(name)) return "dictionary";
  if (!rows.length || !headers.some(Boolean)) return "empty";
  if (/^(填写说明|使用说明|导入说明|说明|readme)$/i.test(name)) return "instruction";
  return "data";
}

function detectHeaderRow(worksheet: ExcelJS.Worksheet): number {
  const max = Math.min(20, worksheet.rowCount);
  const maxColumn = effectiveMaxColumn(worksheet);
  let bestRow = 1;
  let bestScore = Number.NEGATIVE_INFINITY;
  for (let rowNo = 1; rowNo <= max; rowNo += 1) {
    const nonEmpty = rowValues(worksheet.getRow(rowNo), maxColumn).filter(Boolean);
    if (nonEmpty.length < 2) continue;
    const distinct = distinctValues(nonEmpty);
    const duplicatePenalty = nonEmpty.length - distinct.length;
    const keywordCells = nonEmpty.filter((value) => HEADER_FIELD_PATTERN.test(value)).length;
    const keywordDistinct = keywordDistinctCount(nonEmpty, HEADER_FIELD_PATTERN);
    const dataLikeCount = nonEmpty.filter(isDataLikeCell).length;
    const dataLikeRatio = dataLikeCount / nonEmpty.length;
    const longTextPenalty = nonEmpty.filter((value) => value.length > 30 && !HEADER_FIELD_PATTERN.test(value)).length * 8;
    const technicalPenalty = nonEmpty.some((value) => TECHNICAL_HEADER_PATTERN.test(value)) ? 80 : 0;
    const titlePenalty = distinct.length <= 2 && nonEmpty.length > 4 ? 60 : 0;
    const dataPenalty = dataLikeCount * 8 + (dataLikeRatio > 0.45 ? 45 : 0);
    const groupedHeaderBonus = shouldCombinePreviousHeaderRow(worksheet, rowNo, maxColumn) ? 35 : 0;
    const latePenalty = rowNo > 10 ? (rowNo - 10) * 4 : 0;
    const score =
      distinct.length * 12 +
      nonEmpty.length +
      keywordDistinct * 20 +
      keywordCells * 2 +
      groupedHeaderBonus -
      duplicatePenalty * 4 -
      dataPenalty -
      longTextPenalty -
      technicalPenalty -
      titlePenalty -
      latePenalty;
    if (score > bestScore) {
      bestScore = score;
      bestRow = rowNo;
    }
  }
  return bestRow;
}

async function parseExcelSource(file: GenericFileInput, sourceConfigs: GenericSourceConfig[], fileIndex: number): Promise<GenericTable[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(file.buffer as any, { ignoreNodes: EXCEL_PARSE_IGNORE_NODES });
  const tables: GenericTable[] = [];
  const includeAuxiliarySheets = sourceConfigs.some((config) => config.sourceId === "__include_auxiliary_sheets__");

  for (const worksheet of workbook.worksheets) {
    const provisionalId = `excel:${fileIndex + 1}:${file.fileName}:${worksheet.name}`;
    if (!includeAuxiliarySheets && worksheetState(worksheet) !== "visible") continue;
    const config = sourceConfigs.find((item) => item.sourceId === provisionalId);
    const maxColumn = effectiveMaxColumn(worksheet);
    const headerRow = config?.headerRow ?? detectHeaderRow(worksheet);
    const dataStartRow = config?.dataStartRow ?? headerRow + 1;
    const used = new Set<string>();
    const rawHeaders = rawHeaderValues(worksheet, headerRow, maxColumn);
    const headers = rawHeaders.map((value, index) => uniqueHeader(value, index, used));
    const rows: Record<string, string>[] = [];
    const fillDownColumns = new Set(normalizeColumnNames(config?.fillDownColumns));
    const lastSeen: Record<string, string> = {};

    for (let rowNo = dataStartRow; rowNo <= worksheet.rowCount; rowNo += 1) {
      const raw = rowValues(worksheet.getRow(rowNo), maxColumn);
      if (!raw.some(Boolean)) continue;
      const row: Record<string, string> = {};
      headers.forEach((header, index) => {
        const value = raw[index] ?? "";
        if (fillDownColumns.has(header)) {
          if (value) {
            lastSeen[header] = value;
            row[header] = value;
          } else {
            row[header] = lastSeen[header] ?? "";
          }
          return;
        }
        row[header] = value;
      });
      rows.push(row);
    }

    const role = classifyWorksheet(worksheet, rows, headers);
    if (!includeAuxiliarySheets && role !== "data") continue;

    tables.push({
      source: {
        id: provisionalId,
        fileName: file.fileName,
        kind: "excel",
        sheetName: worksheet.name,
        sheetState: worksheetState(worksheet),
        role,
        defaultHeaderRow: headerRow,
        defaultDataStartRow: dataStartRow,
        rowCount: rows.length,
        columnCount: headers.length,
        headers,
        suggestedFillDownColumns: suggestedFillDownColumns(headers),
        previewRows: rows.slice(0, 20)
      },
      rows
    });
  }

  return tables;
}

async function parsePdfSource(file: GenericFileInput, fileIndex: number): Promise<GenericTable[]> {
  const pdfjs = (await import("pdfjs-dist/legacy/build/pdf.mjs")) as any;
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(file.buffer),
    disableFontFace: true,
    useSystemFonts: true
  }).promise;
  const rows: Record<string, string>[] = [];

  for (let pageNo = 1; pageNo <= doc.numPages; pageNo += 1) {
    const page = await doc.getPage(pageNo);
    const content = await page.getTextContent({ includeMarkedContent: false });
    const lines: Array<{ y: number; items: Array<{ x: number; text: string }> }> = [];
    for (const item of content.items) {
      const text = cleanText(item.str);
      if (!text) continue;
      const y = Number(item.transform[5]);
      const x = Number(item.transform[4]);
      let line = lines.find((candidate) => Math.abs(candidate.y - y) < 2);
      if (!line) {
        line = { y, items: [] };
        lines.push(line);
      }
      line.items.push({ x, text });
    }
    lines
      .sort((a, b) => b.y - a.y)
      .forEach((line, index) => {
        rows.push({
          page: String(pageNo),
          line: String(index + 1),
          text: line.items.sort((a, b) => a.x - b.x).map((item) => item.text).join(" ")
        });
      });
  }

  const source: GenericSourceSummary = {
    id: `pdf:${fileIndex + 1}:${file.fileName}:文本行`,
    fileName: file.fileName,
    kind: "pdf",
    sheetName: "文本行",
    sheetState: "visible",
    role: "data",
    defaultHeaderRow: 1,
    defaultDataStartRow: 1,
    rowCount: rows.length,
    columnCount: 3,
    headers: ["page", "line", "text"],
    suggestedFillDownColumns: [],
    previewRows: rows.slice(0, 20)
  };

  return [{ source, rows }];
}

function isExcelFile(file: GenericFileInput): boolean {
  return /\.xlsx$/i.test(file.fileName);
}

function isPdfFile(file: GenericFileInput): boolean {
  return /\.pdf$/i.test(file.fileName) || file.mimeType === "application/pdf";
}

function scorePresetMatrixFile(fileName: string): number {
  let score = 0;
  if (/文档矩阵|审阅流程矩阵/.test(fileName)) score += 120;
  if (/矩阵/.test(fileName)) score += 80;
  if (/审阅流程/.test(fileName)) score += 30;
  if (/导入|文件夹属性|属性数据|模板/.test(fileName)) score -= 80;
  return score;
}

function scorePresetPdfFile(fileName: string): number {
  let score = 0;
  if (/主项表|设计分工|项目主项/.test(fileName)) score += 120;
  if (/PDM|pdf/i.test(fileName)) score += 30;
  return score;
}

function pickScoredFile(files: Array<{ file: GenericFileInput; index: number }>, scorer: (fileName: string) => number) {
  return [...files].sort((a, b) => scorer(b.file.fileName) - scorer(a.file.fileName))[0];
}

function derivedTable(
  id: string,
  fileName: string,
  sheetName: string,
  headers: string[],
  rows: Record<string, string>[],
  role: GenericSourceSummary["role"] = "derived"
): GenericTable {
  return {
    source: {
      id,
      fileName,
      kind: "derived",
      sheetName,
      sheetState: "visible",
      role,
      defaultHeaderRow: 1,
      defaultDataStartRow: 2,
      rowCount: rows.length,
      columnCount: headers.length,
      headers,
      suggestedFillDownColumns: suggestedFillDownColumns(headers),
      previewRows: rows.slice(0, 20)
    },
    rows
  };
}

function folderRowRecord(row: FolderAttributeRow): Record<string, string> {
  return {
    文件夹路径: row.folderPath,
    名称: row.name,
    更新名称: row.updateName,
    编码: row.code,
    创建时间: row.createDate,
    更新时间: row.updateDate,
    创建者: row.createBy,
    更新者: row.updateBy,
    所属单位: row.belongingUnit,
    层级结构: row.levelName,
    装置: row.device,
    单元: row.unit,
    设计单位: row.designCompany,
    出图配置: row.plotConfiguration
  };
}

function reviewRowRecord(row: ReviewFlowRow): Record<string, string> {
  return {
    序号: row.recordNo,
    创建者: row.creator,
    创建时间: row.createdDate,
    记录状态: row.recordStatus,
    设计文档: row.designDocument,
    功能区名称: row.functionAreaName,
    单元名称: row.unitName,
    装置: row.deviceName,
    专业分类: row.disciplineName,
    数字化PMC: row.digitalPmc,
    专业工程师: row.specialtyEngineer,
    专业负责人: row.professionalLeader,
    装置专业键: normalizeJoinKey(row.deviceName, row.disciplineName)
  };
}

async function parseStructuredDerivedSources(files: GenericFileInput[]): Promise<GenericParseResult> {
  const issues: IssueRow[] = [];
  const excelFiles = files.map((file, index) => ({ file, index })).filter(({ file }) => isExcelFile(file));
  const pdfFiles = files.map((file, index) => ({ file, index })).filter(({ file }) => isPdfFile(file));
  const matrix = pickScoredFile(excelFiles, scorePresetMatrixFile);
  const pdf = pickScoredFile(pdfFiles, scorePresetPdfFile);

  if (!matrix) {
    return {
      tables: [],
      issues: [errorIssue("structured_matrix_missing", "结构化解析未找到可识别的矩阵 Excel", { source: "generic" })]
    };
  }
  if (!pdf) {
    return {
      tables: [],
      issues: [errorIssue("structured_pdf_missing", "结构化解析未找到可识别的主项或分工 PDF", { source: "generic" })]
    };
  }

  if (excelFiles.length > 1) {
    issues.push(issue("structured_matrix_selected", "已按文件名选择矩阵 Excel", { source: matrix.file.fileName }));
  }
  if (pdfFiles.length > 1) {
    issues.push(issue("structured_pdf_selected", "已按文件名选择主项或分工 PDF", { source: pdf.file.fileName }));
  }

  try {
    const result = await generatePreset(matrix.file.buffer, pdf.file.buffer, {
      rootPath: "文档文件"
    });
    issues.push(...result.issues);

    const matrixRows = result.matrix.rows.map((row) => ({
      源行: String(row.sourceRow),
      类别: row.category,
      装置名称: row.deviceName,
      专业分类: row.disciplineName,
      数字化PMC: row.digitalPmc,
      专业工程师: row.specialtyEngineer,
      专业负责人: row.professionalLeader,
      归档负责人: row.archiveOwner ?? "",
      装置匹配键: normalizeDeviceName(row.deviceName),
      装置专业键: normalizeJoinKey(row.deviceName, row.disciplineName)
    }));

    const deviceRows = result.project.devices.map((device) => ({
      主项号: device.code,
      装置名称: device.name,
      装置: displayCodeName(device.code, device.name),
      功能区编码: device.areaCode,
      功能区名称: device.areaName,
      功能区: displayArea(device.areaCode, device.areaName),
      单元号: device.unitCode,
      单元名称: device.unitName,
      单元: displayCodeName(device.unitCode, device.unitName),
      设计单位: device.designCompany,
      源页: String(device.sourcePage),
      原始行: device.rawLine,
      单独装置文件夹: device.separateFolder ? "是" : "否",
      装置匹配键: normalizeDeviceName(device.name)
    }));

    const unitRows = result.project.units.map((unit) => ({
      单元号: unit.code,
      单元名称: unit.name,
      单元: displayCodeName(unit.code, unit.name),
      功能区编码: unit.areaCode,
      功能区名称: unit.areaName,
      功能区: displayArea(unit.areaCode, unit.areaName),
      设计单位: unit.designCompany,
      源页: String(unit.sourcePage),
      原始行: unit.rawLine,
      单元匹配键: normalizeDeviceName(unit.name)
    }));

    const folderRows = result.folderRows.map(folderRowRecord);
    const reviewRows = result.reviewRows.map(reviewRowRecord);
    const issueRows = issues.map((item) => ({
      类型: item.type,
      级别: item.severity,
      说明: item.message,
      来源: item.source ?? "",
      源行: item.row == null ? "" : String(item.row),
      关键值: item.key ?? "",
      原始值: item.value ?? ""
    }));

    return {
      tables: [
        derivedTable(
          `derived:structured:matrix:${matrix.index + 1}:${matrix.file.fileName}:矩阵明细`,
          matrix.file.fileName,
          "矩阵明细",
          ["源行", "类别", "装置名称", "专业分类", "数字化PMC", "专业工程师", "专业负责人", "归档负责人", "装置匹配键", "装置专业键"],
          matrixRows
        ),
        derivedTable(
          `derived:structured:pdf:${pdf.index + 1}:${pdf.file.fileName}:PDF台账`,
          pdf.file.fileName,
          "PDF台账",
          ["主项号", "装置名称", "装置", "功能区编码", "功能区名称", "功能区", "单元号", "单元名称", "单元", "设计单位", "源页", "原始行", "单独装置文件夹", "装置匹配键"],
          deviceRows
        ),
        derivedTable(
          `derived:structured:pdf:${pdf.index + 1}:${pdf.file.fileName}:分组台账`,
          pdf.file.fileName,
          "分组台账",
          ["单元号", "单元名称", "单元", "功能区编码", "功能区名称", "功能区", "设计单位", "源页", "原始行", "单元匹配键"],
          unitRows
        ),
        derivedTable(
          "derived:structured:output:层级目录",
          "结构化生成结果",
          "层级目录结果",
          ["文件夹路径", "名称", "更新名称", "编码", "创建时间", "更新时间", "创建者", "更新者", "所属单位", "层级结构", "装置", "单元", "设计单位", "出图配置"],
          folderRows
        ),
        derivedTable(
          "derived:structured:output:矩阵组合",
          "结构化生成结果",
          "矩阵组合结果",
          ["序号", "创建者", "创建时间", "记录状态", "设计文档", "功能区名称", "单元名称", "装置", "专业分类", "数字化PMC", "专业工程师", "专业负责人", "装置专业键"],
          reviewRows
        ),
        derivedTable(
          "derived:structured:output:问题清单",
          "结构化生成结果",
          "结构化问题清单",
          ["类型", "级别", "说明", "来源", "源行", "关键值", "原始值"],
          issueRows,
          "issue"
        )
      ],
      issues
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : "结构化解析失败";
    return {
      tables: [],
      issues: [errorIssue("structured_parse_failed", message, { source: "generic" })]
    };
  }
}

async function parseGenericSourcesDetailed(files: GenericFileInput[], config: GenericCombineConfig = {}): Promise<GenericParseResult> {
  const tables: GenericTable[] = [];
  const issues: IssueRow[] = [];
  const sourceConfigs = config.includeAuxiliarySheets
    ? [{ sourceId: "__include_auxiliary_sheets__" }, ...(config.sourceConfigs ?? [])]
    : (config.sourceConfigs ?? []);
  for (const [index, file] of files.entries()) {
    if (isExcelFile(file)) {
      tables.push(...(await parseExcelSource(file, sourceConfigs, index)));
    } else if (isPdfFile(file)) {
      tables.push(...(await parsePdfSource(file, index)));
    }
  }
  if (config.structuredExtraction === "auto" || config.businessPreset === "preset") {
    const derived = await parseStructuredDerivedSources(files);
    tables.push(...derived.tables);
    issues.push(...derived.issues);
  }
  return { tables, issues };
}

export async function parseGenericSources(files: GenericFileInput[], config: GenericCombineConfig = {}): Promise<GenericTable[]> {
  return (await parseGenericSourcesDetailed(files, config)).tables;
}

function defaultOutputColumns(base: GenericTable, joins: GenericJoinRule[], tableMap: Map<string, GenericTable>): GenericOutputColumn[] {
  const columns: GenericOutputColumn[] = base.source.headers.map((header) => ({
    label: header,
    sourceId: base.source.id,
    column: header
  }));
  for (const join of joins) {
    const table = tableMap.get(join.sourceId);
    if (!table) continue;
    const prefix = join.prefix || table.source.sheetName || table.source.fileName;
    for (const header of table.source.headers) {
      columns.push({ label: `${prefix}.${header}`, sourceId: table.source.id, column: header });
    }
  }
  return columns;
}

function joinDuplicateMessage(strategy: NonNullable<GenericJoinRule["duplicateStrategy"]>): string {
  if (strategy === "expand") return "关联表中存在重复键，将按匹配记录展开为多行";
  if (strategy === "merge") return "关联表中存在重复键，将多条匹配记录合并到同一行";
  return "关联表中存在重复键，将使用第一条匹配记录";
}

function mergeRows(rows: Record<string, string>[], separator: string): Record<string, string> {
  const merged: Record<string, string> = {};
  const headers = Object.keys(rows[0] ?? {});
  for (const header of headers) {
    const values = rows.map((row) => cleanText(row[header])).filter(Boolean);
    merged[header] = Array.from(new Set(values)).join(separator);
  }
  return merged;
}

function sourceTokenPrefixes(source: GenericSourceSummary): string[] {
  return [
    source.id,
    source.sheetName,
    source.fileName.replace(/\.[^.]+$/, "")
  ]
    .map((value) => cleanText(value))
    .filter(Boolean)
    .sort((a, b) => b.length - a.length);
}

function hasOwnValue(row: Record<string, string>, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(row, key);
}

function prefixedTokenColumn(token: string, source: GenericSourceSummary): string | undefined {
  for (const prefix of sourceTokenPrefixes(source)) {
    if (token.startsWith(`${prefix}::`)) return token.slice(prefix.length + 2);
    if (token.startsWith(`${prefix}.`)) return token.slice(prefix.length + 1);
  }
  return undefined;
}

function resolveTemplateToken(
  token: string,
  data: Map<string, Record<string, string>>,
  tableMap: Map<string, GenericTable>,
  preferredSourceId?: string
): string {
  const key = cleanText(token);
  if (!key) return "";

  const preferredRow = preferredSourceId ? data.get(preferredSourceId) : undefined;
  if (preferredRow && hasOwnValue(preferredRow, key)) return cleanText(preferredRow[key]);

  for (const row of data.values()) {
    if (hasOwnValue(row, key)) return cleanText(row[key]);
  }

  for (const [sourceId, row] of data.entries()) {
    const source = tableMap.get(sourceId)?.source;
    if (!source) continue;
    const column = prefixedTokenColumn(key, source);
    if (column && hasOwnValue(row, column)) return cleanText(row[column]);
  }

  return "";
}

function renderTemplate(
  template: string,
  data: Map<string, Record<string, string>>,
  tableMap: Map<string, GenericTable>,
  preferredSourceId?: string
): string {
  return template.replace(/\{([^{}]+)\}/g, (_match, token: string) => resolveTemplateToken(token, data, tableMap, preferredSourceId));
}

function outputColumnValue(
  data: Map<string, Record<string, string>>,
  column: GenericOutputColumn,
  tableMap: Map<string, GenericTable>
): string {
  if (column.constant != null) return column.constant;
  if (column.template != null) return cleanText(renderTemplate(column.template, data, tableMap, column.sourceId));
  if (!column.sourceId || !column.column) return "";
  return cleanText(data.get(column.sourceId)?.[column.column] ?? "");
}

function styleGenericHeader(row: ExcelJS.Row): void {
  row.font = { bold: true, color: { argb: "FFFFFFFF" } };
  row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF243447" } };
}

function safeSheetName(rawName: string, used: Set<string>): string {
  const fallback = "结果";
  const clean = (cleanText(rawName) || fallback).replace(/[\\/?*\[\]:]/g, "_").slice(0, 31) || fallback;
  let name = clean;
  let suffix = 2;
  while (used.has(name)) {
    const marker = `_${suffix}`;
    name = `${clean.slice(0, Math.max(1, 31 - marker.length))}${marker}`;
    suffix += 1;
  }
  used.add(name);
  return name;
}

function columnValue(row: Record<string, string>, column: GenericOutputColumn, sourceId: string, tableMap: Map<string, GenericTable>): string {
  if (column.constant != null) return column.constant;
  if (column.template != null) return cleanText(renderTemplate(column.template, new Map([[sourceId, row]]), tableMap, sourceId));
  if (column.sourceId && column.sourceId !== sourceId) return "";
  return cleanText(column.column ? row[column.column] : "");
}

function addGenericRowsSheet(
  workbook: ExcelJS.Workbook,
  usedSheetNames: Set<string>,
  sheetName: string,
  sourceId: string,
  rows: Record<string, string>[],
  columns: GenericOutputColumn[],
  tableMap: Map<string, GenericTable> = new Map()
): void {
  const sheet = workbook.addWorksheet(safeSheetName(sheetName, usedSheetNames), { views: [{ state: "frozen", ySplit: 1 }] });
  const headers = columns.length ? columns.map((column) => column.label) : ["无输出字段"];
  sheet.addRow(headers);
  styleGenericHeader(sheet.getRow(1));

  for (const row of rows) {
    sheet.addRow(columns.map((column) => columnValue(row, column, sourceId, tableMap)));
  }

  headers.forEach((header, index) => {
    sheet.getColumn(index + 1).width = Math.min(42, Math.max(14, cleanText(header).length + 6));
  });
}

function addIssueOverviewSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, issues: IssueRow[]): void {
  const issueSheet = workbook.addWorksheet(safeSheetName("问题清单", usedSheetNames), { views: [{ state: "frozen", ySplit: 1 }] });
  issueSheet.addRow(["类型", "级别", "说明", "来源", "源行", "关键值", "原始值"]);
  styleGenericHeader(issueSheet.getRow(1));
  if (!issues.length) {
    issueSheet.addRow(["none", "info", "未发现问题", "", "", "", ""]);
  } else {
    for (const item of issues) {
      issueSheet.addRow([item.type, item.severity, item.message, item.source ?? "", item.row ?? "", item.key ?? "", item.value ?? ""]);
    }
  }
  issueSheet.columns = [
    { width: 28 },
    { width: 12 },
    { width: 58 },
    { width: 22 },
    { width: 12 },
    { width: 36 },
    { width: 48 }
  ];
}

function addSourceOverviewSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, sources: GenericSourceSummary[]): void {
  const sourceSheet = workbook.addWorksheet(safeSheetName("数据源概览", usedSheetNames));
  sourceSheet.addRow(["数据源", "文件", "类型", "Sheet", "行数", "列数", "字段"]);
  styleGenericHeader(sourceSheet.getRow(1));
  for (const source of sources) {
    sourceSheet.addRow([source.id, source.fileName, source.kind, source.sheetName, source.rowCount, source.columnCount, source.headers.join(",")]);
  }
  sourceSheet.columns = [{ width: 44 }, { width: 34 }, { width: 12 }, { width: 20 }, { width: 12 }, { width: 12 }, { width: 80 }];
}

export async function combineGeneric(
  files: GenericFileInput[],
  config: GenericCombineConfig = {},
  options: { previewLimit?: number; issueLimit?: number } = {}
): Promise<GenericPreviewResult> {
  const parsed = await parseGenericSourcesDetailed(files, config);
  const tables = parsed.tables;
  const issues: IssueRow[] = [...parsed.issues];
  const sources = tables.map((table) => table.source);
  const tableMap = new Map(tables.map((table) => [table.source.id, table]));
  const base = config.baseSourceId ? tableMap.get(config.baseSourceId) : tables[0];
  if (!base) {
    const message = config.baseSourceId ? "指定的主表数据源不存在，请重新解析数据源后选择主表" : "没有可用的数据源，请上传 .xlsx 或 .pdf 文件";
    return {
      sources,
      outputPreview: [],
      issues: [
        issue(config.baseSourceId ? "base_source_missing" : "no_source", message, {
          severity: "error",
          key: config.baseSourceId
        })
      ],
      stats: { sourceCount: sources.length, outputRowCount: 0, issueCount: 1 }
    };
  }

  const joins = config.joins ?? [];
  const outputColumns = config.outputColumns?.length ? config.outputColumns : defaultOutputColumns(base, joins, tableMap);
  let workingRows = base.rows.map((row, rowIndex) => ({
    rowIndex,
    data: new Map<string, Record<string, string>>([[base.source.id, row]])
  }));

  for (const join of joins) {
    const right = tableMap.get(join.sourceId);
    const duplicateStrategy = join.duplicateStrategy ?? "first";
    const separator = join.multiValueSeparator || ",";
    if (!right) {
      issues.push(issue("join_source_missing", "关联数据源不存在，已跳过", { key: join.sourceId }));
      continue;
    }
    if (!base.source.headers.includes(join.leftKey) || !right.source.headers.includes(join.rightKey)) {
      issues.push(
        issue("join_key_missing", "关联键字段不存在，已跳过该关联", {
          key: `${join.leftKey} -> ${join.rightKey}`,
          source: right.source.id
        })
      );
      continue;
    }

    const index = new Map<string, Record<string, string>[]>();
    for (const row of right.rows) {
      const key = normalizeKey(row[join.rightKey]);
      if (!key) continue;
      const group = index.get(key) ?? [];
      group.push(row);
      index.set(key, group);
    }
    for (const [key, rows] of index.entries()) {
      if (rows.length > 1) {
        issues.push(issue("join_duplicate_key", joinDuplicateMessage(duplicateStrategy), { source: right.source.id, key, value: duplicateStrategy }));
      }
    }

    const nextRows: typeof workingRows = [];
    for (const row of workingRows) {
      const left = row.data.get(base.source.id);
      const key = normalizeKey(left?.[join.leftKey]);
      const matches = key ? index.get(key) : undefined;
      if (!matches?.length) {
        if (join.joinType === "left") {
          nextRows.push(row);
        }
        if (key) {
          issues.push(issue("join_unmatched", "主表记录未匹配到关联表", { source: right.source.id, key }));
        }
      } else if (duplicateStrategy === "expand") {
        for (const match of matches) {
          nextRows.push({
            rowIndex: row.rowIndex,
            data: new Map([...row.data.entries(), [right.source.id, match]])
          });
        }
      } else if (duplicateStrategy === "merge") {
        nextRows.push({
          rowIndex: row.rowIndex,
          data: new Map([...row.data.entries(), [right.source.id, mergeRows(matches, separator)]])
        });
      } else {
        nextRows.push({
          rowIndex: row.rowIndex,
          data: new Map([...row.data.entries(), [right.source.id, matches[0]]])
        });
      }
    }
    workingRows = nextRows;
  }

  const output = workingRows.map((row) => {
    const out: Record<string, string> = {};
    for (const column of outputColumns) {
      out[column.label] = outputColumnValue(row.data, column, tableMap);
    }
    return out;
  });
  const issueLimit = options.issueLimit ?? 500;

  return {
    sources,
    outputPreview: output.slice(0, options.previewLimit ?? PREVIEW_LIMIT),
    issues: issueLimit === Number.MAX_SAFE_INTEGER ? issues : issues.slice(0, issueLimit),
    stats: {
      sourceCount: sources.length,
      outputRowCount: output.length,
      issueCount: issues.length
    }
  };
}

export async function writeGenericWorkbook(files: GenericFileInput[], config: GenericCombineConfig = {}): Promise<Buffer> {
  if (config.resultSheets?.length) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "excel-composer";
    workbook.created = new Date();
    workbook.modified = new Date();
    const usedSheetNames = new Set<string>();
    const parsed = await parseGenericSourcesDetailed(files, config);
    const tableMap = new Map(parsed.tables.map((table) => [table.source.id, table]));
    const issues = [...parsed.issues];

    for (const resultSheet of config.resultSheets) {
      const table = tableMap.get(resultSheet.sourceId);
      if (!table) {
        issues.push(issue("result_sheet_source_missing", "输出任务的数据源不存在，已跳过该 Sheet", {
          severity: "error",
          key: resultSheet.sourceId
        }));
        continue;
      }
      const columns = resultSheet.outputColumns?.length
        ? resultSheet.outputColumns
        : table.source.headers.map((header) => ({
          label: header,
          sourceId: table.source.id,
          column: header
        }));
      addGenericRowsSheet(
        workbook,
        usedSheetNames,
        resultSheet.sheetName || table.source.sheetName,
        table.source.id,
        table.rows,
        columns,
        tableMap
      );
    }

    if (!workbook.worksheets.length) {
      addGenericRowsSheet(workbook, usedSheetNames, "组合结果", "", [], []);
    }
    addIssueOverviewSheet(workbook, usedSheetNames, issues);
    addSourceOverviewSheet(workbook, usedSheetNames, parsed.tables.map((table) => table.source));
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
  }

  const full = await combineGeneric(files, config, {
    previewLimit: Number.MAX_SAFE_INTEGER,
    issueLimit: Number.MAX_SAFE_INTEGER
  });
  if (config.templateExport?.sourceId) {
    const templateOutput = await writeGenericTemplateWorkbook(files, config, full);
    if (templateOutput) return templateOutput;
  }
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "excel-composer";
  workbook.created = new Date();
  workbook.modified = new Date();
  const usedSheetNames = new Set<string>();

  const headers = full.outputPreview[0] ? Object.keys(full.outputPreview[0]) : (config.outputColumns ?? []).map((col) => col.label);
  const outputColumns = (headers.length ? headers : ["无输出字段"]).map((header) => ({
    label: header,
    column: header
  }));
  addGenericRowsSheet(workbook, usedSheetNames, "组合结果", "", full.outputPreview, outputColumns);
  addIssueOverviewSheet(workbook, usedSheetNames, full.issues);
  addSourceOverviewSheet(workbook, usedSheetNames, full.sources);

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

async function writeGenericTemplateWorkbook(
  files: GenericFileInput[],
  config: GenericCombineConfig,
  full: GenericPreviewResult
): Promise<Buffer | undefined> {
  const templateConfig = config.templateExport;
  if (!templateConfig?.sourceId) return undefined;

  const allSources = await parseGenericSources(files, {
    ...config,
    includeAuxiliarySheets: true
  });
  const templateTable = allSources.find((table) => table.source.id === templateConfig.sourceId);
  if (!templateTable || templateTable.source.kind !== "excel") return undefined;

  const templateFile = files.find((file) => file.fileName === templateTable.source.fileName);
  if (!templateFile) return undefined;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateFile.buffer as any);
  workbook.creator = "excel-composer";
  workbook.modified = new Date();

  const sheet = workbook.getWorksheet(templateTable.source.sheetName);
  if (!sheet) return undefined;

  const sourceConfig = config.sourceConfigs?.find((item) => item.sourceId === templateTable.source.id);
  const headerRowNo = templateConfig.headerRow ?? sourceConfig?.headerRow ?? templateTable.source.defaultHeaderRow;
  const dataStartRow = templateConfig.dataStartRow ?? sourceConfig?.dataStartRow ?? templateTable.source.defaultDataStartRow;
  const headerRow = sheet.getRow(headerRowNo);
  const templateHeaders = rowValues(headerRow, sheet.columnCount);
  const outputHeaders = full.outputPreview[0]
    ? Object.keys(full.outputPreview[0])
    : (config.outputColumns ?? []).map((column) => column.label);
  const outputHeaderSet = new Set(outputHeaders);
  const columnMap = templateHeaders.map((header) => (header && outputHeaderSet.has(header) ? header : ""));

  const styleRow = sheet.getRow(dataStartRow);
  const cellTemplates = templateHeaders.map((_, index) => {
    const cell = styleRow.getCell(index + 1);
    return {
      style: cloneExcelValue(cell.style),
      dataValidation: cloneExcelValue(cell.dataValidation),
      numFmt: cell.numFmt
    };
  });

  const rowsToRemove = Math.max(0, sheet.rowCount - dataStartRow + 1);
  if (rowsToRemove > 0) {
    sheet.spliceRows(dataStartRow, rowsToRemove);
  }

  full.outputPreview.forEach((outputRow, rowIndex) => {
    const row = sheet.getRow(dataStartRow + rowIndex);
    columnMap.forEach((outputHeader, columnIndex) => {
      const cell = row.getCell(columnIndex + 1);
      const template = cellTemplates[columnIndex];
      if (template?.style) cell.style = cloneExcelValue(template.style);
      if (template?.dataValidation) cell.dataValidation = cloneExcelValue(template.dataValidation);
      if (template?.numFmt) cell.numFmt = template.numFmt;
      cell.value = outputHeader ? (outputRow[outputHeader] ?? "") : "";
    });
    row.commit();
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
