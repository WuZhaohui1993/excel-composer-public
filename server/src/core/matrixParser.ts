import ExcelJS from "exceljs";
import type { IssueRow, MatrixRow, ParsedMatrix } from "./types.js";
import { cleanText, normalizeDisciplineName, normalizeJoinKey, normalizePeople } from "./normalizers.js";

function cellText(row: ExcelJS.Row, col: number): string {
  return cleanText(row.getCell(col).value);
}

function uniqueInOrder(values: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const value of values) {
    if (!value || seen.has(value)) continue;
    seen.add(value);
    result.push(value);
  }
  return result;
}

export async function parseReviewMatrix(buffer: Buffer, options: { sheetName?: string; dataStartRow?: number } = {}): Promise<ParsedMatrix> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as any);

  const worksheet = options.sheetName
    ? workbook.getWorksheet(options.sheetName)
    : workbook.getWorksheet("设计文档审阅流程") ?? workbook.worksheets[0];

  if (!worksheet) {
    return {
      sheetName: "",
      rows: [],
      disciplines: [],
      devices: [],
      issues: [{ type: "matrix_sheet_missing", severity: "error", message: "未找到审阅流程矩阵工作表", source: "matrix" }]
    };
  }

  const issues: IssueRow[] = [];
  const rows: MatrixRow[] = [];
  const duplicateKeys = new Map<string, number>();
  let currentCategory = "";
  let currentDeviceName = "";
  const dataStartRow = options.dataStartRow ?? 4;

  for (let rowNumber = dataStartRow; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const category = cellText(row, 1) || currentCategory;
    const deviceName = cellText(row, 2) || currentDeviceName;
    const disciplineName = normalizeDisciplineName(cellText(row, 3));

    if (category) currentCategory = category;
    if (deviceName) currentDeviceName = deviceName;
    if (!deviceName && !disciplineName) continue;
    if (!deviceName || !disciplineName) {
      issues.push({
        type: "matrix_key_missing",
        severity: "warning",
        message: "矩阵行缺少装置名称或专业分类，已跳过",
        source: worksheet.name,
        row: rowNumber,
        value: [deviceName, disciplineName].filter(Boolean).join(" / ")
      });
      continue;
    }

    const key = normalizeJoinKey(deviceName, disciplineName);
    const firstRow = duplicateKeys.get(key);
    if (firstRow) {
      issues.push({
        type: "matrix_duplicate_key",
        severity: "warning",
        message: `矩阵中存在重复装置+专业组合，首次出现在第 ${firstRow} 行`,
        source: worksheet.name,
        row: rowNumber,
        key
      });
    } else {
      duplicateKeys.set(key, rowNumber);
    }

    rows.push({
      sourceRow: rowNumber,
      category,
      deviceName,
      disciplineName,
      digitalPmc: normalizePeople(cellText(row, 5)),
      specialtyEngineer: normalizePeople(cellText(row, 6)),
      professionalLeader: normalizePeople(cellText(row, 7)),
      archiveOwner: normalizePeople(cellText(row, 8))
    });
  }

  return {
    sheetName: worksheet.name,
    rows,
    disciplines: uniqueInOrder(rows.map((row) => row.disciplineName)),
    devices: uniqueInOrder(rows.map((row) => row.deviceName)),
    issues
  };
}
