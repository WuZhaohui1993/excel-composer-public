import ExcelJS from "exceljs";
import type { FolderAttributeRow, GenerationResult, IssueRow, ReviewFlowRow } from "./types.js";

const folderHeaders = [
  "文件夹路径",
  "名称",
  "更新名称",
  "编码",
  "创建时间",
  "更新时间",
  "创建者",
  "更新者",
  "所属单位",
  "层级结构",
  "装置",
  "单元",
  "设计单位",
  "出图配置"
];

const reviewHeaders = [
  "序号",
  "创建者",
  "创建时间",
  "记录状态",
  "设计文档",
  "功能区名称",
  "单元名称",
  "装置",
  "专业分类",
  "数字化PMC",
  "专业工程师",
  "专业负责人"
];

const issueHeaders = ["类型", "级别", "说明", "来源", "源行", "关键值", "原始值"];

function styleHeader(row: ExcelJS.Row): void {
  row.font = { bold: true, color: { argb: "FFFFFFFF" } };
  row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF243447" } };
  row.alignment = { vertical: "middle", horizontal: "center" };
}

function setColumns(sheet: ExcelJS.Worksheet, widths: number[]): void {
  widths.forEach((width, index) => {
    sheet.getColumn(index + 1).width = width;
  });
}

function addFolderSheet(workbook: ExcelJS.Workbook, rows: FolderAttributeRow[]): void {
  const sheet = workbook.addWorksheet("文件夹属性", { views: [{ state: "frozen", ySplit: 1 }] });
  sheet.addRow(folderHeaders);
  styleHeader(sheet.getRow(1));
  for (const row of rows) {
    sheet.addRow([
      row.folderPath,
      row.name,
      row.updateName,
      row.code,
      row.createDate,
      row.updateDate,
      row.createBy,
      row.updateBy,
      row.belongingUnit,
      row.levelName,
      row.device,
      row.unit,
      row.designCompany,
      row.plotConfiguration
    ]);
  }
  sheet.autoFilter = { from: "A1", to: "N1" };
  setColumns(sheet, [72, 28, 18, 14, 20, 20, 16, 16, 18, 14, 26, 26, 18, 18]);
}

function addReviewSheet(workbook: ExcelJS.Workbook, rows: ReviewFlowRow[]): void {
  const sheet = workbook.addWorksheet("设计文档审阅流程", { views: [{ state: "frozen", ySplit: 1 }] });
  sheet.addRow(reviewHeaders);
  styleHeader(sheet.getRow(1));
  for (const row of rows) {
    sheet.addRow([
      row.recordNo,
      row.creator,
      row.createdDate,
      row.recordStatus,
      row.designDocument,
      row.functionAreaName,
      row.unitName,
      row.deviceName,
      row.disciplineName,
      row.digitalPmc,
      row.specialtyEngineer,
      row.professionalLeader
    ]);
  }
  sheet.autoFilter = { from: "A1", to: "L1" };
  setColumns(sheet, [16, 12, 18, 12, 16, 18, 28, 28, 24, 16, 34, 20]);
}

function addIssueSheet(workbook: ExcelJS.Workbook, rows: IssueRow[]): void {
  const sheet = workbook.addWorksheet("问题清单", { views: [{ state: "frozen", ySplit: 1 }] });
  sheet.addRow(issueHeaders);
  styleHeader(sheet.getRow(1));
  if (rows.length === 0) {
    sheet.addRow(["none", "info", "未发现问题", "", "", "", ""]);
  } else {
    for (const issue of rows) {
      sheet.addRow([issue.type, issue.severity, issue.message, issue.source ?? "", issue.row ?? "", issue.key ?? "", issue.value ?? ""]);
    }
  }
  sheet.autoFilter = { from: "A1", to: "G1" };
  setColumns(sheet, [28, 12, 58, 22, 12, 36, 48]);
}

export async function writePresetWorkbook(result: GenerationResult): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "excel-composer";
  workbook.created = new Date();
  workbook.modified = new Date();

  addFolderSheet(workbook, result.folderRows);
  addReviewSheet(workbook, result.reviewRows);
  addIssueSheet(workbook, result.issues);

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
