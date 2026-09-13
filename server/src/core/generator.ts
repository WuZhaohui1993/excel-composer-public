import type {
  DeviceNode,
  FolderAttributeRow,
  GenerateConfig,
  GenerationResult,
  IssueRow,
  ParsedProject,
  ReviewFlowRow,
  UnitNode
} from "./types.js";
import { parseReviewMatrix } from "./matrixParser.js";
import { parseProjectPdf } from "./pdfParser.js";
import {
  AREA_NAMES,
  compareByCodeName,
  compactText,
  displayArea,
  displayCodeName,
  normalizeDeviceName,
  normalizeJoinKey
} from "./normalizers.js";

const DEFAULT_ROOT = "示例项目/文档文件";

function issue(type: string, message: string, extra: Partial<IssueRow> = {}): IssueRow {
  return { type, severity: "warning", message, ...extra };
}

function unitDisplay(unit: Pick<UnitNode, "code" | "name">): string {
  return displayCodeName(unit.code, unit.name);
}

function deviceDisplay(device: Pick<DeviceNode, "code" | "name">): string {
  return displayCodeName(device.code, device.name);
}

function findDevice(project: ParsedProject, deviceName: string): DeviceNode | undefined {
  const key = normalizeDeviceName(deviceName);
  const exact = project.devices.find((device) => normalizeDeviceName(device.name) === key);
  if (exact) return exact;
  const fuzzy = project.devices.find((device) => {
    const normalized = normalizeDeviceName(device.name);
    return normalized.includes(key) || key.includes(normalized);
  });
  if (fuzzy) return fuzzy;

  const aliases: Record<string, DeviceNode> = {
    装置变电所: {
      code: "30302",
      name: "装置变电所",
      areaCode: "30",
      areaName: "公用工程",
      unitCode: "",
      unitName: "",
      designCompany: "",
      sourcePage: 0,
      sourceY: 0,
      rawLine: "业务别名：PDF 中 30302/30307 装置变电所为跨行主项",
      separateFolder: false
    },
    固废暂存设施: {
      code: "50163",
      name: "固废暂存设施",
      areaCode: "50",
      areaName: "物流设施",
      unitCode: "",
      unitName: "",
      designCompany: "",
      sourcePage: 0,
      sourceY: 0,
      rawLine: "业务别名：PDF 名称为固废暂存库，平台口径为固废暂存设施",
      separateFolder: false
    },
    行政生活设施: {
      code: "60101",
      name: "行政生活设施",
      areaCode: "60",
      areaName: "基础设施",
      unitCode: "",
      unitName: "",
      designCompany: "",
      sourcePage: 0,
      sourceY: 0,
      rawLine: "业务别名：PDF 名称为生产楼，平台口径为行政生活设施",
      separateFolder: false
    }
  };
  return aliases[key];
}

function findUnit(project: ParsedProject, deviceName: string): UnitNode | undefined {
  const key = normalizeDeviceName(deviceName);
  return project.units.find((unit) => normalizeDeviceName(unit.name) === key);
}

function folderPathForDevice(rootPath: string, device: DeviceNode): string {
  const area = displayArea(device.areaCode, device.areaName);
  const unit = unitDisplay({ code: device.unitCode, name: device.unitName });
  return device.separateFolder ? `${rootPath}/${area}/${unit}/${deviceDisplay(device)}` : `${rootPath}/${area}/${unit}`;
}

function buildFolderRows(project: ParsedProject, disciplines: string[], config: GenerateConfig): FolderAttributeRow[] {
  const rootPath = config.rootPath || DEFAULT_ROOT;
  const createBy = config.folderCreator ?? "";
  const rows: FolderAttributeRow[] = [];
  const push = (row: FolderAttributeRow) => rows.push(row);

  for (const [areaCode, areaName] of Object.entries(AREA_NAMES)) {
    push({
      folderPath: rootPath,
      name: displayArea(areaCode, areaName),
      updateName: "",
      code: "",
      createDate: "",
      updateDate: "",
      createBy,
      updateBy: "",
      belongingUnit: "",
      levelName: "功能区",
      device: "",
      unit: "",
      designCompany: "",
      plotConfiguration: ""
    });
  }

  for (const unit of project.units) {
    push({
      folderPath: `${rootPath}/${displayArea(unit.areaCode, unit.areaName)}`,
      name: unitDisplay(unit),
      updateName: "",
      code: "",
      createDate: "",
      updateDate: "",
      createBy,
      updateBy: "",
      belongingUnit: "",
      levelName: "单元",
      device: "",
      unit: unitDisplay(unit),
      designCompany: unit.designCompany,
      plotConfiguration: unit.designCompany
    });
  }

  for (const device of project.devices) {
    if (device.separateFolder) {
      push({
        folderPath: `${rootPath}/${displayArea(device.areaCode, device.areaName)}/${unitDisplay({ code: device.unitCode, name: device.unitName })}`,
        name: deviceDisplay(device),
        updateName: "",
        code: "",
        createDate: "",
        updateDate: "",
        createBy,
        updateBy: "",
        belongingUnit: "",
        levelName: "装置",
        device: deviceDisplay(device),
        unit: unitDisplay({ code: device.unitCode, name: device.unitName }),
        designCompany: device.designCompany,
        plotConfiguration: device.designCompany
      });
    }
  }

  for (const device of project.devices) {
    const path = folderPathForDevice(rootPath, device);
    for (const discipline of disciplines) {
      push({
        folderPath: path,
        name: discipline,
        updateName: "",
        code: "",
        createDate: "",
        updateDate: "",
        createBy,
        updateBy: "",
        belongingUnit: "",
        levelName: "",
        device: deviceDisplay(device),
        unit: unitDisplay({ code: device.unitCode, name: device.unitName }),
        designCompany: device.designCompany,
        plotConfiguration: device.designCompany
      });
    }
  }

  return rows;
}

function buildReviewRows(project: ParsedProject, matrixRows: GenerationResult["matrix"]["rows"], issues: IssueRow[], config: GenerateConfig): ReviewFlowRow[] {
  const recordPrefix = config.recordPrefix || "sjwj";
  const designDocumentName = config.designDocumentName || "文档文件";
  const rows: ReviewFlowRow[] = [];
  const seen = new Set<string>();

  matrixRows.forEach((matrixRow, index) => {
    const device = findDevice(project, matrixRow.deviceName);
    const unit = device ? project.units.find((candidate) => candidate.code === device.unitCode && candidate.name === device.unitName) : findUnit(project, matrixRow.deviceName);
    const key = normalizeJoinKey(matrixRow.deviceName, matrixRow.disciplineName);
    if (seen.has(key)) {
      issues.push(issue("review_duplicate_key", "审阅流程输出存在重复装置+专业组合", { source: "matrix", row: matrixRow.sourceRow, key }));
    }
    seen.add(key);

    if (!device && !unit) {
      issues.push(
        issue("review_pdf_match_missing", "矩阵装置未能在 PDF 主项表中匹配到功能区/单元信息", {
          source: "matrix",
          row: matrixRow.sourceRow,
          key,
          value: matrixRow.deviceName
        })
      );
    }

    const area = device
      ? displayArea(device.areaCode, device.areaName)
      : unit
        ? displayArea(unit.areaCode, unit.areaName)
        : "";
    const resolvedUnit = device
      ? unitDisplay({ code: device.unitCode, name: device.unitName })
      : unit
        ? unitDisplay(unit)
        : "";
    const resolvedDevice = device
      ? deviceDisplay(device)
      : unit
        ? unitDisplay(unit)
        : compactText(matrixRow.deviceName);

    rows.push({
      recordNo: `${recordPrefix}${String(index + 1).padStart(6, "0")}`,
      creator: "",
      createdDate: "",
      recordStatus: "",
      designDocument: designDocumentName,
      functionAreaName: area,
      unitName: device?.separateFolder ? resolvedUnit : resolvedUnit,
      deviceName: resolvedDevice,
      disciplineName: matrixRow.disciplineName,
      digitalPmc: matrixRow.digitalPmc,
      specialtyEngineer: matrixRow.specialtyEngineer,
      professionalLeader: matrixRow.professionalLeader
    });
  });

  return rows;
}

export async function generatePreset(matrixBuffer: Buffer, pdfBuffer: Buffer, config: GenerateConfig = {}): Promise<GenerationResult> {
  const matrix = await parseReviewMatrix(matrixBuffer, {
    sheetName: config.matrixSheetName,
    dataStartRow: config.matrixDataStartRow
  });
  const project = await parseProjectPdf(pdfBuffer);
  const issues: IssueRow[] = [...matrix.issues, ...project.issues];
  const folderRows = buildFolderRows(project, matrix.disciplines, config);
  const reviewRows = buildReviewRows(project, matrix.rows, issues, config);

  if (Math.abs(folderRows.length - 3716) > 250) {
    issues.push(
      issue("folder_count_deviation", "文件夹属性行数与参考样例差异较大，请检查 PDF 解析台账", {
        source: "generator",
        value: `${folderRows.length} / reference 3716`
      })
    );
  }
  if (Math.abs(reviewRows.length - 3938) > 80) {
    issues.push(
      issue("review_count_deviation", "设计文档审阅流程行数与参考样例差异较大，请检查矩阵 sheet 或表头行", {
        source: "generator",
        value: `${reviewRows.length} / reference 3938`
      })
    );
  }

  return {
    matrix,
    project: {
      ...project,
      units: [...project.units].sort(compareByCodeName),
      devices: [...project.devices].sort(compareByCodeName)
    },
    folderRows,
    reviewRows,
    issues,
    stats: {
      unitCount: project.units.length,
      deviceCount: project.devices.length,
      disciplineCount: matrix.disciplines.length,
      folderRowCount: folderRows.length,
      reviewRowCount: reviewRows.length,
      issueCount: issues.length
    }
  };
}
