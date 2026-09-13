export type IssueSeverity = "warning" | "error";

export interface IssueRow {
  type: string;
  severity: IssueSeverity;
  message: string;
  source?: string;
  row?: number;
  key?: string;
  value?: string;
}

export interface MatrixRow {
  sourceRow: number;
  category: string;
  deviceName: string;
  disciplineName: string;
  digitalPmc: string;
  specialtyEngineer: string;
  professionalLeader: string;
  archiveOwner?: string;
}

export interface ParsedMatrix {
  sheetName: string;
  rows: MatrixRow[];
  disciplines: string[];
  devices: string[];
  issues: IssueRow[];
}

export interface UnitNode {
  code: string;
  name: string;
  areaCode: string;
  areaName: string;
  designCompany: string;
  sourcePage: number;
  sourceY: number;
  rawLine: string;
}

export interface DeviceNode {
  code: string;
  name: string;
  areaCode: string;
  areaName: string;
  unitCode: string;
  unitName: string;
  designCompany: string;
  sourcePage: number;
  sourceY: number;
  rawLine: string;
  separateFolder: boolean;
}

export interface ParsedProject {
  units: UnitNode[];
  devices: DeviceNode[];
  issues: IssueRow[];
  previewLines: string[];
}

export interface FolderAttributeRow {
  folderPath: string;
  name: string;
  updateName: string;
  code: string;
  createDate: string;
  updateDate: string;
  createBy: string;
  updateBy: string;
  belongingUnit: string;
  levelName: string;
  device: string;
  unit: string;
  designCompany: string;
  plotConfiguration: string;
}

export interface ReviewFlowRow {
  recordNo: string;
  creator: string;
  createdDate: string;
  recordStatus: string;
  designDocument: string;
  functionAreaName: string;
  unitName: string;
  deviceName: string;
  disciplineName: string;
  digitalPmc: string;
  specialtyEngineer: string;
  professionalLeader: string;
}

export interface GenerateConfig {
  matrixSheetName?: string;
  matrixHeaderRow?: number;
  matrixDataStartRow?: number;
  rootPath?: string;
  designDocumentName?: string;
  recordPrefix?: string;
  folderCreator?: string;
  previewLimit?: number;
}

export interface GenerationResult {
  matrix: ParsedMatrix;
  project: ParsedProject;
  folderRows: FolderAttributeRow[];
  reviewRows: ReviewFlowRow[];
  issues: IssueRow[];
  stats: {
    unitCount: number;
    deviceCount: number;
    disciplineCount: number;
    folderRowCount: number;
    reviewRowCount: number;
    issueCount: number;
  };
}

export interface PreviewResult {
  projectPreview: DeviceNode[];
  folderPreview: FolderAttributeRow[];
  reviewPreview: ReviewFlowRow[];
  issues: IssueRow[];
  stats: GenerationResult["stats"];
}
