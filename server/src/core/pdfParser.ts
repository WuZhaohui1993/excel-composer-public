import type { DeviceNode, IssueRow, ParsedProject, UnitNode } from "./types.js";
import {
  AREA_NAMES,
  DESIGN_COMPANIES,
  cleanText,
  codeArea,
  compactText,
  compareByCodeName,
  displayCodeName,
  normalizeCode
} from "./normalizers.js";

interface TextItem {
  str: string;
  x: number;
  y: number;
}

interface PdfLine {
  page: number;
  y: number;
  items: TextItem[];
  text: string;
}

function field(line: PdfLine, minX: number, maxX: number): string {
  return line.items
    .filter((item) => item.x >= minX && item.x < maxX)
    .sort((a, b) => a.x - b.x)
    .map((item) => item.str)
    .join("");
}

function companyFromLine(line: PdfLine): string {
  const text = compactText(field(line, 780, 855) || line.text);
  return DESIGN_COMPANIES.find((company) => text.includes(company)) ?? "";
}

function extractCode(text: string): string {
  const match = cleanText(text).match(/\d{5}[A-Z]?|\d{5}\s*[~～-]\s*\d{1,5}/);
  return match ? normalizeCode(match[0]) : "";
}

function normalizeName(name: string): string {
  return compactText(name)
    .replace(/^[0-9]+$/, "")
    .replace(/^[/、，,]+/, "")
    .replace(/全场消防/g, "全厂消防");
}

function isHeaderOrRemark(line: PdfLine): boolean {
  return /文档编号|文件名称|功能区类别|单元类别|装置主项类别|设计分工|备注|子主项号|主项号|代码|名称|版次|注：|具体|主项划分|技术商后确定/.test(line.text);
}

function addOrUpdateUnit(units: Map<string, UnitNode>, unit: UnitNode): UnitNode {
  const key = `${unit.code}|${unit.name}`;
  const current = units.get(key);
  if (!current) {
    units.set(key, unit);
    return unit;
  }
  if (!current.designCompany && unit.designCompany) current.designCompany = unit.designCompany;
  return current;
}

function addOrUpdateDevice(devices: Map<string, DeviceNode>, device: DeviceNode): void {
  const key = `${device.code}|${device.name}|${device.unitCode}`;
  const current = devices.get(key);
  if (!current) {
    devices.set(key, device);
    return;
  }
  if (!current.designCompany && device.designCompany) current.designCompany = device.designCompany;
}

function isUnitLikeCode(code: string): boolean {
  return /^\d{5}$/.test(code) || /^\d{5}-\d{5}$/.test(code);
}

async function extractLines(buffer: Buffer): Promise<PdfLine[]> {
  const pdfjs = (await import("pdfjs-dist/legacy/build/pdf.mjs")) as any;
  const loadingTask = pdfjs.getDocument({
    data: new Uint8Array(buffer),
    disableFontFace: true,
    useSystemFonts: true
  });
  const doc = await loadingTask.promise;
  const lines: PdfLine[] = [];

  for (let pageNo = 1; pageNo <= doc.numPages; pageNo += 1) {
    const page = await doc.getPage(pageNo);
    const content = await page.getTextContent({ includeMarkedContent: false });
    const rawItems: TextItem[] = content.items
      .filter((item: any) => typeof item.str === "string" && item.str.trim())
      .map((item: any) => ({
        str: cleanText(item.str),
        x: Number(item.transform[4]),
        y: Number(item.transform[5])
      }));
    const groups: Array<{ y: number; items: TextItem[] }> = [];
    for (const item of rawItems.sort((a, b) => b.y - a.y || a.x - b.x)) {
      let group = groups.find((candidate) => Math.abs(candidate.y - item.y) < 2);
      if (!group) {
        group = { y: item.y, items: [] };
        groups.push(group);
      }
      group.items.push(item);
    }
    for (const group of groups) {
      const items = group.items.sort((a, b) => a.x - b.x);
      lines.push({
        page: pageNo,
        y: group.y,
        items,
        text: items.map((item) => item.str).join(" ")
      });
    }
  }

  return lines;
}

export async function parseProjectPdf(buffer: Buffer): Promise<ParsedProject> {
  const lines = await extractLines(buffer);
  const issues: IssueRow[] = [];
  const units = new Map<string, UnitNode>();
  const devices = new Map<string, DeviceNode>();
  let currentUnit: UnitNode | undefined;

  for (const line of lines) {
    if (line.page === 1 || isHeaderOrRemark(line)) continue;

    const unitNameRaw = normalizeName(field(line, 225, 370));
    const unitCodeRaw = extractCode(field(line, 370, 430));
    const childNameRaw = normalizeName(field(line, 560, 705));
    const childCodeRaw = extractCode(field(line, 705, 780));
    const company = companyFromLine(line);

    let unitName = unitNameRaw;
    let unitCode = unitCodeRaw;

    if (!unitCode && /电子化学品装置/.test(compactText(line.text))) {
      unitName = "电子化学品装置";
      unitCode = "22804-22806";
    }

    if (unitCode && unitName) {
      const area = codeArea(unitCode);
      if (!area) {
        issues.push({
          type: "pdf_unknown_area",
          severity: "warning",
          message: "PDF 主项号无法识别功能区",
          source: `PDF 第 ${line.page} 页`,
          key: unitCode,
          value: line.text
        });
      } else {
        currentUnit = addOrUpdateUnit(units, {
          code: unitCode,
          name: unitName,
          areaCode: area.areaCode,
          areaName: area.areaName,
          designCompany: childCodeRaw ? "" : company,
          sourcePage: line.page,
          sourceY: line.y,
          rawLine: line.text
        });
      }
    }

    if (childCodeRaw && childNameRaw && isUnitLikeCode(childCodeRaw) && (!unitCode || unitCode === childCodeRaw)) {
      const area = codeArea(childCodeRaw);
      if (area) {
        currentUnit = addOrUpdateUnit(units, {
          code: childCodeRaw,
          name: childNameRaw,
          areaCode: area.areaCode,
          areaName: area.areaName,
          designCompany: company,
          sourcePage: line.page,
          sourceY: line.y,
          rawLine: line.text
        });
        addOrUpdateDevice(devices, {
          code: childCodeRaw,
          name: childNameRaw,
          areaCode: area.areaCode,
          areaName: area.areaName,
          unitCode: childCodeRaw,
          unitName: childNameRaw,
          designCompany: company,
          sourcePage: line.page,
          sourceY: line.y,
          rawLine: line.text,
          separateFolder: false
        });
      }
    } else if (childCodeRaw && childNameRaw && currentUnit) {
      const area = codeArea(childCodeRaw) ?? { areaCode: currentUnit.areaCode, areaName: currentUnit.areaName };
      addOrUpdateDevice(devices, {
        code: childCodeRaw,
        name: childNameRaw,
        areaCode: area.areaCode,
        areaName: area.areaName,
        unitCode: currentUnit.code,
        unitName: currentUnit.name,
        designCompany: company,
        sourcePage: line.page,
        sourceY: line.y,
        rawLine: line.text,
        separateFolder: childCodeRaw !== currentUnit.code || normalizeName(childNameRaw) !== normalizeName(currentUnit.name)
      });
    } else if (unitCode && unitName && company && currentUnit) {
      addOrUpdateDevice(devices, {
        code: unitCode,
        name: unitName,
        areaCode: currentUnit.areaCode,
        areaName: currentUnit.areaName,
        unitCode: currentUnit.code,
        unitName: currentUnit.name,
        designCompany: company,
        sourcePage: line.page,
        sourceY: line.y,
        rawLine: line.text,
        separateFolder: false
      });
    }
  }

  for (const [code, areaName] of Object.entries(AREA_NAMES)) {
    if (![...units.values()].some((unit) => unit.areaCode === code)) {
      issues.push({
        type: "pdf_area_empty",
        severity: "warning",
        message: `PDF 未解析到 ${code}${areaName} 下的主项`,
        source: "PDF"
      });
    }
  }

  const unitList = [...units.values()].sort(compareByCodeName);
  const deviceList = [...devices.values()].sort(compareByCodeName);
  const previewLines = lines
    .filter((line) => line.page > 1 && !isHeaderOrRemark(line))
    .slice(0, 120)
    .map((line) => `P${line.page} ${Math.round(line.y)} ${line.text}`);

  return {
    units: unitList,
    devices: deviceList,
    issues,
    previewLines
  };
}
