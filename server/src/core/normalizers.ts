const FULL_WIDTH_MAP: Record<string, string> = {
  "(": "（",
  ")": "）",
  "，": ",",
  ";": ",",
  "；": ",",
  "、": ","
};

export const AREA_NAMES: Record<string, string> = {
  "10": "全厂系统",
  "20": "工艺装置",
  "30": "公用工程",
  "40": "辅助生产设施",
  "50": "物流设施",
  "60": "基础设施"
};

export const DESIGN_COMPANIES = ["华陆公司", "赛鼎公司", "航天科技"];

function normalizeWhitespace(value: string): string {
  return value
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/([\u4e00-\u9fff])\s+([\u4e00-\u9fff])/g, "$1$2");
}

export function cleanText(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "object") {
    const record = value as Record<string, unknown>;
    if (Array.isArray(record.richText)) {
      return cleanText(record.richText.map((part) => {
        if (part && typeof part === "object" && "text" in (part as Record<string, unknown>)) {
          return (part as { text?: unknown }).text ?? "";
        }
        return part ?? "";
      }).join(""));
    }
    if ("result" in record) {
      return cleanText(record.result);
    }
    if ("text" in record) {
      return cleanText(record.text);
    }
    if ("error" in record) {
      return cleanText(record.error);
    }
  }
  return normalizeWhitespace(String(value));
}

export function normalizeParentheses(value: string): string {
  return value.replace(/[()]/g, (char) => FULL_WIDTH_MAP[char] ?? char);
}

export function compactText(value: string): string {
  return normalizeParentheses(cleanText(value)).replace(/\s+/g, "");
}

export function normalizePeople(value: unknown): string {
  return cleanText(value)
    .replace(/[；;，、]/g, ",")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean)
    .join(",");
}

export function normalizeDisciplineName(value: unknown): string {
  return normalizeParentheses(cleanText(value))
    .replace(/\s+/g, "")
    .replace(/（([A-Z]{3})）$/, "（$1）");
}

export function normalizeDeviceName(value: unknown): string {
  let text = compactText(cleanText(value));
  text = text.replace(/^(\d{5})([A-Z])?[-~～](\d{5})([A-Z])?/, "");
  text = text.replace(/^\d{5}[A-Z]?/, "");
  text = text.replace(/全场消防/g, "全厂消防");
  return text;
}

export function normalizeJoinKey(deviceName: unknown, disciplineName?: unknown): string {
  const device = normalizeDeviceName(deviceName);
  const discipline = disciplineName == null ? "" : `|${normalizeDisciplineName(disciplineName)}`;
  return `${device}${discipline}`;
}

export function codeArea(code: string): { areaCode: string; areaName: string } | undefined {
  const normalized = code.replace(/[A-Z].*$/, "");
  const prefix = normalized.slice(0, 2);
  const directAreaName = AREA_NAMES[prefix];
  if (directAreaName) return { areaCode: prefix, areaName: directAreaName };
  const decadePrefix = `${normalized.slice(0, 1)}0`;
  const decadeAreaName = AREA_NAMES[decadePrefix];
  return decadeAreaName ? { areaCode: decadePrefix, areaName: decadeAreaName } : undefined;
}

export function displayArea(areaCode: string, areaName: string): string {
  return `${areaCode}${areaName}`;
}

export function displayCodeName(code: string, name: string): string {
  return `${code}${compactText(name)}`;
}

export function disciplineCode(value: string): string {
  return normalizeDisciplineName(value).match(/（([A-Z]{3})）/)?.[1] ?? "";
}

export function normalizeCode(value: string): string {
  return cleanText(value)
    .replace(/[～~]/g, "-")
    .replace(/\s+/g, "")
    .replace(/^(\d{5})-(\d{1,5})$/, (_, start: string, end: string) => {
      const prefix = start.slice(0, Math.max(0, 5 - end.length));
      return `${start}-${prefix}${end}`;
    });
}

export function compareByCodeName<T extends { code: string; name: string }>(a: T, b: T): number {
  return a.code.localeCompare(b.code, "zh-Hans-CN", { numeric: true }) || a.name.localeCompare(b.name, "zh-Hans-CN");
}
