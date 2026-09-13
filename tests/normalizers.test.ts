import { describe, expect, it } from "vitest";
import { cleanText, normalizeDeviceName, normalizeDisciplineName, normalizePeople, normalizeCode } from "../server/src/core/normalizers.js";

describe("normalizers", () => {
  it("cleans device codes and spacing", () => {
    expect(normalizeDeviceName("10082A装置总图运输 A")).toBe("装置总图运输A");
    expect(normalizeDeviceName("22804-22806电子化学品装置")).toBe("电子化学品装置");
  });

  it("normalizes discipline parentheses", () => {
    expect(normalizeDisciplineName("工艺设计类(PRE)")).toBe("工艺设计类（PRE）");
  });

  it("uses English commas for people fields", () => {
    expect(normalizePeople("杨波；白科、李亚龙")).toBe("杨波,白科,李亚龙");
  });

  it("expands short range suffixes", () => {
    expect(normalizeCode("22804~2806")).toBe("22804-22806");
  });

  it("cleans ExcelJS rich text and formula result values", () => {
    expect(cleanText({ richText: [{ text: "装置\n" }, { text: "代码" }] })).toBe("装置代码");
    expect(cleanText({ formula: "A1+B1", result: " 10082 " })).toBe("10082");
    expect(cleanText({ text: "项目主页", hyperlink: "https://example.com" })).toBe("项目主页");
  });
});
