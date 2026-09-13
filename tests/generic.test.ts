import ExcelJS from "exceljs";
import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { combineGeneric, parseGenericSources, writeGenericWorkbook } from "../server/src/core/genericCombiner.js";

async function workbookBuffer(sheetName: string, rows: unknown[][]): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName);
  rows.forEach((row) => sheet.addRow(row));
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function richTextValue(...texts: string[]) {
  return { richText: texts.map((text) => ({ text })) };
}

function setRowValues(sheet: ExcelJS.Worksheet, rowNo: number, values: unknown[]): void {
  values.forEach((value, index) => {
    sheet.getCell(rowNo, index + 1).value = value as any;
  });
}

describe("generic combiner", () => {
  it("left joins two excel sources and exports a workbook", async () => {
    const files = [
      {
        fileName: "主表.xlsx",
        buffer: await workbookBuffer("人员", [
          ["姓名", "部门"],
          ["张三", "工程部"],
          ["李四", "采购部"]
        ])
      },
      {
        fileName: "电话.xlsx",
        buffer: await workbookBuffer("通讯录", [
          ["姓名", "电话"],
          ["张三", "1001"]
        ])
      }
    ];

    const result = await combineGeneric(files, {
      baseSourceId: "excel:1:主表.xlsx:人员",
      joins: [{
        sourceId: "excel:2:电话.xlsx:通讯录",
        leftKey: "姓名",
        rightKey: "姓名",
        joinType: "left",
        prefix: "通讯录"
      }],
      outputColumns: [
        { label: "姓名", sourceId: "excel:1:主表.xlsx:人员", column: "姓名" },
        { label: "部门", sourceId: "excel:1:主表.xlsx:人员", column: "部门" },
        { label: "电话", sourceId: "excel:2:电话.xlsx:通讯录", column: "电话" }
      ]
    });

    expect(result.stats.outputRowCount).toBe(2);
    expect(result.outputPreview[0]).toEqual({ 姓名: "张三", 部门: "工程部", 电话: "1001" });
    expect(result.outputPreview[1]).toEqual({ 姓名: "李四", 部门: "采购部", 电话: "" });

    const output = await writeGenericWorkbook(files, {
      baseSourceId: "excel:1:主表.xlsx:人员",
      joins: [{ sourceId: "excel:2:电话.xlsx:通讯录", leftKey: "姓名", rightKey: "姓名", joinType: "left" }]
    });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(output as any);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["组合结果", "问题清单", "数据源概览"]);
  });

  it("preserves generic text values and exports a problem workbook for invalid config", async () => {
    const files = [{
      fileName: "文本.xlsx",
      buffer: await workbookBuffer("明细", [
        ["名称", "备注"],
        ["A", "一；二、三"]
      ])
    }];

    const result = await combineGeneric(files, {
      baseSourceId: "excel:1:文本.xlsx:明细",
      outputColumns: [{ label: "备注", sourceId: "excel:1:文本.xlsx:明细", column: "备注" }]
    });
    expect(result.outputPreview[0]).toEqual({ 备注: "一；二、三" });

    const output = await writeGenericWorkbook(files, { baseSourceId: "missing" });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(output as any);
    expect(workbook.getWorksheet("问题清单")?.getCell("A2").value).toBe("base_source_missing");
  });

  it("handles one-to-many joins by expanding rows or merging values", async () => {
    const files = [
      {
        fileName: "订单.xlsx",
        buffer: await workbookBuffer("订单", [
          ["订单号", "客户"],
          ["O-001", "张三"],
          ["O-002", "李四"]
        ])
      },
      {
        fileName: "明细.xlsx",
        buffer: await workbookBuffer("明细", [
          ["订单号", "行号", "商品"],
          ["O-001", "1", "阀门"],
          ["O-001", "2", "管件"],
          ["O-002", "1", "电缆"]
        ])
      }
    ];

    const expandResult = await combineGeneric(files, {
      baseSourceId: "excel:1:订单.xlsx:订单",
      joins: [{
        sourceId: "excel:2:明细.xlsx:明细",
        leftKey: "订单号",
        rightKey: "订单号",
        joinType: "left",
        duplicateStrategy: "expand"
      }],
      outputColumns: [
        { label: "订单号", sourceId: "excel:1:订单.xlsx:订单", column: "订单号" },
        { label: "客户", sourceId: "excel:1:订单.xlsx:订单", column: "客户" },
        { label: "行号", sourceId: "excel:2:明细.xlsx:明细", column: "行号" },
        { label: "商品", sourceId: "excel:2:明细.xlsx:明细", column: "商品" }
      ]
    });

    expect(expandResult.stats.outputRowCount).toBe(3);
    expect(expandResult.outputPreview.map((row) => row.商品)).toEqual(["阀门", "管件", "电缆"]);
    expect(expandResult.issues.find((item) => item.type === "join_duplicate_key")?.message).toContain("展开为多行");

    const mergeResult = await combineGeneric(files, {
      baseSourceId: "excel:1:订单.xlsx:订单",
      joins: [{
        sourceId: "excel:2:明细.xlsx:明细",
        leftKey: "订单号",
        rightKey: "订单号",
        joinType: "left",
        duplicateStrategy: "merge",
        multiValueSeparator: "、"
      }],
      outputColumns: [
        { label: "订单号", sourceId: "excel:1:订单.xlsx:订单", column: "订单号" },
        { label: "商品", sourceId: "excel:2:明细.xlsx:明细", column: "商品" }
      ]
    });

    expect(mergeResult.stats.outputRowCount).toBe(2);
    expect(mergeResult.outputPreview[0]).toEqual({ 订单号: "O-001", 商品: "阀门、管件" });
    expect(mergeResult.outputPreview[1]).toEqual({ 订单号: "O-002", 商品: "电缆" });
    expect(mergeResult.issues.find((item) => item.type === "join_duplicate_key")?.message).toContain("合并到同一行");
  });

  it("renders template output columns from row fields and custom literals", async () => {
    const files = [
      {
        fileName: "主项.xlsx",
        buffer: await workbookBuffer("主项", [
          ["代码", "名称"],
          ["10", "全厂系统"],
          ["20", "工艺装置"]
        ])
      },
      {
        fileName: "分工.xlsx",
        buffer: await workbookBuffer("分工", [
          ["代码", "设计单位"],
          ["10", "华陆公司"],
          ["20", "赛鼎公司"]
        ])
      }
    ];

    const result = await combineGeneric(files, {
      baseSourceId: "excel:1:主项.xlsx:主项",
      joins: [{
        sourceId: "excel:2:分工.xlsx:分工",
        leftKey: "代码",
        rightKey: "代码",
        joinType: "left"
      }],
      outputColumns: [
        { label: "功能区", template: "{代码}/{名称}" },
        { label: "带表名前缀", template: "{分工.设计单位}-{主项.名称}" }
      ]
    });

    expect(result.outputPreview).toEqual([
      { 功能区: "10/全厂系统", 带表名前缀: "华陆公司-全厂系统" },
      { 功能区: "20/工艺装置", 带表名前缀: "赛鼎公司-工艺装置" }
    ]);
  });

  it("ignores formatted empty tail columns when parsing excel sources", async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("资产分类");
    sheet.addRow(["资产类别", "资产类别名称"]);
    sheet.addRow(["C01", "炉类"]);
    sheet.getCell("A2").dataValidation = { type: "list", allowBlank: true, formulae: ['"C01,C02"'] };
    sheet.getCell("XEZ1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };

    const sources = await parseGenericSources([{
      fileName: "资产分类.xlsx",
      buffer: Buffer.from(await workbook.xlsx.writeBuffer())
    }]);

    expect(sources[0].source.columnCount).toBe(2);
    expect(sources[0].source.headers).toEqual(["资产类别", "资产类别名称"]);
    expect(sources[0].rows[0]).toEqual({ 资产类别: "C01", 资产类别名称: "炉类" });
  });

  it("preserves sparse-offset columns when detecting excel headers", async () => {
    const directoryWorkbook = new ExcelJS.Workbook();
    const directorySheet = directoryWorkbook.addWorksheet("表4-3资产属性表目录");
    directorySheet.getCell("B1").value = "表4-3 资产属性表目录";
    directorySheet.getCell("C1").value = "表4-3 资产属性表目录";
    directorySheet.getCell("B2").value = "资产类别";
    directorySheet.getCell("C2").value = "对应的属性Sheet";
    directorySheet.getCell("B3").value = "C0101-裂解炉";
    directorySheet.getCell("C3").value = "EQ-01_容器类";
    directorySheet.getCell("B4").value = "C0104-加热炉";
    directorySheet.getCell("C4").value = "EQ-06_工业炉类";
    directorySheet.getCell("XEZ1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };

    const attributeWorkbook = new ExcelJS.Workbook();
    const attributeSheet = attributeWorkbook.addWorksheet("表4-4资产属性和计量单位目录");
    attributeSheet.addRow(["资产属性分类名", "属性ID", "属性中文名"]);
    attributeSheet.addRow(["EQ-01_容器类", "ZJYJ_Tag Number", "位号"]);
    attributeSheet.addRow(["EQ-01_容器类", "ZJYJ_Drawing Number", "图号"]);
    attributeSheet.addRow(["EQ-06_工业炉类", "ZJYJ_Equipment Type", "设备型式"]);

    const files = [
      { fileName: "附件1.xlsx", buffer: Buffer.from(await directoryWorkbook.xlsx.writeBuffer()) },
      { fileName: "附件2.xlsx", buffer: Buffer.from(await attributeWorkbook.xlsx.writeBuffer()) }
    ];
    const directorySourceId = "excel:1:附件1.xlsx:表4-3资产属性表目录";
    const attributeSourceId = "excel:2:附件2.xlsx:表4-4资产属性和计量单位目录";

    const sources = await parseGenericSources(files);
    const directory = sources.find((source) => source.source.id === directorySourceId);
    expect(directory?.source.defaultHeaderRow).toBe(2);
    expect(directory?.source.columnCount).toBe(3);
    expect(directory?.source.headers).toEqual(["列A", "资产类别", "对应的属性Sheet"]);
    expect(directory?.rows[0]).toEqual({
      列A: "",
      资产类别: "C0101-裂解炉",
      对应的属性Sheet: "EQ-01_容器类"
    });

    const result = await combineGeneric(files, {
      baseSourceId: directorySourceId,
      joins: [{
        sourceId: attributeSourceId,
        leftKey: "对应的属性Sheet",
        rightKey: "资产属性分类名",
        joinType: "left",
        duplicateStrategy: "expand"
      }],
      outputColumns: [
        { label: "资产类别", sourceId: directorySourceId, column: "资产类别" },
        { label: "对应的属性Sheet", sourceId: directorySourceId, column: "对应的属性Sheet" },
        { label: "属性ID", sourceId: attributeSourceId, column: "属性ID" },
        { label: "属性中文名", sourceId: attributeSourceId, column: "属性中文名" }
      ]
    });

    expect(result.stats.outputRowCount).toBe(3);
    expect(result.outputPreview).toEqual([
      { 资产类别: "C0101-裂解炉", 对应的属性Sheet: "EQ-01_容器类", 属性ID: "ZJYJ_Tag Number", 属性中文名: "位号" },
      { 资产类别: "C0101-裂解炉", 对应的属性Sheet: "EQ-01_容器类", 属性ID: "ZJYJ_Drawing Number", 属性中文名: "图号" },
      { 资产类别: "C0104-加热炉", 对应的属性Sheet: "EQ-06_工业炉类", 属性ID: "ZJYJ_Equipment Type", 属性中文名: "设备型式" }
    ]);
  });

  it("parses rich-text engineering sheets with grouped two-row headers", async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    sheet.mergeCells("A1:B1");
    sheet.mergeCells("C1:L1");
    sheet.mergeCells("M1:P1");
    sheet.getCell("A1").value = richTextValue("功能区", "类别") as any;
    sheet.getCell("C1").value = richTextValue("单元", "类别") as any;
    sheet.getCell("M1").value = richTextValue("装置主项", "类别") as any;
    sheet.getCell("Q1").value = richTextValue("详细设计", "设计分工") as any;
    sheet.getCell("R1").value = "备注";

    setRowValues(sheet, 2, [
      richTextValue("装置\n", "代码"),
      "名称",
      "序号",
      "序号",
      "名称",
      "名称",
      "名称",
      "名称",
      "主项号",
      "主项号",
      "单元规模",
      "单元规模",
      "序号",
      "名称",
      "子主项号",
      "子主项号",
      "详细设计设计分工",
      "备注"
    ]);
    setRowValues(sheet, 3, [
      "10",
      richTextValue("全厂\n", "系统"),
      "1",
      "1",
      richTextValue("全厂", "总图运输"),
      "全厂总图运输",
      "全厂总图运输",
      "全厂总图运输",
      "10082",
      "10082",
      "",
      "",
      "1",
      "全厂总图运输",
      "10082",
      "10082",
      richTextValue("华陆", "公司"),
      "含征地红线内的总图运输、道路、竖向、围墙大门"
    ]);
    setRowValues(sheet, 4, [
      "10",
      "全厂系统",
      "1",
      "1",
      "全厂总图运输",
      "全厂总图运输",
      "全厂总图运输",
      "全厂总图运输",
      "10082",
      "10082",
      "",
      "",
      "2",
      "装置总图A",
      "10082A",
      "10082A",
      "赛鼎公司",
      "包含装置区道路、竖向、围墙和相关设施分工"
    ]);

    const sources = await parseGenericSources([{
      fileName: "项目主项表及设计分工.xlsx",
      buffer: Buffer.from(await workbook.xlsx.writeBuffer())
    }]);
    const table = sources[0];

    expect(table.source.defaultHeaderRow).toBe(2);
    expect(table.source.defaultDataStartRow).toBe(3);
    expect(table.source.headers.some((header) => header.includes("[object Object]"))).toBe(false);
    expect(table.source.headers.slice(0, 4)).toEqual([
      "功能区类别.装置代码",
      "功能区类别.名称",
      "单元类别.序号",
      "单元类别.序号_2"
    ]);
    expect(table.source.headers).toContain("装置主项类别.子主项号");
    expect(table.source.suggestedFillDownColumns).toContain("功能区类别.装置代码");
    expect(table.rows[0]["功能区类别.装置代码"]).toBe("10");
    expect(table.rows[0]["功能区类别.名称"]).toBe("全厂系统");
    expect(table.rows[0]["装置主项类别.子主项号"]).toBe("10082");
  });

  it("exports selected source tables as separate result sheets", async () => {
    const files = [
      {
        fileName: "项目.xlsx",
        buffer: await workbookBuffer("项目", [
          ["项目编号", "项目名称"],
          ["P-001", "管道改造"]
        ])
      },
      {
        fileName: "人员.xlsx",
        buffer: await workbookBuffer("人员", [
          ["姓名", "部门"],
          ["张三", "工程部"]
        ])
      }
    ];

    const output = await writeGenericWorkbook(files, {
      resultSheets: [
        { sourceId: "excel:1:项目.xlsx:项目", sheetName: "项目清单" },
        { sourceId: "excel:2:人员.xlsx:人员", sheetName: "人员清单" }
      ]
    });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(output as any);

    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["项目清单", "人员清单", "问题清单", "数据源概览"]);
    expect(workbook.getWorksheet("项目清单")?.getRow(2).values).toEqual([, "P-001", "管道改造"]);
    expect(workbook.getWorksheet("人员清单")?.getRow(2).values).toEqual([, "张三", "工程部"]);
  });

  it("defaults to business sheets and supports data-start and fill-down settings", async () => {
    const workbook = new ExcelJS.Workbook();
    const instruction = workbook.addWorksheet("填写说明");
    instruction.addRow(["序号", "提示内容"]);
    instruction.addRow(["1", "请勿删除说明页"]);

    const main = workbook.addWorksheet("主表");
    main.addRow(["导入模板"]);
    main.addRow(["", "", ""]);
    main.addRow(["装置", "专业", "负责人"]);
    main.addRow(["说明", "说明", "说明"]);
    main.addRow(["全厂总图运输", "总图运输设计类（GPE）", "张三"]);
    main.addRow(["", "管道设计类（PIE）", "李四"]);

    const dict = workbook.addWorksheet("hidden#ho_zzmc5");
    dict.state = "hidden";
    dict.addRow(["全厂总图运输"]);
    dict.addRow(["空分装置"]);

    const files = [{ fileName: "导入模板.xlsx", buffer: Buffer.from(await workbook.xlsx.writeBuffer()) }];
    const sources = await parseGenericSources(files);

    expect(sources.map((source) => source.source.sheetName)).toEqual(["主表"]);
    expect(sources[0].source.defaultHeaderRow).toBe(3);
    expect(sources[0].source.defaultDataStartRow).toBe(4);

    const result = await combineGeneric(files, {
      baseSourceId: "excel:1:导入模板.xlsx:主表",
      sourceConfigs: [{
        sourceId: "excel:1:导入模板.xlsx:主表",
        headerRow: 3,
        dataStartRow: 5,
        fillDownColumns: ["装置"]
      }],
      outputColumns: [
        { label: "装置", sourceId: "excel:1:导入模板.xlsx:主表", column: "装置" },
        { label: "专业", sourceId: "excel:1:导入模板.xlsx:主表", column: "专业" }
      ]
    });

    expect(result.outputPreview).toEqual([
      { 装置: "全厂总图运输", 专业: "总图运输设计类（GPE）" },
      { 装置: "全厂总图运输", 专业: "管道设计类（PIE）" }
    ]);

    const withAuxiliary = await parseGenericSources(files, { includeAuxiliarySheets: true });
    expect(withAuxiliary.map((source) => [source.source.sheetName, source.source.role])).toEqual([
      ["填写说明", "instruction"],
      ["主表", "data"],
      ["hidden#ho_zzmc5", "dictionary"]
    ]);
  });

  it("can fill a template workbook while preserving instruction and hidden sheets", async () => {
    const source = await workbookBuffer("源数据", [
      ["序号", "设计文档", "装置", "专业分类"],
      ["sjwj000001", "文档文件", "10082全厂总图运输", "外管设计类（OPE）"],
      ["sjwj000002", "文档文件", "10082全厂总图运输", "热工设计类（TPE）"]
    ]);

    const template = new ExcelJS.Workbook();
    const instruction = template.addWorksheet("填写说明");
    instruction.addRow(["序号", "提示内容"]);
    instruction.addRow(["1", "请勿删除说明页"]);

    const main = template.addWorksheet("主表");
    main.columns = [{ width: 24 }, { width: 24 }, { width: 24 }, { width: 24 }];
    main.addRow(["{.sys_record_no}", "{.ho_Design_Document}", "{.ho_zzmc}", "{.ho_discipline_sjname}"]);
    main.addRow([]);
    main.addRow([]);
    main.addRow(["序号", "设计文档", "装置", "专业分类"]);
    main.addRow(["old", "old", "old", "old"]);
    main.getRow(5).getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE8CC" } };

    const dict = template.addWorksheet("hidden#ho_zzmc5");
    dict.state = "hidden";
    dict.addRow(["10082全厂总图运输"]);

    const files = [
      { fileName: "源数据.xlsx", buffer: source },
      { fileName: "设计文档审阅流程 导入.xlsx", buffer: Buffer.from(await template.xlsx.writeBuffer()) }
    ];

    const output = await writeGenericWorkbook(files, {
      baseSourceId: "excel:1:源数据.xlsx:源数据",
      templateExport: {
        sourceId: "excel:2:设计文档审阅流程 导入.xlsx:主表",
        headerRow: 4,
        dataStartRow: 5
      },
      outputColumns: [
        { label: "序号", sourceId: "excel:1:源数据.xlsx:源数据", column: "序号" },
        { label: "设计文档", sourceId: "excel:1:源数据.xlsx:源数据", column: "设计文档" },
        { label: "装置", sourceId: "excel:1:源数据.xlsx:源数据", column: "装置" },
        { label: "专业分类", sourceId: "excel:1:源数据.xlsx:源数据", column: "专业分类" }
      ]
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(output as any);

    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["填写说明", "主表", "hidden#ho_zzmc5"]);
    expect(workbook.getWorksheet("hidden#ho_zzmc5")?.state).toBe("hidden");
    expect(workbook.getWorksheet("主表")?.getRow(1).values).toEqual([, "{.sys_record_no}", "{.ho_Design_Document}", "{.ho_zzmc}", "{.ho_discipline_sjname}"]);
    expect(workbook.getWorksheet("主表")?.getRow(5).values).toEqual([, "sjwj000001", "文档文件", "10082全厂总图运输", "外管设计类（OPE）"]);
    expect(workbook.getWorksheet("主表")?.getRow(6).values).toEqual([, "sjwj000002", "文档文件", "10082全厂总图运输", "热工设计类（TPE）"]);
    expect(workbook.getWorksheet("主表")?.getColumn(1).width).toBe(24);
    expect(workbook.getWorksheet("组合结果")).toBeUndefined();
  });
});
