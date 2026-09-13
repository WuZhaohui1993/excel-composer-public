import React from "react";
import ReactDOM from "react-dom/client";
import { AnimatePresence, motion } from "framer-motion";
import { AlertTriangle, ArrowDown, ArrowUp, CheckCircle2, Database, Download, Eraser, Eye, FileSpreadsheet, FileText, GripVertical, Info, Layers3, Link2, Loader2, PanelLeftClose, PanelLeftOpen, Plus, Search, Settings2, Sparkles, UploadCloud, X } from "lucide-react";
import "./styles.css";

interface IssueRow {
  type: string;
  severity: "warning" | "error";
  message: string;
  source?: string;
  row?: number;
  key?: string;
  value?: string;
}

interface DevicePreview {
  code: string;
  name: string;
  areaName: string;
  unitName: string;
  designCompany: string;
}

interface FolderPreview {
  folderPath: string;
  name: string;
  levelName: string;
  device: string;
  unit: string;
  designCompany: string;
}

interface ReviewPreview {
  recordNo: string;
  functionAreaName: string;
  unitName: string;
  deviceName: string;
  disciplineName: string;
  digitalPmc: string;
  specialtyEngineer: string;
  professionalLeader: string;
}

interface PresetPreview {
  projectPreview: DevicePreview[];
  folderPreview: FolderPreview[];
  reviewPreview: ReviewPreview[];
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

interface GenericSourceSummary {
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

interface GenericJoinRule {
  sourceId: string;
  leftKey: string;
  rightKey: string;
  joinType: "left" | "inner";
  prefix?: string;
  duplicateStrategy?: "first" | "expand" | "merge";
  multiValueSeparator?: string;
}

interface GenericJoinDraft extends GenericJoinRule {
  id: string;
}

interface GenericOutputColumn {
  label: string;
  sourceId?: string;
  column?: string;
  constant?: string;
  template?: string;
}

interface GenericOutputColumnDraft extends GenericOutputColumn {
  id: string;
}

interface GenericTemplateExportDraft {
  sourceId: string;
}

interface GenericResultSheetConfig {
  sourceId: string;
  sheetName?: string;
  outputColumns?: GenericOutputColumn[];
}

interface GenericPreview {
  sources: GenericSourceSummary[];
  outputPreview: Record<string, string>[];
  issues: IssueRow[];
  stats: {
    sourceCount: number;
    outputRowCount: number;
    issueCount: number;
  };
}

type AlertKind = "info" | "warning" | "error";

interface AlertMessage {
  kind: AlertKind;
  title: string;
  message: string;
}

interface GuideStep {
  id: string;
  title: string;
  body: string;
  target: string;
  action: string;
  done?: boolean;
  optional?: boolean;
}

interface GuideRect {
  top: number;
  left: number;
  width: number;
  height: number;
  bottom: number;
}

const presetDefaultConfig = {
  rootPath: "示例项目/文档文件",
  designDocumentName: "文档文件",
  recordPrefix: "sjwj",
  folderCreator: "",
  previewLimit: 100
};

type RightView = "fields" | "result" | "issues";
type DragPayload =
  | { kind: "source"; sourceId: string }
  | { kind: "field"; sourceId: string; column: string };
type PointerDragState = DragPayload & {
  startX: number;
  startY: number;
  x: number;
  y: number;
  active: boolean;
};

const dragMimeType = "application/x-excel-tools-drag";
const accessTokenStorageKey = "excel-tools.access-token";

function shortPath(value: string) {
  if (value.length <= 52) return value;
  return `${value.slice(0, 24)}…${value.slice(-24)}`;
}

function fileLabel(file: File | null, fallback: string) {
  return file ? file.name : fallback;
}

function Stat({ label, value }: { label: string; value: number }) {
  return (
    <div className="stat">
      <span>{label}</span>
      <strong>{value.toLocaleString("zh-CN")}</strong>
    </div>
  );
}

function AlertPopup({ alert, onClose }: { alert: AlertMessage | null; onClose: () => void }) {
  return (
    <AnimatePresence>
      {alert && (
        <motion.div
          className={`alert-popup ${alert.kind}`}
          initial={{ opacity: 0, y: -12, scale: 0.98 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          exit={{ opacity: 0, y: -12, scale: 0.98 }}
          transition={{ duration: 0.18 }}
          role="alert"
        >
          <div className="alert-popup-icon">
            {alert.kind === "error" || alert.kind === "warning" ? <AlertTriangle size={18} /> : <Info size={18} />}
          </div>
          <div className="alert-popup-content">
            <strong>{alert.title}</strong>
            <p>{alert.message}</p>
          </div>
          <button className="alert-popup-close" onClick={onClose} title="关闭提示">
            <X size={16} />
          </button>
        </motion.div>
      )}
    </AnimatePresence>
  );
}

function GuidedTour({
  steps,
  activeIndex,
  setActiveIndex,
  onClose,
  title
}: {
  steps: GuideStep[];
  activeIndex: number;
  setActiveIndex: (index: number) => void;
  onClose: () => void;
  title: string;
}) {
  const [rect, setRect] = React.useState<GuideRect | null>(null);
  const step = steps[Math.min(activeIndex, Math.max(0, steps.length - 1))];
  const completedCount = steps.filter((item) => item.done || item.optional).length;

  React.useEffect(() => {
    if (!step) return undefined;
    let frame = 0;

    function refresh() {
      const element = document.querySelector<HTMLElement>(step.target);
      if (!element) {
        setRect(null);
        return;
      }
      element.scrollIntoView({ block: "center", inline: "nearest", behavior: "smooth" });
      frame = window.setTimeout(() => {
        const next = element.getBoundingClientRect();
        setRect({
          top: next.top,
          left: next.left,
          width: next.width,
          height: next.height,
          bottom: next.bottom
        });
      }, 220);
    }

    refresh();
    window.addEventListener("resize", refresh);
    window.addEventListener("scroll", refresh, true);
    return () => {
      window.clearTimeout(frame);
      window.removeEventListener("resize", refresh);
      window.removeEventListener("scroll", refresh, true);
    };
  }, [step?.target]);

  if (!step) return null;

  const panelStyle: React.CSSProperties = rect
    ? (() => {
      const panelWidth = Math.min(360, window.innerWidth - 32);
      const panelHeight = Math.min(360, window.innerHeight - 32);
      const spaceBelow = window.innerHeight - rect.bottom;
      const spaceAbove = rect.top;
      const top = spaceBelow >= panelHeight + 16
        ? rect.bottom + 12
        : spaceAbove >= panelHeight + 16
          ? rect.top - panelHeight - 12
          : Math.max(16, Math.min(window.innerHeight - panelHeight - 16, (window.innerHeight - panelHeight) / 2));
      const left = rect.left + panelWidth <= window.innerWidth - 16
        ? rect.left
        : Math.max(16, window.innerWidth - panelWidth - 16);
      return { top, left };
    })()
    : { top: 24, left: 24 };

  return (
    <div className="guide-layer" aria-live="polite">
      <div className="guide-dim" />
      {rect && (
        <motion.div
          className="guide-highlight"
          initial={false}
          animate={{
            top: rect.top - 6,
            left: rect.left - 6,
            width: rect.width + 12,
            height: rect.height + 12
          }}
          transition={{ duration: 0.18 }}
        />
      )}
      <motion.div
        className="guide-panel"
        style={panelStyle}
        initial={{ opacity: 0, y: 8, scale: 0.98 }}
        animate={{ opacity: 1, y: 0, scale: 1 }}
        exit={{ opacity: 0, y: 8, scale: 0.98 }}
        transition={{ duration: 0.18 }}
      >
        <div className="guide-panel-head">
          <span>{title}</span>
          <button className="alert-popup-close" onClick={onClose} title="关闭教程"><X size={16} /></button>
        </div>
        <div className="guide-progress">
          <span>第 {activeIndex + 1} 步 / {steps.length}</span>
          <span>{completedCount}/{steps.length} 已完成</span>
        </div>
        <h3>{step.title}</h3>
        <p>{step.body}</p>
        <div className={`guide-status ${step.done ? "done" : step.optional ? "optional" : "todo"}`}>
          {step.done ? "已完成" : step.optional ? "可选步骤" : step.action}
        </div>
        <div className="guide-step-dots">
          {steps.map((item, index) => (
            <button
              key={item.id}
              className={`${index === activeIndex ? "active" : ""} ${item.done ? "done" : ""}`}
              onClick={() => setActiveIndex(index)}
              title={item.title}
            />
          ))}
        </div>
        <div className="guide-actions">
          <button className="secondary" disabled={activeIndex === 0} onClick={() => setActiveIndex(activeIndex - 1)}>上一步</button>
          {activeIndex === steps.length - 1 ? (
            <button className="primary" onClick={onClose}>完成</button>
          ) : (
            <button className="primary" onClick={() => setActiveIndex(activeIndex + 1)}>下一步</button>
          )}
        </div>
      </motion.div>
    </div>
  );
}

function HeaderControls({
  accessToken,
  setAccessToken,
  compact = false,
  onOpenPreset,
  onStartGuide
}: {
  accessToken: string;
  setAccessToken: (value: string) => void;
  compact?: boolean;
  onOpenPreset: () => void;
  onStartGuide?: () => void;
}) {
  const presetButton = (
    <button className="preset-entry-button" onClick={onOpenPreset} title="打开固定流程">
      <Sparkles size={15} />固定流程
    </button>
  );
  const guideButton = onStartGuide ? (
    <button className="guide-entry-button" onClick={onStartGuide} title="打开操作教程">
      <Info size={15} />教程
    </button>
  ) : null;

  if (compact) {
    return (
      <motion.header className="compact-header" initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.24 }}>
        <div className="compact-title-row">
          <div>
            <p className="eyebrow">通用模式</p>
            <h1>excel组合工具</h1>
          </div>
          <div className="header-actions">
            {presetButton}
            {guideButton}
            <details className="header-settings" data-guide="generic-token">
              <summary><Settings2 size={15} />设置</summary>
          <label className="field token-field" data-guide="generic-token">
                <span>访问口令</span>
	                <input
	                  type="password"
	                  value={accessToken}
	                  onChange={(event) => setAccessToken(event.target.value)}
	                  placeholder="输入访问口令，本机保存"
	                />
                <small className="token-save-hint">{accessToken ? "已保存在本机浏览器，清空输入框可删除。" : "输入后保存在本机浏览器。"}</small>
              </label>
            </details>
          </div>
        </div>
        <p className="subtitle compact-subtitle">
          一次生成一张目标表，复杂需求分次处理。
        </p>
      </motion.header>
    );
  }

  return (
    <>
      <motion.header initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.32 }}>
        <p className="eyebrow">通用模式</p>
        <h1>excel组合工具</h1>
        <p className="subtitle">
          上传 Excel 或 PDF，选字段、按相同字段补充数据，生成一张符合要求的结果表。
        </p>
	      </motion.header>
	      {presetButton}
	      {guideButton}
      <label className="field token-field">
        <span>访问口令</span>
	              <input
	                type="password"
	                value={accessToken}
	                onChange={(event) => setAccessToken(event.target.value)}
	                placeholder="输入访问口令，本机保存"
	              />
        <small className="token-save-hint">{accessToken ? "已保存在本机浏览器，清空输入框可删除。" : "输入后保存在本机浏览器。"}</small>
      </label>
    </>
  );
}

function FileInput({
  icon,
  label,
  hint,
  accept,
  file,
  onChange,
  guideId
}: {
  icon: React.ReactNode;
  label: string;
  hint: string;
  accept: string;
  file: File | null;
  onChange: (file: File | null) => void;
  guideId?: string;
}) {
  return (
    <label className="file-input" data-guide={guideId}>
      <input
        type="file"
        accept={accept}
        onClick={(event) => {
          event.currentTarget.value = "";
        }}
        onChange={(event) => onChange(event.currentTarget.files?.[0] ?? null)}
      />
      <span className="file-icon">{icon}</span>
      <span>
        <strong>{label}</strong>
        <small>{fileLabel(file, hint)}</small>
      </span>
    </label>
  );
}

function sourceLabel(source: GenericSourceSummary) {
  const kindLabel = source.kind === "pdf" ? "PDF" : source.kind === "derived" ? "结构化" : "Excel";
  return `${kindLabel} / ${source.fileName} / ${source.sheetName}`;
}

function sourceShortLabel(source?: GenericSourceSummary) {
  if (!source) return "";
  const kindLabel = source.kind === "pdf" ? "PDF" : source.kind === "derived" ? "结构化" : "Excel";
  return `${kindLabel} / ${source.sheetName}`;
}

function sourceRoleLabel(source: GenericSourceSummary) {
  const labels: Record<GenericSourceSummary["role"], string> = {
    data: "数据表",
    instruction: "说明页",
    dictionary: "字典表",
    empty: "空表",
    derived: "结构化表",
    issue: "问题"
  };
  return labels[source.role];
}

function buildPresetFormData(matrix: File, pdf: File, config: typeof presetDefaultConfig): FormData {
  const formData = new FormData();
  formData.append("matrix", matrix);
  formData.append("pdf", pdf);
  formData.append("config", JSON.stringify(config));
  return formData;
}

function buildGenericFormData(files: File[], config: unknown): FormData {
  const formData = new FormData();
  files.forEach((file) => formData.append("files", file));
  formData.append("config", JSON.stringify(config));
  return formData;
}

function parseOutputColumns(text: string): GenericOutputColumn[] {
  const columns: GenericOutputColumn[] = [];
  for (const line of text.split("\n").map((item) => item.trim()).filter(Boolean)) {
    const [left, ...rest] = line.split("=");
    const label = left.trim();
    const rawValue = rest.join("=").trim();
    if (!label || !rawValue) continue;
    if (/^".*"$/.test(rawValue)) {
      columns.push({ label, constant: rawValue.slice(1, -1) });
      continue;
    }
    if (/\{[^{}]+\}/.test(rawValue)) {
      columns.push({ label, template: rawValue });
      continue;
    }
    const dot = rawValue.lastIndexOf(".");
    if (dot < 1) continue;
    columns.push({ label, sourceId: rawValue.slice(0, dot), column: rawValue.slice(dot + 1) });
  }
  return columns;
}

function cleanSheetName(value: string) {
  return (value.trim() || "结果").replace(/[\\/?*\[\]:]/g, "_").slice(0, 31) || "结果";
}

function newId(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function setDragPayload(event: React.DragEvent, payload: DragPayload) {
  event.dataTransfer.setData(dragMimeType, JSON.stringify(payload));
  event.dataTransfer.setData("text/plain", JSON.stringify(payload));
  event.dataTransfer.effectAllowed = "copy";
}

function getDragPayload(event: React.DragEvent): DragPayload | null {
  const raw = event.dataTransfer.getData(dragMimeType) || event.dataTransfer.getData("text/plain");
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw) as DragPayload;
    if (parsed.kind === "source" || parsed.kind === "field") return parsed;
  } catch {
    return null;
  }
  return null;
}

function PresetWorkspace({
  accessToken,
  setAccessToken,
  onClose
}: {
  accessToken: string;
  setAccessToken: (value: string) => void;
  onClose: () => void;
}) {
  const [matrixFile, setMatrixFile] = React.useState<File | null>(null);
  const [pdfFile, setPdfFile] = React.useState<File | null>(null);
  const [config, setConfig] = React.useState(presetDefaultConfig);
  const [preview, setPreview] = React.useState<PresetPreview | null>(null);
  const [activeTable, setActiveTable] = React.useState<"project" | "folder" | "review" | "issues">("review");
  const [loading, setLoading] = React.useState(false);
  const [exporting, setExporting] = React.useState(false);
  const [error, setError] = React.useState("");
  const [alert, setAlert] = React.useState<AlertMessage | null>(null);
  const [guideOpen, setGuideOpen] = React.useState(false);
  const [guideStepIndex, setGuideStepIndex] = React.useState(0);
  const canRun = Boolean(matrixFile && pdfFile);

  function showAlert(kind: AlertKind, title: string, message: string) {
    setAlert({ kind, title, message });
  }

  async function requestPreview() {
    if (!matrixFile || !pdfFile) return;
    setLoading(true);
    setError("");
    try {
      const response = await fetch("/api/preset/preview", {
        method: "POST",
        headers: accessToken ? { "x-access-token": accessToken } : undefined,
        body: buildPresetFormData(matrixFile, pdfFile, config)
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.message || payload.error || "预览失败");
      setPreview(payload);
      if (payload.stats.issueCount) {
        showAlert("warning", "预览发现问题", `生成结果包含 ${payload.stats.issueCount} 个问题，请查看问题清单后再导出。`);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : "预览失败";
      setError(message);
      showAlert("error", "预览失败", message);
    } finally {
      setLoading(false);
    }
  }

  async function exportWorkbook() {
    if (!matrixFile || !pdfFile) return;
    setExporting(true);
    setError("");
    try {
      const response = await fetch("/api/preset/export", {
        method: "POST",
        headers: accessToken ? { "x-access-token": accessToken } : undefined,
        body: buildPresetFormData(matrixFile, pdfFile, config)
      });
      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.message || payload.error || "导出失败");
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `固定流程组合结果_${new Date().toISOString().slice(0, 10)}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      const message = err instanceof Error ? err.message : "导出失败";
      setError(message);
      showAlert("error", "导出失败", message);
    } finally {
      setExporting(false);
    }
  }

  const guideSteps: GuideStep[] = [
    {
      id: "token",
      title: "填写访问口令",
      body: "服务器开启口令时，首次填写即可，浏览器会保存到本机，之后生成预览或导出会直接带上。",
      target: "[data-guide='preset-token']",
      action: "点访问口令，首次输入口令",
      done: Boolean(accessToken),
      optional: Boolean(accessToken)
    },
    {
      id: "files",
      title: "上传两份源文件",
      body: "这里需要矩阵 Excel 和主项 PDF 两个文件，缺任意一个都不能生成。",
      target: "[data-guide='preset-files']",
      action: "分别上传 .xlsx 和 .pdf",
      done: canRun
    },
    {
      id: "config",
      title: "确认规则参数",
      body: "默认参数已经按当前流程填好。只在目录名、设计文档名或序号前缀变化时修改。",
      target: "[data-guide='preset-config']",
      action: "检查参数，通常无需修改",
      done: true,
      optional: true
    },
    {
      id: "preview",
      title: "生成预览",
      body: "先生成预览，核对主项台账、审阅流程和问题清单，再决定是否导出。",
      target: "[data-guide='preset-actions']",
      action: "点击生成预览",
      done: Boolean(preview)
    },
    {
      id: "export",
      title: "导出 Excel",
      body: "预览结果确认后，导出当前固定流程结果。",
      target: "[data-guide='preset-actions']",
      action: "点击导出 Excel",
      done: false
    }
  ];
  const nextGuideIndex = Math.max(0, guideSteps.findIndex((step) => !step.done && !step.optional));

  function startPresetGuide(index = nextGuideIndex >= 0 ? nextGuideIndex : 0) {
    setGuideStepIndex(index);
    setGuideOpen(true);
  }

  return (
    <main className="app-shell preset-shell">
      <AlertPopup alert={alert} onClose={() => setAlert(null)} />
      <AnimatePresence>
        {guideOpen && (
          <GuidedTour
            steps={guideSteps}
            activeIndex={Math.min(guideStepIndex, guideSteps.length - 1)}
            setActiveIndex={setGuideStepIndex}
            onClose={() => setGuideOpen(false)}
            title="固定流程教程"
          />
        )}
      </AnimatePresence>
      <section className="left-pane">
        <motion.header className="preset-header" initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.28 }}>
          <div className="preset-header-top">
            <p className="eyebrow">固定流程</p>
            <div className="preset-header-actions">
              <button className="guide-entry-button" onClick={() => startPresetGuide(0)} title="打开固定流程教程">
                <Info size={15} />教程
              </button>
              <button className="preset-close-button" onClick={onClose} title="返回通用组合">
                <X size={16} />返回
              </button>
            </div>
          </div>
          <details className="header-settings preset-settings" data-guide="preset-token">
            <summary><Settings2 size={15} />访问口令</summary>
            <label className="field token-field">
              <span>访问口令</span>
              <input
                type="password"
                value={accessToken}
                onChange={(event) => setAccessToken(event.target.value)}
                placeholder="输入访问口令，本机保存"
              />
              <small className="token-save-hint">{accessToken ? "已保存在本机浏览器，清空输入框可删除。" : "输入后保存在本机浏览器。"}</small>
            </label>
          </details>
        </motion.header>
        <motion.div className="panel upload-panel" data-guide="preset-files" initial={{ opacity: 0, y: 16 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.36, delay: 0.08 }}>
          <div className="panel-title"><UploadCloud size={18} /><span>源文件</span></div>
          <FileInput icon={<FileSpreadsheet size={22} />} label="文档矩阵" hint="上传 .xlsx" accept=".xlsx" file={matrixFile} onChange={setMatrixFile} />
          <FileInput icon={<FileText size={22} />} label="项目主项表及分工" hint="上传 .pdf" accept=".pdf" file={pdfFile} onChange={setPdfFile} />
        </motion.div>
        <motion.div className="panel" data-guide="preset-config" initial={{ opacity: 0, y: 16 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.36, delay: 0.14 }}>
          <div className="panel-title"><FileSpreadsheet size={18} /><span>规则参数</span></div>
          <label className="field"><span>施工图根目录</span><input value={config.rootPath} onChange={(event) => setConfig({ ...config, rootPath: event.target.value })} /></label>
          <div className="field-grid">
            <label className="field"><span>设计文档</span><input value={config.designDocumentName} onChange={(event) => setConfig({ ...config, designDocumentName: event.target.value })} /></label>
            <label className="field"><span>序号前缀</span><input value={config.recordPrefix} onChange={(event) => setConfig({ ...config, recordPrefix: event.target.value })} /></label>
          </div>
        </motion.div>
        <div className="action-row" data-guide="preset-actions">
          <button className="primary" disabled={!canRun || loading} onClick={requestPreview}>{loading ? <Loader2 className="spin" size={18} /> : <CheckCircle2 size={18} />}生成预览</button>
          <button className="secondary" disabled={!canRun || exporting} onClick={exportWorkbook}>{exporting ? <Loader2 className="spin" size={18} /> : <Download size={18} />}导出 Excel</button>
        </div>
      </section>
      <section className="right-pane">
        <div className="workspace-header">
          <div><p className="eyebrow">预览</p><h2>生成结果核查</h2></div>
          {preview && <span className={preview.stats.issueCount ? "badge warn" : "badge ok"}>{preview.stats.issueCount} 个问题</span>}
        </div>
        {preview ? (
          <>
            <motion.div className="stats-grid" initial={{ opacity: 0 }} animate={{ opacity: 1 }} transition={{ duration: 0.28 }}>
              <Stat label="单元" value={preview.stats.unitCount} />
              <Stat label="装置" value={preview.stats.deviceCount} />
              <Stat label="专业" value={preview.stats.disciplineCount} />
              <Stat label="文件夹属性" value={preview.stats.folderRowCount} />
              <Stat label="审阅流程" value={preview.stats.reviewRowCount} />
            </motion.div>
            <div className="tabs">
              <button className={activeTable === "project" ? "active" : ""} onClick={() => setActiveTable("project")}>主项台账</button>
              <button className={activeTable === "folder" ? "active" : ""} onClick={() => setActiveTable("folder")}>文件夹属性</button>
              <button className={activeTable === "review" ? "active" : ""} onClick={() => setActiveTable("review")}>审阅流程</button>
              <button className={activeTable === "issues" ? "active" : ""} onClick={() => setActiveTable("issues")}>问题清单</button>
            </div>
            <div className="table-shell">
              {activeTable === "project" && <table><thead><tr><th>编码</th><th>名称</th><th>功能区</th><th>单元</th><th>设计单位</th></tr></thead><tbody>{preview.projectPreview.map((row) => <tr key={`${row.code}-${row.name}`}><td>{row.code}</td><td>{row.name}</td><td>{row.areaName}</td><td>{row.unitName}</td><td>{row.designCompany}</td></tr>)}</tbody></table>}
              {activeTable === "folder" && <table><thead><tr><th>路径</th><th>名称</th><th>层级</th><th>装置</th><th>单元</th><th>设计单位</th></tr></thead><tbody>{preview.folderPreview.map((row, index) => <tr key={`${row.folderPath}-${row.name}-${index}`}><td title={row.folderPath}>{shortPath(row.folderPath)}</td><td>{row.name}</td><td>{row.levelName}</td><td>{row.device}</td><td>{row.unit}</td><td>{row.designCompany}</td></tr>)}</tbody></table>}
              {activeTable === "review" && <table><thead><tr><th>序号</th><th>功能区</th><th>单元</th><th>装置</th><th>专业</th><th>PMC</th><th>专业工程师</th><th>负责人</th></tr></thead><tbody>{preview.reviewPreview.map((row) => <tr key={row.recordNo}><td>{row.recordNo}</td><td>{row.functionAreaName}</td><td>{row.unitName}</td><td>{row.deviceName}</td><td>{row.disciplineName}</td><td>{row.digitalPmc}</td><td>{row.specialtyEngineer}</td><td>{row.professionalLeader}</td></tr>)}</tbody></table>}
              {activeTable === "issues" && <table><thead><tr><th>类型</th><th>级别</th><th>说明</th><th>来源</th><th>关键值</th></tr></thead><tbody>{preview.issues.map((row, index) => <tr key={`${row.type}-${index}`}><td>{row.type}</td><td>{row.severity}</td><td>{row.message}</td><td>{row.source || ""}{row.row ? `:${row.row}` : ""}</td><td>{row.key || row.value || ""}</td></tr>)}</tbody></table>}
            </div>
          </>
        ) : (
          <div className="empty-state"><FileSpreadsheet size={38} /><h2>等待源文件</h2><p>上传矩阵 Excel 和主项 PDF 后生成预览，导出时会得到一个包含三张工作表的新 Excel。</p></div>
        )}
      </section>
    </main>
  );
}

function GenericWorkspace({
  accessToken,
  setAccessToken,
  onOpenPreset
}: {
  accessToken: string;
  setAccessToken: (value: string) => void;
  onOpenPreset: () => void;
}) {
  const [files, setFiles] = React.useState<File[]>([]);
  const [sources, setSources] = React.useState<GenericSourceSummary[]>([]);
  const [sourceHeaderRows, setSourceHeaderRows] = React.useState<Record<string, number>>({});
  const [sourceDataStartRows, setSourceDataStartRows] = React.useState<Record<string, number>>({});
  const [sourceFillDownColumns, setSourceFillDownColumns] = React.useState<Record<string, string>>({});
  const [includeAuxiliarySheets, setIncludeAuxiliarySheets] = React.useState(false);
  const [templateExportEnabled, setTemplateExportEnabled] = React.useState(false);
  const [templateExport, setTemplateExport] = React.useState<GenericTemplateExportDraft>({ sourceId: "" });
  const [baseSourceId, setBaseSourceId] = React.useState("");
  const [joins, setJoins] = React.useState<GenericJoinDraft[]>([]);
  const [outputColumns, setOutputColumns] = React.useState<GenericOutputColumnDraft[]>([]);
  const [fieldQuery, setFieldQuery] = React.useState("");
  const [rightView, setRightView] = React.useState<RightView>("fields");
  const [focusPreview, setFocusPreview] = React.useState(false);
  const [advancedOpen, setAdvancedOpen] = React.useState(false);
  const [outputText, setOutputText] = React.useState("");
  const [constantLabel, setConstantLabel] = React.useState("");
  const [constantValue, setConstantValue] = React.useState("");
  const [templateLabel, setTemplateLabel] = React.useState("");
  const [templateValue, setTemplateValue] = React.useState("");
  const templateTextareaRef = React.useRef<HTMLTextAreaElement | null>(null);
  const [notice, setNotice] = React.useState("");
  const [dragging, setDragging] = React.useState<DragPayload | null>(null);
  const [pointerDrag, setPointerDrag] = React.useState<PointerDragState | null>(null);
  const [activeDrop, setActiveDrop] = React.useState("");
  const [preview, setPreview] = React.useState<GenericPreview | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [exporting, setExporting] = React.useState(false);
  const [error, setError] = React.useState("");
  const [alert, setAlert] = React.useState<AlertMessage | null>(null);
  const [guideOpen, setGuideOpen] = React.useState(false);
  const [guideStepIndex, setGuideStepIndex] = React.useState(0);
  const sourceMap = React.useMemo(() => new Map(sources.map((source) => [source.id, source])), [sources]);
  const baseSource = sourceMap.get(baseSourceId);

  function showAlert(kind: AlertKind, title: string, message: string) {
    setAlert({ kind, title, message });
  }

  React.useEffect(() => {
    if (!pointerDrag) return undefined;
    const drag = pointerDrag;

    function currentDropTarget(x: number, y: number) {
      return document.elementFromPoint(x, y)?.closest<HTMLElement>("[data-drop-target]")?.dataset.dropTarget ?? "";
    }

    function onPointerMove(event: PointerEvent) {
      const distance = Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY);
      const active = drag.active || distance > 4;
      setPointerDrag({ ...drag, x: event.clientX, y: event.clientY, active });
      if (active) {
        setDragging(drag);
        setActiveDrop(currentDropTarget(event.clientX, event.clientY));
      }
    }

    function onPointerUp(event: PointerEvent) {
      if (drag.active || Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) > 4) {
        applyPointerDrop(currentDropTarget(event.clientX, event.clientY), drag);
      }
      setPointerDrag(null);
      finishDrag();
    }

    function onMouseMove(event: MouseEvent) {
      const distance = Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY);
      const active = drag.active || distance > 4;
      setPointerDrag({ ...drag, x: event.clientX, y: event.clientY, active });
      if (active) {
        setDragging(drag);
        setActiveDrop(currentDropTarget(event.clientX, event.clientY));
      }
    }

    function onMouseUp(event: MouseEvent) {
      if (drag.active || Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) > 4) {
        applyPointerDrop(currentDropTarget(event.clientX, event.clientY), drag);
      }
      setPointerDrag(null);
      finishDrag();
    }

    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
    return () => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
    };
  }, [pointerDrag, joins, baseSourceId, outputColumns]);

  function resetWorkspaceState(nextNotice = "") {
    setSources([]);
    setSourceHeaderRows({});
    setSourceDataStartRows({});
    setSourceFillDownColumns({});
    setBaseSourceId("");
    setTemplateExport({ sourceId: "" });
    setTemplateExportEnabled(false);
    setJoins([]);
    setOutputColumns([]);
    setFieldQuery("");
    setOutputText("");
    setTemplateLabel("");
    setTemplateValue("");
    setPreview(null);
    setRightView("fields");
    setError("");
    setAlert(null);
    setNotice(nextNotice);
  }

  function handleFilesChange(nextFiles: File[]) {
    setFiles(nextFiles);
    resetWorkspaceState(nextFiles.length ? "文件已选择，请重新解析表和字段。" : "");
  }

  function genericConfig(options?: { resultSheets?: GenericResultSheetConfig[] }) {
    const draggedColumns = outputColumns.map(({ id: _id, ...column }) => column);
    const textColumns = parseOutputColumns(outputText);
    const configuredColumns = [...draggedColumns, ...textColumns];
    const sourceIds = new Set([
      ...Object.keys(sourceHeaderRows),
      ...Object.keys(sourceDataStartRows),
      ...Object.keys(sourceFillDownColumns)
    ]);
    return {
      baseSourceId: effectiveBaseSourceId() || undefined,
      structuredExtraction: "none",
      includeAuxiliarySheets,
      sourceConfigs: Array.from(sourceIds).map((sourceId) => {
        const fillDownColumns = sourceFillDownColumns[sourceId]
          ?.split(/[，,；;、]/)
          .map((item) => item.trim())
          .filter(Boolean);
        return {
          sourceId,
          headerRow: sourceHeaderRows[sourceId],
          dataStartRow: sourceDataStartRows[sourceId],
          fillDownColumns: fillDownColumns?.length ? fillDownColumns : undefined
        };
      }),
      joins: joins
        .filter((join) => join.sourceId && join.leftKey && join.rightKey)
        .map(({ id: _id, ...join }) => join),
      outputColumns: configuredColumns.length ? configuredColumns : undefined,
      templateExport: templateExportEnabled && templateExport.sourceId
        ? {
          sourceId: templateExport.sourceId,
          headerRow: sourceHeaderRows[templateExport.sourceId],
          dataStartRow: sourceDataStartRows[templateExport.sourceId]
        }
        : undefined,
      resultSheets: options?.resultSheets?.length ? options.resultSheets : undefined
    };
  }

  function allowDrop(event: React.DragEvent) {
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
  }

  function beginDrag(event: React.DragEvent, payload: DragPayload) {
    setDragPayload(event, payload);
    setDragging(payload);
  }

  function beginPointerDrag(event: React.PointerEvent, payload: DragPayload) {
    if ((event.target as HTMLElement).closest("button,input,select,textarea")) return;
    setPointerDrag({
      ...payload,
      startX: event.clientX,
      startY: event.clientY,
      x: event.clientX,
      y: event.clientY,
      active: false
    });
  }

  function beginMouseDrag(event: React.MouseEvent, payload: DragPayload) {
    if (event.button !== 0 || (event.target as HTMLElement).closest("button,input,select,textarea")) return;
    setPointerDrag({
      ...payload,
      startX: event.clientX,
      startY: event.clientY,
      x: event.clientX,
      y: event.clientY,
      active: false
    });
  }

  function finishDrag() {
    setDragging(null);
    setActiveDrop("");
  }

  function enterDrop(event: React.DragEvent, target: string) {
    allowDrop(event);
    setActiveDrop(target);
  }

  function setBaseFromPayload(payload: DragPayload | null) {
    if (!payload) return;
    setBaseSourceId(payload.sourceId);
  }

  function uniqueOutputLabelFrom(current: GenericOutputColumnDraft[], label: string) {
    const used = new Set(current.map((column) => column.label));
    if (!used.has(label)) return label;
    let suffix = 2;
    while (used.has(`${label}_${suffix}`)) suffix += 1;
    return `${label}_${suffix}`;
  }

  function warnManualJoinNeeded(sourceId: string) {
    const source = sourceMap.get(sourceId);
    const currentBaseSource = effectiveBaseSource();
    if (!source || !currentBaseSource || source.id === currentBaseSource.id) return;
    if (joins.some((join) => join.sourceId === source.id && join.leftKey && join.rightKey)) return;
    const message = `${sourceShortLabel(source)} 的字段已加入结果列；请手动添加补充表，并选择两边相同含义的字段。`;
    setNotice(message);
    showAlert("warning", "需要手动设置补充表", message);
  }

  function appendOutputField(sourceId: string, column: string) {
    const source = sourceMap.get(sourceId);
    if (!source) return;
    const currentBaseSourceId = effectiveBaseSourceId();
    if (!currentBaseSourceId) {
      setNotice("请先手动选择主表，主表决定结果表有多少行。");
    } else if (source.id !== currentBaseSourceId) {
      warnManualJoinNeeded(source.id);
    }
    setOutputColumns((current) => {
      if (current.some((item) => item.sourceId === sourceId && item.column === column && item.constant == null)) {
        return current;
      }
      return [
        ...current,
        {
          id: newId("out"),
          label: uniqueOutputLabelFrom(current, column),
          sourceId,
          column,
          constant: undefined
        }
      ];
    });
  }

  function appendOutputFields(sourceId: string, headers: string[]) {
    const source = sourceMap.get(sourceId);
    if (!source || !headers.length) return;
    const currentBaseSourceId = effectiveBaseSourceId();
    if (!currentBaseSourceId) {
      setNotice("请先手动选择主表，主表决定结果表有多少行。");
    } else if (source.id !== currentBaseSourceId) {
      warnManualJoinNeeded(source.id);
    }
    setOutputColumns((current) => {
      let next = [...current];
      for (const header of headers) {
        if (next.some((item) => item.sourceId === sourceId && item.column === header && item.constant == null)) continue;
        next = [
          ...next,
          {
            id: newId("out"),
            label: uniqueOutputLabelFrom(next, header),
            sourceId,
            column: header,
            constant: undefined
          }
        ];
      }
      return next;
    });
  }

  function fieldTemplateToken(source: GenericSourceSummary, column: string) {
    const duplicated = sources.filter((item) => item.headers.includes(column)).length > 1;
    return duplicated && source.sheetName ? `{${source.sheetName}.${column}}` : `{${column}}`;
  }

  function insertIntoTemplate(token: string) {
    const textarea = templateTextareaRef.current;
    const selectionStart = textarea?.selectionStart;
    const selectionEnd = textarea?.selectionEnd;
    let cursorPosition = 0;
    setTemplateValue((current) => {
      const start = typeof selectionStart === "number" ? Math.min(selectionStart, current.length) : current.length;
      const end = typeof selectionEnd === "number" ? Math.min(selectionEnd, current.length) : start;
      cursorPosition = start + token.length;
      return `${current.slice(0, start)}${token}${current.slice(end)}`;
    });
    window.setTimeout(() => {
      const nextTextarea = templateTextareaRef.current;
      if (!nextTextarea) return;
      nextTextarea.focus();
      nextTextarea.setSelectionRange(cursorPosition, cursorPosition);
    }, 0);
  }

  function appendFieldToTemplate(sourceId: string, column: string) {
    const source = sourceMap.get(sourceId);
    if (!source) return;
    const currentBaseSourceId = effectiveBaseSourceId();
    if (!currentBaseSourceId) {
      setNotice("请先手动选择主表，主表决定结果表有多少行。");
    } else if (source.id !== currentBaseSourceId) {
      warnManualJoinNeeded(source.id);
    }
    insertIntoTemplate(fieldTemplateToken(source, column));
    setNotice(`已把“${column}”插入组合内容。`);
  }

  function moveOutputColumn(index: number, direction: -1 | 1) {
    setOutputColumns((current) => {
      const target = index + direction;
      if (target < 0 || target >= current.length) return current;
      const next = [...current];
      [next[index], next[target]] = [next[target], next[index]];
      return next;
    });
  }

  function normalized(value: string) {
    return value.replace(/\s+/g, "").toLowerCase();
  }

  function fieldMatches(source: GenericSourceSummary, header: string, query: string) {
    if (!query) return true;
    const sample = source.previewRows.find((row) => row[header])?.[header] ?? "";
    const haystack = `${source.fileName} ${source.sheetName} ${header} ${sample}`.toLowerCase();
    return haystack.includes(query.toLowerCase());
  }

  function visibleHeaders(source: GenericSourceSummary) {
    return source.headers.filter((header) => fieldMatches(source, header, fieldQuery));
  }

  function visibleSources() {
    return sources.filter((source) => {
      if (!fieldQuery.trim()) return true;
      return sourceLabel(source).toLowerCase().includes(fieldQuery.toLowerCase()) || visibleHeaders(source).length > 0;
    });
  }

  function handleOutputDrop(event: React.DragEvent) {
    event.preventDefault();
    const payload = getDragPayload(event);
    if (payload?.kind === "field") {
      appendOutputField(payload.sourceId, payload.column);
    }
    finishDrag();
  }

  function updateJoin(index: number, nextJoin: GenericJoinDraft) {
    setJoins((current) => current.map((join, itemIndex) => (itemIndex === index ? nextJoin : join)));
  }

  function applyJoinSourcePayload(index: number, payload: DragPayload) {
    const source = sourceMap.get(payload.sourceId);
    const join = joins[index];
    if (!join) return;
    updateJoin(index, {
      ...join,
      sourceId: payload.sourceId,
      rightKey: payload.kind === "field" ? payload.column : join.rightKey,
      prefix: source?.sheetName ?? join.prefix
    });
  }

  function applyJoinKeyPayload(index: number, side: "left" | "right", payload: DragPayload) {
    if (payload.kind !== "field") return;
    const source = sourceMap.get(payload.sourceId);
    const join = joins[index];
    if (!join) return;
    if (side === "left") {
      setBaseSourceId(payload.sourceId);
      updateJoin(index, { ...join, leftKey: payload.column });
      return;
    }
    updateJoin(index, {
      ...join,
      sourceId: payload.sourceId,
      rightKey: payload.column,
      prefix: source?.sheetName ?? join.prefix
    });
  }

  function applyPointerDrop(target: string, payload: DragPayload) {
    if (target === "output" && payload.kind === "field") {
      appendOutputField(payload.sourceId, payload.column);
      return;
    }
    if (target === "base") {
      setBaseFromPayload(payload);
      return;
    }
    const joinSourceId = target.match(/^join-source-(.+)$/)?.[1];
    if (joinSourceId) {
      const index = joins.findIndex((join) => join.id === joinSourceId);
      if (index >= 0) applyJoinSourcePayload(index, payload);
      return;
    }
    const joinLeftId = target.match(/^join-left-(.+)$/)?.[1];
    if (joinLeftId) {
      const index = joins.findIndex((join) => join.id === joinLeftId);
      if (index >= 0) applyJoinKeyPayload(index, "left", payload);
      return;
    }
    const joinRightId = target.match(/^join-right-(.+)$/)?.[1];
    if (joinRightId) {
      const index = joins.findIndex((join) => join.id === joinRightId);
      if (index >= 0) applyJoinKeyPayload(index, "right", payload);
    }
  }

  function handleJoinSourceDrop(index: number, event: React.DragEvent) {
    event.preventDefault();
    const payload = getDragPayload(event);
    if (!payload) return;
    applyJoinSourcePayload(index, payload);
    finishDrag();
  }

  function handleJoinKeyDrop(index: number, side: "left" | "right", event: React.DragEvent) {
    event.preventDefault();
    const payload = getDragPayload(event);
    if (!payload) return;
    applyJoinKeyPayload(index, side, payload);
    finishDrag();
  }

  function addConstantColumn() {
    const label = constantLabel.trim();
    if (!label) {
      showAlert("warning", "固定列缺少列名", "请先填写固定列名，再添加固定值列。");
      return;
    }
    setOutputColumns((current) => [
      ...current,
      { id: newId("out"), label: uniqueOutputLabelFrom(current, label), constant: constantValue }
    ]);
    setConstantLabel("");
    setConstantValue("");
  }

  function addTemplateColumn() {
    const label = templateLabel.trim() || "组合列";
    const template = templateValue.trim();
    if (!/\{[^{}]+\}/.test(template)) {
      showAlert("warning", "组合列缺少字段", "请先在组合内容里加入字段，例如 {代码}/{名称}。");
      return;
    }
    setOutputColumns((current) => [
      ...current,
      { id: newId("out"), label: uniqueOutputLabelFrom(current, label), template }
    ]);
    setTemplateLabel("");
    setTemplateValue("");
    setNotice(`已添加组合列“${label}”。`);
  }

  function canOutputSource(source: GenericSourceSummary) {
    return source.rowCount > 0 && source.headers.length > 0 && source.role !== "empty";
  }

  function implicitSingleBaseSource() {
    if (baseSourceId) return undefined;
    const candidates = sources.filter(canOutputSource);
    return candidates.length === 1 ? candidates[0] : undefined;
  }

  function effectiveBaseSourceId() {
    return baseSourceId || implicitSingleBaseSource()?.id || "";
  }

  function effectiveBaseSource() {
    return baseSource ?? implicitSingleBaseSource();
  }

  function defaultSingleSheetName(source: GenericSourceSummary) {
    return cleanSheetName(source.sheetName || source.fileName.replace(/\.[^.]+$/, ""));
  }

  function downloadBlob(blob: Blob, fileName: string) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);
  }

  async function exportWithConfig(config: unknown, fileName: string) {
    if (!files.length) return;
    setExporting(true);
    setError("");
    try {
      const response = await fetch("/api/generic/export", {
        method: "POST",
        headers: accessToken ? { "x-access-token": accessToken } : undefined,
        body: buildGenericFormData(files, config)
      });
      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.message || payload.error || "导出失败");
      }
      const blob = await response.blob();
      downloadBlob(blob, fileName);
    } catch (err) {
      const message = err instanceof Error ? err.message : "导出失败";
      setError(message);
      showAlert("error", "导出失败", message);
    } finally {
      setExporting(false);
    }
  }

  async function exportSourceSheet(sourceId: string) {
    const source = sourceMap.get(sourceId);
    if (!source) return;
    await exportWithConfig(
      genericConfig({ resultSheets: [{ sourceId, sheetName: defaultSingleSheetName(source) }] }),
      `${defaultSingleSheetName(source)}_${new Date().toISOString().slice(0, 10)}.xlsx`
    );
  }

  async function inspectSources() {
    if (!files.length) return;
    setLoading(true);
    setError("");
    try {
      const response = await fetch("/api/generic/inspect", {
        method: "POST",
        headers: accessToken ? { "x-access-token": accessToken } : undefined,
        body: buildGenericFormData(files, {
          structuredExtraction: "none",
          includeAuxiliarySheets,
          sourceConfigs: Object.entries(sourceHeaderRows).map(([sourceId, headerRow]) => ({
            sourceId,
            headerRow,
            dataStartRow: sourceDataStartRows[sourceId],
            fillDownColumns: sourceFillDownColumns[sourceId]?.split(/[，,；;、]/).map((item) => item.trim()).filter(Boolean)
          }))
        })
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.message || payload.error || "解析失败");
      setSources(payload.sources);
      setSourceHeaderRows(Object.fromEntries(payload.sources.map((source: GenericSourceSummary) => [source.id, source.defaultHeaderRow])));
      setSourceDataStartRows(Object.fromEntries(payload.sources.map((source: GenericSourceSummary) => [source.id, source.defaultDataStartRow])));
      setSourceFillDownColumns(Object.fromEntries(payload.sources.map((source: GenericSourceSummary) => [source.id, source.suggestedFillDownColumns.join(",")])));
      setBaseSourceId("");
      setTemplateExport({ sourceId: "" });
      setTemplateExportEnabled(false);
      setJoins([]);
      setOutputColumns([]);
      setFieldQuery("");
      setOutputText("");
      setPreview(null);
      setRightView("fields");
      setNotice(`已解析 ${payload.sources.length} 张可用表。`);
    } catch (err) {
      const message = err instanceof Error ? err.message : "解析失败";
      setError(message);
      showAlert("error", "解析失败", message);
    } finally {
      setLoading(false);
    }
  }

  async function requestPreview() {
    if (!files.length) return;
    if (!sources.length) {
      showAlert("warning", "请先解析", "请先解析表和字段，再生成预览。");
      return;
    }
    if (!effectiveBaseSourceId()) {
      showAlert("warning", "请先选择主表", "多张表时需要先选择哪张表决定输出行数。");
      return;
    }
    setLoading(true);
    setError("");
    try {
      const response = await fetch("/api/generic/preview", {
        method: "POST",
        headers: accessToken ? { "x-access-token": accessToken } : undefined,
        body: buildGenericFormData(files, genericConfig())
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.message || payload.error || "预览失败");
      setPreview(payload);
      if (!sources.length) {
        setSources(payload.sources);
        setSourceHeaderRows(Object.fromEntries(payload.sources.map((source: GenericSourceSummary) => [source.id, source.defaultHeaderRow])));
        setSourceDataStartRows(Object.fromEntries(payload.sources.map((source: GenericSourceSummary) => [source.id, source.defaultDataStartRow])));
        setSourceFillDownColumns(Object.fromEntries(payload.sources.map((source: GenericSourceSummary) => [source.id, source.suggestedFillDownColumns.join(",")])));
        setTemplateExport({ sourceId: "" });
      }
      setRightView("result");
      setNotice(`预览已生成：${payload.stats.outputRowCount} 行，${payload.stats.issueCount} 个问题。`);
      if (payload.stats.issueCount) {
        showAlert("warning", "预览发现问题", `组合结果已生成，但发现 ${payload.stats.issueCount} 个问题，请切换到“问题”页查看。`);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : "预览失败";
      setError(message);
      showAlert("error", "预览失败", message);
    } finally {
      setLoading(false);
    }
  }

  async function exportGeneric() {
    if (!files.length) return;
    if (!sources.length) {
      showAlert("warning", "请先解析", "请先解析表和字段，再导出结果表。");
      return;
    }
    if (!effectiveBaseSourceId()) {
      showAlert("warning", "请先选择主表", "多张表时需要先选择哪张表决定输出行数。");
      return;
    }
    await exportWithConfig(genericConfig(), `通用组合结果_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function addJoin() {
    setJoins([...joins, {
      id: newId("join"),
      sourceId: "",
      leftKey: "",
      rightKey: "",
      joinType: "left",
      duplicateStrategy: "first",
      multiValueSeparator: ",",
      prefix: ""
    }]);
  }

  const previewRows = preview?.outputPreview ?? [];
  const previewHeaders = previewRows[0] ? Object.keys(previewRows[0]) : [];
  const issueRows = preview?.issues ?? [];
  const excelSources = sources.filter((source) => source.kind === "excel");
  const templateSource = sourceMap.get(templateExport.sourceId);
  const activeBaseSource = effectiveBaseSource();
  const activeBaseSourceId = effectiveBaseSourceId();
  const unjoinedOutputCount = outputColumns.filter((column) => {
    if (!column.sourceId || column.sourceId === activeBaseSourceId) return false;
    return !joins.some((join) => join.sourceId === column.sourceId && join.leftKey && join.rightKey);
  }).length;
  const canPreviewGeneric = Boolean(files.length && sources.length && activeBaseSourceId && !loading);
  const canExportGeneric = Boolean(files.length && sources.length && activeBaseSourceId && !exporting);
  const baseStatusText = activeBaseSource
    ? sourceShortLabel(activeBaseSource)
    : sources.length > 1
      ? `${sources.length} 张表，先选主表`
      : sources.length
        ? "当前表"
        : files.length
          ? "待解析"
          : "待上传";
  const columnStatusText = outputColumns.length ? `${outputColumns.length} 列` : "全部字段";
  const guidanceText = !files.length
    ? "先上传要组合的 Excel 或 PDF。"
    : !sources.length
      ? "文件已选择，下一步解析出表和字段。"
      : !activeBaseSource
        ? "多张表时先选择主表。"
        : unjoinedOutputCount
          ? `${unjoinedOutputCount} 个结果列来自未设置关联的表。`
          : preview
            ? "预览已生成，可导出本次结果表。"
            : "配置完成后先预览，确认结果正确再导出。";
  const guideSteps: GuideStep[] = [
    {
      id: "token",
      title: "填写访问口令",
      body: "服务器开启口令时，首次在左上角“设置”里输入口令即可，之后会直接带上。",
      target: "[data-guide='generic-token']",
      action: "点“设置”，首次输入访问口令",
      done: Boolean(accessToken),
      optional: Boolean(accessToken)
    },
    {
      id: "files",
      title: "上传源文件",
      body: "选择要组合的 Excel 或 PDF。可以一次选择多个文件，系统会把里面的 sheet 和 PDF 文本行整理成可用表。",
      target: "[data-guide='generic-files']",
      action: "点击上传区域，选择文件",
      done: files.length > 0
    },
    {
      id: "inspect",
      title: "解析表和字段",
      body: "点解析后，右侧会出现表和字段。表名独立一行，字段在表名下面。",
      target: "[data-guide='generic-inspect']",
      action: "点击解析表和字段",
      done: sources.length > 0
    },
    {
      id: "fields",
      title: "选择主表和字段",
      body: "右侧可以设为主表、搜索字段，也可以点“加入”。需要精简列时，只保留本次目标格式需要的字段。",
      target: "[data-guide='generic-field-bank']",
      action: "设为主表并加入结果列",
      done: Boolean(activeBaseSource),
      optional: Boolean(activeBaseSource)
    },
    {
      id: "preview",
      title: "生成预览",
      body: "导出前先预览。系统会给出前 100 行结果，并把未匹配、重复键等问题列出来。",
      target: "[data-guide='generic-preview']",
      action: "点击生成预览",
      done: Boolean(preview)
    },
    {
      id: "workbook",
      title: "可选：自己选择",
      body: "在这里选择主表、补充表、两边相同含义的字段，以及最终结果列。",
      target: "[data-guide='generic-advanced']",
      action: "展开选择主表和补充表",
      done: advancedOpen,
      optional: !advancedOpen
    },
    {
      id: "export",
      title: "导出 Excel",
      body: "预览没问题后导出本次结果表；如果还有其他目标格式，重新配置后再单独导出。",
      target: "[data-guide='generic-export']",
      action: "点击导出组合",
      done: false
    }
  ];
  const firstTodoGuideIndex = guideSteps.findIndex((step) => !step.done && !step.optional);
  const nextGuideIndex = firstTodoGuideIndex >= 0 ? firstTodoGuideIndex : guideSteps.length - 1;
  const nextGuideStep = guideSteps[nextGuideIndex];
  const guideCompletedCount = guideSteps.filter((step) => step.done || step.optional).length;

  function startGenericGuide(index = nextGuideIndex >= 0 ? nextGuideIndex : 0) {
    setGuideStepIndex(index);
    setGuideOpen(true);
  }

  return (
    <main className={`app-shell generic-shell ${focusPreview ? "focus-preview" : ""}`}>
      <AlertPopup alert={alert} onClose={() => setAlert(null)} />
      <AnimatePresence>
        {guideOpen && (
          <GuidedTour
            steps={guideSteps}
            activeIndex={Math.min(guideStepIndex, guideSteps.length - 1)}
            setActiveIndex={setGuideStepIndex}
            onClose={() => setGuideOpen(false)}
            title="通用组合教程"
          />
        )}
      </AnimatePresence>
      {pointerDrag?.active && (
        <div className="drag-ghost" style={{ left: pointerDrag.x + 12, top: pointerDrag.y + 12 }}>
          {pointerDrag.kind === "field" ? pointerDrag.column : sourceShortLabel(sourceMap.get(pointerDrag.sourceId))}
        </div>
      )}
      <section className="left-pane">
        <HeaderControls accessToken={accessToken} setAccessToken={setAccessToken} onOpenPreset={onOpenPreset} onStartGuide={() => startGenericGuide(0)} compact />
        <div className="panel upload-panel">
          <div className="panel-title"><UploadCloud size={18} /><span>通用源文件</span></div>
          <label className="file-input generic-files" data-guide="generic-files">
            <input
              type="file"
              accept=".xlsx,.pdf"
              multiple
              onClick={(event) => {
                event.currentTarget.value = "";
              }}
              onChange={(event) => handleFilesChange(Array.from(event.currentTarget.files ?? []))}
            />
            <span className="file-icon"><FileSpreadsheet size={22} /></span>
            <span><strong>上传 Excel / PDF</strong><small>{files.length ? files.map((file) => file.name).join("；") : "可一次选择多个文件"}</small></span>
          </label>
          <details className="parse-options">
            <summary>少用选项</summary>
            <label className="toggle-row">
              <input
                type="checkbox"
                checked={includeAuxiliarySheets}
                onChange={(event) => setIncludeAuxiliarySheets(event.target.checked)}
              />
              <span>同时显示说明页、字典表和空表</span>
            </label>
          </details>
          <div className="action-row" data-guide="generic-inspect">
            <button className="primary" disabled={!files.length || loading} onClick={inspectSources} title="解析表和字段">{loading ? <Loader2 className="spin" size={18} /> : <CheckCircle2 size={18} />}解析</button>
            <button className="secondary" disabled={!canPreviewGeneric} onClick={requestPreview} title="生成预览">{loading ? <Loader2 className="spin" size={18} /> : <CheckCircle2 size={18} />}预览</button>
          </div>
        </div>
        <div className="panel simple-combine-panel" data-guide="generic-simple">
          <div className="panel-title simple-panel-title">
            <span><Sparkles size={18} />生成一张结果表</span>
          </div>
          <div className="simple-inline-status">
            <span>表：<strong>{baseStatusText}</strong></span>
            <span>列：<strong>{columnStatusText}</strong></span>
          </div>
          <div className="simple-action-grid">
            <button className="secondary" data-guide="generic-preview" disabled={!canPreviewGeneric} onClick={requestPreview}>
              {loading ? <Loader2 className="spin" size={17} /> : <Eye size={17} />}预览
            </button>
            <button className="secondary" data-guide="generic-export" disabled={!canExportGeneric} onClick={exportGeneric}>
              {exporting ? <Loader2 className="spin" size={17} /> : <Download size={17} />}{templateExportEnabled ? "导出模板" : "导出"}
            </button>
          </div>
          <small className="simple-hint">{guidanceText}</small>
        </div>
        <details className="advanced-workspace" data-guide="generic-advanced" open={advancedOpen} onToggle={(event) => setAdvancedOpen(event.currentTarget.open)}>
          <summary>
            <span>选择主表、补充表和结果列</span>
            <small>主表决定行数，补充表按相同字段补资料</small>
          </summary>
        <div className="panel compose-panel" data-guide="generic-compose">
          <div className="panel-title"><Database size={18} /><span>本次结果怎么生成</span></div>
          <div className="plain-guide">
            <span><strong>主表</strong>决定输出多少行</span>
            <span><strong>补充表</strong>提供要补进来的字段</span>
            <span><strong>相同字段</strong>用来把两张表对上</span>
          </div>
          <div className="compose-section output-compose-section">
            <div className="compose-section-title">
              <span>结果要保留的列</span>
              <div className="mini-actions">
                <small>{outputColumns.length ? `${outputColumns.length} 列` : "默认全部列"}</small>
                <button className="mini-action" disabled={!activeBaseSource} onClick={() => activeBaseSource && appendOutputFields(activeBaseSource.id, activeBaseSource.headers)} title="加入主表全部字段">加入主表字段</button>
                <button className="mini-action icon-mini" disabled={!outputColumns.length} onClick={() => setOutputColumns([])} title="清空结果列"><Eraser size={14} /></button>
              </div>
            </div>
            <div
              className={`drop-zone output-drop ${activeDrop === "output" ? "active-drop" : ""} ${dragging?.kind === "field" ? "ready-drop" : ""}`}
              data-drop-target="output"
              onDragOver={allowDrop}
              onDragEnter={(event) => enterDrop(event, "output")}
              onDragLeave={() => setActiveDrop("")}
              onDrop={handleOutputDrop}
            >
              {outputColumns.length === 0 ? (
                <div className="canvas-empty">从右侧字段点“加入”，这里放的就是最终 Excel 的列；不选时会导出主表和已关联补充表的全部列。</div>
              ) : (
                <div className="output-column-list">
                  {outputColumns.map((column) => {
                    const source = column.sourceId ? sourceMap.get(column.sourceId) : undefined;
                    return (
                      <div className="output-column" key={column.id}>
                        <input
                          value={column.label}
                          onChange={(event) => setOutputColumns((current) => current.map((item) => item.id === column.id ? { ...item, label: event.target.value } : item))}
                        />
                        <small>{column.template != null ? `组合：${column.template}` : column.constant != null ? `"${column.constant}"` : `${sourceShortLabel(source)}.${column.column}`}</small>
                        <div className="output-actions">
                          <button className="icon-action ghost" disabled={outputColumns.indexOf(column) === 0} onClick={() => moveOutputColumn(outputColumns.indexOf(column), -1)} title="上移"><ArrowUp size={14} /></button>
                          <button className="icon-action ghost" disabled={outputColumns.indexOf(column) === outputColumns.length - 1} onClick={() => moveOutputColumn(outputColumns.indexOf(column), 1)} title="下移"><ArrowDown size={14} /></button>
                          <button className="icon-action ghost" onClick={() => setOutputColumns(outputColumns.filter((item) => item.id !== column.id))} title="删除结果列"><X size={15} /></button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <div className="constant-row">
              <input value={constantLabel} onChange={(event) => setConstantLabel(event.target.value)} placeholder="固定列名" />
              <input value={constantValue} onChange={(event) => setConstantValue(event.target.value)} placeholder="固定值" />
              <button className="icon-action" onClick={addConstantColumn} title="添加固定列"><Plus size={16} /></button>
            </div>
            <div className="template-editor">
              <label className="template-name-field">
                <span>组合列名</span>
                <input value={templateLabel} onChange={(event) => setTemplateLabel(event.target.value)} placeholder="例如 文件夹路径" />
              </label>
              <label className="template-content-field">
                <span>组合内容</span>
                <textarea
                  ref={templateTextareaRef}
                  value={templateValue}
                  onChange={(event) => setTemplateValue(event.target.value)}
                  placeholder="例如 文档文件/{功能区类别.装置代码}/{功能区类别.名称}/{单元类别.主项号}-{装置主项类别.名称}"
                />
              </label>
              <button className="template-add-button" onClick={addTemplateColumn}>
                <Plus size={16} />添加组合列
              </button>
            </div>
            <details className="advanced-config">
              <summary>文本配置</summary>
              <textarea value={outputText} onChange={(event) => setOutputText(event.target.value)} placeholder={'可选：批量补充固定列或组合列\n固定列 = "固定值"\n组合列 = {代码}/{名称}'} />
            </details>
            <div className="template-export-box">
              <label className="toggle-row">
                <input
                  type="checkbox"
                  checked={templateExportEnabled}
                  disabled={!excelSources.length}
                  onChange={(event) => {
                    const enabled = event.target.checked;
                    setTemplateExportEnabled(enabled);
                    if (!enabled) setTemplateExport({ sourceId: "" });
                  }}
                />
                  <span>按模板生成本次结果表</span>
              </label>
              {templateExportEnabled && (
                <div className="template-export-controls">
                  <select
                    value={templateExport.sourceId}
                    onChange={(event) => setTemplateExport({ sourceId: event.target.value })}
                  >
                    <option value="">选择模板 sheet</option>
                    {excelSources.map((source) => <option key={source.id} value={source.id}>{sourceLabel(source)}</option>)}
                  </select>
                  {templateSource && (
                    <small>
                      保留模板结构，只写入 {templateSource.sheetName} 第 {sourceDataStartRows[templateSource.id] ?? templateSource.defaultDataStartRow} 行起。
                    </small>
                  )}
                </div>
              )}
            </div>
          </div>
          <div
            className={`drop-zone base-drop ${activeBaseSource ? "filled" : ""} ${activeDrop === "base" ? "active-drop" : ""}`}
            data-drop-target="base"
            onDragOver={allowDrop}
            onDragEnter={(event) => enterDrop(event, "base")}
            onDragLeave={() => setActiveDrop("")}
            onDrop={(event) => {
              event.preventDefault();
              setBaseFromPayload(getDragPayload(event));
              finishDrag();
            }}
          >
            <span>主表（决定结果行数）</span>
            <strong>{activeBaseSource ? sourceLabel(activeBaseSource) : "选择一张主表"}</strong>
            {activeBaseSource && <small>{activeBaseSource.rowCount.toLocaleString("zh-CN")} 行 / {activeBaseSource.columnCount} 列</small>}
            <select value={activeBaseSourceId} onChange={(event) => setBaseSourceId(event.target.value)}>
              <option value="">选择主表</option>
              {sources.map((source) => <option key={source.id} value={source.id}>{sourceLabel(source)}</option>)}
            </select>
          </div>

          <div className="compose-section">
            <div className="compose-section-title">
              <span>补充表（按相同字段补资料）</span>
              <div className="mini-actions">
                <button className="icon-action" disabled={!activeBaseSourceId || sources.length < 2} onClick={addJoin} title="添加补充表">
                  <Plus size={16} />
                </button>
              </div>
            </div>
            {joins.length === 0 && <div className="canvas-empty">只导出主表时不用设置这里；需要从另一张表带字段时，添加补充表并选择两边相同含义的字段。</div>}
            {joins.map((join, index) => {
            const rightSource = sourceMap.get(join.sourceId);
            return (
              <div className="join-card" key={join.id}>
                <div className="join-card-head">
                  <span><Link2 size={15} />补充表 {index + 1}</span>
                  <button className="icon-action ghost" onClick={() => setJoins(joins.filter((item) => item.id !== join.id))} title="删除匹配">
                    <X size={15} />
                  </button>
                </div>
                <div
                  className={`drop-zone compact-drop ${activeDrop === `join-source-${join.id}` ? "active-drop" : ""}`}
                  data-drop-target={`join-source-${join.id}`}
                  onDragOver={allowDrop}
                  onDragEnter={(event) => enterDrop(event, `join-source-${join.id}`)}
                  onDragLeave={() => setActiveDrop("")}
                  onDrop={(event) => handleJoinSourceDrop(index, event)}
                >
                  <span>要补字段的表</span>
                  <strong>{rightSource ? sourceLabel(rightSource) : "选择一张补充表"}</strong>
                </div>
                <div className="key-drop-grid">
                  <div
                    className={`drop-zone key-drop ${join.leftKey ? "filled" : ""} ${activeDrop === `join-left-${join.id}` ? "active-drop" : ""}`}
                    data-drop-target={`join-left-${join.id}`}
                    onDragOver={allowDrop}
                    onDragEnter={(event) => enterDrop(event, `join-left-${join.id}`)}
                    onDragLeave={() => setActiveDrop("")}
                    onDrop={(event) => handleJoinKeyDrop(index, "left", event)}
                  >
                    <span>主表里的相同字段</span>
                    <strong>{join.leftKey || "选择主表字段"}</strong>
                  </div>
                  <div
                    className={`drop-zone key-drop ${join.rightKey ? "filled" : ""} ${activeDrop === `join-right-${join.id}` ? "active-drop" : ""}`}
                    data-drop-target={`join-right-${join.id}`}
                    onDragOver={allowDrop}
                    onDragEnter={(event) => enterDrop(event, `join-right-${join.id}`)}
                    onDragLeave={() => setActiveDrop("")}
                    onDrop={(event) => handleJoinKeyDrop(index, "right", event)}
                  >
                    <span>补充表里的相同字段</span>
                    <strong>{join.rightKey || "选择补充表字段"}</strong>
                  </div>
                </div>
                <div className="join-select-grid">
                  <select value={join.sourceId} onChange={(event) => {
                    const nextSource = sourceMap.get(event.target.value);
                    updateJoin(index, { ...join, sourceId: event.target.value, rightKey: "", prefix: nextSource?.sheetName ?? "" });
                  }}><option value="">选择补充表</option>{sources.filter((source) => source.rowCount > 0 && source.id !== activeBaseSourceId).map((source) => <option key={source.id} value={source.id}>{sourceLabel(source)}</option>)}</select>
                  <select value={join.leftKey} onChange={(event) => updateJoin(index, { ...join, leftKey: event.target.value })}><option value="">主表里的相同字段</option>{(activeBaseSource?.headers ?? []).map((header) => <option key={header} value={header}>{header}</option>)}</select>
                  <select value={join.rightKey} onChange={(event) => updateJoin(index, { ...join, rightKey: event.target.value })}><option value="">补充表里的相同字段</option>{(rightSource?.headers ?? []).map((header) => <option key={header} value={header}>{header}</option>)}</select>
                  <select value={join.joinType} onChange={(event) => updateJoin(index, { ...join, joinType: event.target.value as "left" | "inner" })}><option value="left">保留主表全部行（常用）</option><option value="inner">只保留两边都匹配的行</option></select>
                  <select
                    value={join.duplicateStrategy ?? "first"}
                    onChange={(event) => updateJoin(index, { ...join, duplicateStrategy: event.target.value as "first" | "expand" | "merge" })}
                    title="一条主表记录匹配到多条补充记录时怎么处理"
                  >
                    <option value="first">一对多时取第一条</option>
                    <option value="expand">一对多时展开成多行</option>
                    <option value="merge">一对多时合并到一格</option>
                  </select>
                  {(join.duplicateStrategy ?? "first") === "merge" && (
                    <input
                      className="join-separator-input"
                      value={join.multiValueSeparator ?? ","}
                      onChange={(event) => updateJoin(index, { ...join, multiValueSeparator: event.target.value })}
                      placeholder="合并分隔符"
                      title="多值合并分隔符"
                    />
                  )}
                </div>
              </div>
            );
          })}
          </div>

        </div>
        </details>
        <div className="panel guide-card">
          <div className="panel-title"><Info size={18} /><span>操作向导</span></div>
          <div className="guide-card-progress">
            <span>{guideCompletedCount}/{guideSteps.length}</span>
            <div><i style={{ width: `${Math.round((guideCompletedCount / guideSteps.length) * 100)}%` }} /></div>
          </div>
          <strong>{nextGuideStep.title}</strong>
          <small>{nextGuideStep.action}</small>
          <button className="secondary inline" onClick={() => startGenericGuide(nextGuideIndex >= 0 ? nextGuideIndex : 0)}>
            <Info size={16} />开始教程
          </button>
        </div>
      </section>
      <section className="right-pane" data-guide="generic-field-bank">
        <div className="workspace-header generic-workspace-header">
          <div>
            <p className="eyebrow">结果工作区</p>
            <h2>{rightView === "fields" ? "已解析的表和字段" : rightView === "result" ? "结果预览" : "问题清单"}</h2>
          </div>
          <div className="source-toolbar">
            {rightView === "fields" && <label className="search-field">
              <Search size={16} />
              <input value={fieldQuery} onChange={(event) => setFieldQuery(event.target.value)} placeholder="搜索字段名或样例值" />
            </label>}
            {preview && <span className={preview.stats.issueCount ? "badge warn" : "badge ok"}>{preview.stats.issueCount} 个问题</span>}
            <button className="view-action" disabled={!canPreviewGeneric} onClick={requestPreview} title="生成预览">
              {loading ? <Loader2 className="spin" size={15} /> : <CheckCircle2 size={15} />}预览
            </button>
            <button className="view-action" disabled={!canExportGeneric} onClick={exportGeneric} title="导出组合">
              {exporting ? <Loader2 className="spin" size={15} /> : <Download size={15} />}{templateExportEnabled ? "导出模板" : "导出"}
            </button>
            <button className="view-action" onClick={() => setFocusPreview((value) => !value)} title={focusPreview ? "显示配置栏" : "隐藏配置栏"}>
              {focusPreview ? <PanelLeftOpen size={15} /> : <PanelLeftClose size={15} />}{focusPreview ? "显示配置" : "专注预览"}
            </button>
          </div>
        </div>
        <div className="view-tabs">
          <button className={rightView === "fields" ? "active" : ""} onClick={() => setRightView("fields")}><Layers3 size={15} />表和字段</button>
          <button className={rightView === "result" ? "active" : ""} disabled={!preview} onClick={() => setRightView("result")}><Eye size={15} />结果</button>
          <button className={rightView === "issues" ? "active" : ""} disabled={!preview} onClick={() => setRightView("issues")}><AlertTriangle size={15} />问题</button>
        </div>
        {rightView === "fields" && (sources.length ? (
          <div className="source-list source-bank">
            {visibleSources().map((source) => {
              const headers = visibleHeaders(source);
              return (
              <div className={`source-item source-card ${source.id === activeBaseSourceId ? "active-source" : ""}`} key={source.id}>
                <div
                  className="source-card-title"
                  draggable
                  onDragStart={(event) => beginDrag(event, { kind: "source", sourceId: source.id })}
                  onDragEnd={finishDrag}
                  onPointerDown={(event) => beginPointerDrag(event, { kind: "source", sourceId: source.id })}
                  onMouseDown={(event) => beginMouseDrag(event, { kind: "source", sourceId: source.id })}
                >
                  <div className="source-card-name">
                    <strong>{sourceLabel(source)}</strong>
                  </div>
                  <div className="source-card-actions">
                    <button onClick={() => appendOutputFields(source.id, headers)} disabled={!headers.length}>加入结果列</button>
                    <button onClick={() => setBaseSourceId(source.id)}>设为主表</button>
	                    <button onClick={() => exportSourceSheet(source.id)} disabled={!files.length || exporting || !canOutputSource(source)}>
	                      {exporting ? "导出中" : "单独导出"}
	                    </button>
                  </div>
                </div>
                <div className="source-meta">
                  <span>{sourceRoleLabel(source)}</span>
                  {source.sheetState !== "visible" && <span>{source.sheetState}</span>}
                  <small>{source.rowCount.toLocaleString("zh-CN")} 行 / {source.columnCount} 列</small>
                </div>
                <details className="source-options">
                  <summary>表格识别设置</summary>
                  <div className="source-options-grid">
                    <label className="compact-field">
                      <span>表头行</span>
                      <input
                        type="number"
                        min={1}
                        value={sourceHeaderRows[source.id] ?? source.defaultHeaderRow}
                        onChange={(event) => setSourceHeaderRows({ ...sourceHeaderRows, [source.id]: Number(event.target.value) || 1 })}
                      />
                    </label>
                    <label className="compact-field">
                      <span>数据行</span>
                      <input
                        type="number"
                        min={1}
                        value={sourceDataStartRows[source.id] ?? source.defaultDataStartRow}
                        onChange={(event) => setSourceDataStartRows({ ...sourceDataStartRows, [source.id]: Number(event.target.value) || 1 })}
                      />
                    </label>
                    <label className="compact-field filldown-field">
                      <span>补齐列</span>
                      <input
                        value={sourceFillDownColumns[source.id] ?? ""}
                        onChange={(event) => setSourceFillDownColumns({ ...sourceFillDownColumns, [source.id]: event.target.value })}
                        placeholder="装置,装置名称"
                      />
                    </label>
                  </div>
                </details>
                <small className="source-origin">{source.kind === "derived" ? "生成表" : "原始表"}</small>
                <div className="field-chip-list">
                  {headers.map((header) => {
                    const sample = source.previewRows.find((row) => row[header])?.[header] ?? "";
                    const isDragging = dragging?.kind === "field" && dragging.sourceId === source.id && dragging.column === header;
                    return (
                      <div
                        className={`field-chip ${isDragging ? "dragging" : ""}`}
                        key={`${source.id}-${header}`}
                        role="button"
                        tabIndex={0}
                        onKeyDown={(event) => {
                          if (event.key === "Enter" || event.key === " ") {
                            event.preventDefault();
                            appendOutputField(source.id, header);
                          }
                        }}
                        onPointerDown={(event) => beginPointerDrag(event, { kind: "field", sourceId: source.id, column: header })}
                        onMouseDown={(event) => beginMouseDrag(event, { kind: "field", sourceId: source.id, column: header })}
                        title={`${sourceLabel(source)} / ${header}${sample ? ` / ${sample}` : ""}`}
                      >
                        <GripVertical size={14} className="drag-grip" />
                        <span>{header}</span>
                        {sample && <small>{sample}</small>}
                        <div className="chip-actions">
                          <button
                            type="button"
                            className="chip-add"
                            onClick={(event) => {
                              event.stopPropagation();
                              appendOutputField(source.id, header);
                            }}
                            title="加入结果列"
                          >
                            加入
                          </button>
                          <button
                            type="button"
                            className="chip-add"
                            onClick={(event) => {
                              event.stopPropagation();
                              appendFieldToTemplate(source.id, header);
                            }}
                            title="插入到组合内容"
                          >
                            插入
                          </button>
                        </div>
                      </div>
                    );
                  })}
                  {!headers.length && <div className="no-fields">没有可显示字段</div>}
                </div>
              </div>
            );})}
          </div>
	        ) : <div className="empty-state"><FileSpreadsheet size={38} /><h2>等待解析</h2><p>上传 Excel 或 PDF 后点“解析”，右侧会列出可选的表和字段。</p></div>)}
        {rightView === "result" && (preview ? (
          <>
            <div className="stats-grid generic-stats"><Stat label="可用表" value={preview.stats.sourceCount} /><Stat label="输出行" value={preview.stats.outputRowCount} /><Stat label="问题" value={preview.stats.issueCount} /></div>
            <div className="result-title"><h2>组合结果</h2></div>
            {previewHeaders.length ? (
              <div className="table-shell generic-table"><table><thead><tr>{previewHeaders.map((header) => <th key={header}>{header}</th>)}</tr></thead><tbody>{previewRows.map((row, rowIndex) => <tr key={rowIndex}>{previewHeaders.map((header) => <td key={header}>{row[header]}</td>)}</tr>)}</tbody></table></div>
            ) : (
              <div className="empty-state inline-empty"><FileSpreadsheet size={34} /><h2>没有输出行</h2><p>请检查主表、补充表关联或结果列配置。</p></div>
            )}
          </>
        ) : <div className="empty-state"><Eye size={38} /><h2>等待预览</h2><p>生成预览后会显示组合结果的前 100 行。</p></div>)}
        {rightView === "issues" && (preview ? (
          <div className="table-shell generic-table issue-table"><table><thead><tr><th>类型</th><th>级别</th><th>说明</th><th>来源</th><th>关键值</th></tr></thead><tbody>{issueRows.length ? issueRows.map((row, index) => <tr key={`${row.type}-${index}`}><td>{row.type}</td><td>{row.severity}</td><td>{row.message}</td><td>{row.source || ""}{row.row ? `:${row.row}` : ""}</td><td>{row.key || row.value || ""}</td></tr>) : <tr><td colSpan={5}>未发现问题</td></tr>}</tbody></table></div>
        ) : <div className="empty-state"><AlertTriangle size={38} /><h2>等待问题清单</h2><p>生成预览后会列出未匹配、重复键和配置异常。</p></div>)}
      </section>
    </main>
  );
}

function App() {
  const [accessToken, setAccessToken] = React.useState(() => {
    try {
      return window.localStorage.getItem(accessTokenStorageKey) ?? "";
    } catch {
      return "";
    }
  });
  const [showPreset, setShowPreset] = React.useState(false);

  React.useEffect(() => {
    try {
      if (accessToken) {
        window.localStorage.setItem(accessTokenStorageKey, accessToken);
      } else {
        window.localStorage.removeItem(accessTokenStorageKey);
      }
    } catch {
      // localStorage may be unavailable in private or restricted browser contexts.
    }
  }, [accessToken]);

  return (
    <>
      <GenericWorkspace accessToken={accessToken} setAccessToken={setAccessToken} onOpenPreset={() => setShowPreset(true)} />
      <AnimatePresence>
        {showPreset && (
          <motion.div
            className="preset-overlay"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.18 }}
          >
            <PresetWorkspace accessToken={accessToken} setAccessToken={setAccessToken} onClose={() => setShowPreset(false)} />
          </motion.div>
        )}
      </AnimatePresence>
    </>
  );
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
