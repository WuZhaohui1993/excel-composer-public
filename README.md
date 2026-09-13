# Excel 表格工具

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE) [![React](https://img.shields.io/badge/React-18.x-61dafb.svg)](https://react.dev/) [![TypeScript](https://img.shields.io/badge/TypeScript-5.x-3178c6.svg)](https://www.typescriptlang.org/)

浏览器端 Excel/PDF 解析和组合工具。上传多个来源后，可以识别表格、配置字段匹配、预览结果并导出新的工作簿。

> 本项目提供通用表格处理能力，不包含任何客户文件、项目名称、真实业务数据、访问口令、服务器配置或 SSH 材料。示例文件应使用虚构内容。

## 功能

- **多文件解析**：识别 Excel 工作表和 PDF 文本行，展示字段与数据概览。
- **字段映射**：选择主表、补充表、匹配字段、保留规则和输出列。
- **结果预览**：预览组合结果、问题清单和数据源概览。
- **模板写回**：在受控模板模式下保留说明页、隐藏字典、列宽和基础样式。
- **导出与校验**：生成新的 `.xlsx` 文件，并对缺失字段、重复匹配和异常行给出反馈。

## 技术栈

| 层 | 技术 |
| --- | --- |
| 前端 | React 18、TypeScript、Vite、Framer Motion |
| 后端 | Node.js、Fastify、Multipart、ExcelJS、PDF.js |
| 测试 | Vitest、TypeScript |

## 快速开始

### 环境要求

Git、Node.js 20+ 和 npm。

### 安装与启动

```bash
npm install
npm run dev
```

前端默认运行在 `http://127.0.0.1:5173`，本地 API 默认运行在 `http://127.0.0.1:3001`。如需访问口令，在本地 `.env` 中设置 `ACCESS_TOKEN`，不要提交该文件。

## 测试

```bash
npm run type-check
npm test
npm run build
```

当前测试使用内存生成的虚构工作簿；不会读取客户文件或生产目录。

## 项目结构

```text
├── src/            # React 前端与交互界面
├── server/         # Fastify API、解析器和工作簿写回
├── tests/          # 通用组合与规范化测试
├── scripts/        # 本地辅助脚本
├── LICENSE
└── THIRD_PARTY_NOTICES.md
```

## 安全边界

上传文件只应在本地或受控服务中处理。生产部署必须设置访问口令、限制文件大小和类型、隔离临时目录、清理导出文件，并使用 HTTPS 和反向代理。

## 参与贡献

请使用虚构 Excel/PDF 作为测试输入，说明字段规则和预期结果；不要提交客户文件、真实数据、访问口令、密钥、服务器信息或生成物。

## 许可证

本项目及其自有代码采用 MIT 许可证。React、Fastify、ExcelJS、PDF.js 和其他依赖的许可证见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) 及锁文件。
