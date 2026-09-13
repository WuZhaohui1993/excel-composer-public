import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import type { GenerateConfig, PreviewResult } from "../core/types.js";
import { generatePreset } from "../core/generator.js";
import { writePresetWorkbook } from "../core/workbookWriter.js";

interface UploadedSources {
  matrix?: Buffer;
  pdf?: Buffer;
  config: GenerateConfig;
}

async function readSources(request: FastifyRequest): Promise<UploadedSources> {
  const parts = request.parts();
  const result: UploadedSources = { config: {} };

  for await (const part of parts) {
    if (part.type === "file") {
      const buffer = await part.toBuffer();
      const filename = part.filename ?? "";
      if (/\.pdf$/i.test(filename) || part.mimetype === "application/pdf" || part.fieldname === "pdf") {
        result.pdf = buffer;
      } else if (/\.xlsx$/i.test(filename) || part.fieldname === "matrix") {
        result.matrix = buffer;
      }
    } else if (part.fieldname === "config") {
      try {
        result.config = JSON.parse(String(part.value));
      } catch {
        result.config = {};
      }
    }
  }

  return result;
}

function requireSources(sources: UploadedSources, reply: FastifyReply): sources is Required<Pick<UploadedSources, "matrix" | "pdf">> & UploadedSources {
  if (!sources.matrix || !sources.pdf) {
    reply.status(400).send({
      error: "missing_sources",
      message: "请同时上传文档矩阵 Excel 和项目主项表 PDF"
    });
    return false;
  }
  return true;
}

function toPreview(result: Awaited<ReturnType<typeof generatePreset>>, limit: number): PreviewResult {
  return {
    projectPreview: result.project.devices.slice(0, limit),
    folderPreview: result.folderRows.slice(0, limit),
    reviewPreview: result.reviewRows.slice(0, limit),
    issues: result.issues.slice(0, Math.max(200, limit)),
    stats: {
      ...result.stats,
      issueCount: result.issues.length
    }
  };
}

export async function registerPresetRoutes(server: FastifyInstance): Promise<void> {
  server.post("/api/preset/preview", async (request, reply) => {
    const sources = await readSources(request);
    if (!requireSources(sources, reply)) return;
    const result = await generatePreset(sources.matrix, sources.pdf, sources.config);
    return toPreview(result, sources.config.previewLimit ?? 100);
  });

  server.post("/api/preset/export", async (request, reply) => {
    const sources = await readSources(request);
    if (!requireSources(sources, reply)) return;
    const result = await generatePreset(sources.matrix, sources.pdf, sources.config);
    const output = await writePresetWorkbook(result);
    const fileName = encodeURIComponent(`示例组合结果_${new Date().toISOString().slice(0, 10)}.xlsx`);
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename*=UTF-8''${fileName}`)
      .send(output);
  });
}
