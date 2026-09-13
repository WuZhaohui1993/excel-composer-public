import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import type { GenericCombineConfig, GenericFileInput } from "../core/genericCombiner.js";
import { combineGeneric, parseGenericSources, writeGenericWorkbook } from "../core/genericCombiner.js";

interface GenericUpload {
  files: GenericFileInput[];
  config: GenericCombineConfig;
}

async function readGenericUpload(request: FastifyRequest): Promise<GenericUpload> {
  const files: GenericFileInput[] = [];
  let config: GenericCombineConfig = {};

  for await (const part of request.parts()) {
    if (part.type === "file") {
      files.push({
        fileName: part.filename ?? "未命名文件",
        mimeType: part.mimetype,
        buffer: await part.toBuffer()
      });
    } else if (part.fieldname === "config") {
      try {
        config = JSON.parse(String(part.value));
      } catch {
        config = {};
      }
    }
  }

  return { files, config };
}

function requireFiles(upload: GenericUpload, reply: FastifyReply): boolean {
  if (!upload.files.length) {
    reply.status(400).send({ error: "missing_files", message: "请至少上传一个 .xlsx 或 .pdf 文件" });
    return false;
  }
  return true;
}

export async function registerGenericRoutes(server: FastifyInstance): Promise<void> {
  server.post("/api/generic/inspect", async (request, reply) => {
    const upload = await readGenericUpload(request);
    if (!requireFiles(upload, reply)) return;
    const tables = await parseGenericSources(upload.files, upload.config);
    return { sources: tables.map((table) => table.source) };
  });

  server.post("/api/generic/preview", async (request, reply) => {
    const upload = await readGenericUpload(request);
    if (!requireFiles(upload, reply)) return;
    return combineGeneric(upload.files, upload.config);
  });

  server.post("/api/generic/export", async (request, reply) => {
    const upload = await readGenericUpload(request);
    if (!requireFiles(upload, reply)) return;
    const output = await writeGenericWorkbook(upload.files, upload.config);
    const fileName = encodeURIComponent(`通用组合结果_${new Date().toISOString().slice(0, 10)}.xlsx`);
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename*=UTF-8''${fileName}`)
      .send(output);
  });
}
