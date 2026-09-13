import path from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";
import fastifyMultipart from "@fastify/multipart";
import fastifyStatic from "@fastify/static";
import fastifyCors from "@fastify/cors";
import Fastify from "fastify";
import { registerPresetRoutes } from "./routes/preset.js";
import { registerGenericRoutes } from "./routes/generic.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function tokenGuard(expectedToken?: string) {
  return async (request: any, reply: any) => {
    if (!expectedToken) return;
    if (request.url === "/api/health") return;
    if (!request.url.startsWith("/api/")) return;
    const token = request.headers["x-access-token"];
    if (token !== expectedToken) {
      reply.status(401).send({ error: "unauthorized", message: "访问口令不正确" });
    }
  };
}

export async function buildServer() {
  const server = Fastify({
    logger: {
      level: process.env.LOG_LEVEL ?? "info"
    },
    bodyLimit: Number(process.env.MAX_BODY_BYTES ?? 200 * 1024 * 1024)
  });

  await server.register(fastifyCors, {
    origin: true,
    credentials: true
  });
  await server.register(fastifyMultipart, {
    limits: {
      fileSize: Number(process.env.MAX_FILE_BYTES ?? 50 * 1024 * 1024),
      files: Number(process.env.MAX_FILES ?? 20)
    }
  });

  server.addHook("preHandler", tokenGuard(process.env.ACCESS_TOKEN));
  server.get("/api/health", async () => ({
    ok: true,
    service: "excel-composer",
    tokenRequired: Boolean(process.env.ACCESS_TOKEN)
  }));

  await registerPresetRoutes(server);
  await registerGenericRoutes(server);

  const publicDir = path.resolve(__dirname, "../../../dist");
  if (fs.existsSync(publicDir)) {
    await server.register(fastifyStatic, {
      root: publicDir,
      prefix: "/"
    });
  }
  server.setNotFoundHandler((request, reply) => {
    if (request.url.startsWith("/api/")) {
      reply.status(404).send({ error: "not_found" });
      return;
    }
    if (fs.existsSync(publicDir)) {
      reply.sendFile("index.html");
      return;
    }
    reply.status(404).send({ error: "frontend_not_built", message: "前端尚未构建，请先运行 npm run build" });
  });

  return server;
}
