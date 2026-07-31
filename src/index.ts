/**
 * HTTP host: serves the MCP endpoint over Streamable HTTP (stateless JSON mode),
 * protected by a bearer token. Exposes an unauthenticated /healthz for probes.
 */

import express from "express";
import { randomUUID } from "node:crypto";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { loadConfig } from "./config.js";
import { setLogLevel, log } from "./logging.js";
import { buildServer } from "./server.js";
import { bearerAuth } from "./auth.js";

async function main(): Promise<void> {
  const config = loadConfig();
  setLogLevel(config.logLevel);

  const app = express();
  app.disable("x-powered-by");
  app.use(express.json({ limit: "2mb" }));

  // Liveness/readiness probe — no auth, no data.
  app.get("/healthz", (_req, res) => {
    res.json({
      status: "ok",
      service: "onenote-mcp",
      version: "err-detail-2",
      authMode: config.graph.authMode,
      writesEnabled: config.enableWrites,
    });
  });

  // MCP endpoint (stateless: a fresh server+transport per request).
  app.post(config.mcpPath, bearerAuth(config.mcpAuthToken), async (req, res) => {
    const server = buildServer(config);
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    res.on("close", () => {
      transport.close();
      server.close();
    });
    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (err) {
      log.error("mcp.request_failed", { requestId: randomUUID() });
      if (!res.headersSent) {
        res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal server error" }, id: null });
      }
    }
  });

  // Streamable HTTP uses POST; reject GET/DELETE on the MCP path cleanly (stateless server).
  const methodNotAllowed = (_req: express.Request, res: express.Response) =>
    res.status(405).json({ jsonrpc: "2.0", error: { code: -32000, message: "Method not allowed." }, id: null });
  app.get(config.mcpPath, methodNotAllowed);
  app.delete(config.mcpPath, methodNotAllowed);

  app.listen(config.port, () => {
    log.info("server.listening", { event: "server.listening" });
    // Intentionally do not log the token, credentials, or target ids.
    console.log(
      JSON.stringify({
        level: "info",
        event: "server.started",
        service: "onenote-mcp",
        port: config.port,
        mcpPath: config.mcpPath,
        authMode: config.graph.authMode,
        writesEnabled: config.enableWrites,
      }),
    );
  });
}

main().catch((err) => {
  // Print only the message (which is authored by us / config.ts) — never a stack with env values.
  console.error(JSON.stringify({ level: "error", event: "server.fatal", message: (err as Error).message }));
  process.exit(1);
});
