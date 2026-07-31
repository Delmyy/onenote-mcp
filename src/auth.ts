/**
 * Edge authentication for MCP client → this server.
 *
 * The MCP client (claude.ai connector, `claude mcp`, ChatGPT, MCP Inspector, ...) must
 * present the shared secret as a bearer token:
 *   Authorization: Bearer <MCP_AUTH_TOKEN>
 *
 * This is the least-privilege network gate in front of the MCP endpoint. The upstream
 * Microsoft Graph credentials never leave the server.
 */

import type { Request, Response, NextFunction } from "express";
import { log } from "./logging.js";

/** Constant-time string comparison to avoid timing side channels. */
function safeEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

export function bearerAuth(expectedToken: string) {
  return (req: Request, res: Response, next: NextFunction): void => {
    const header = req.header("authorization") ?? "";
    const match = /^Bearer\s+(.+)$/i.exec(header);
    if (!match || !safeEqual(match[1], expectedToken)) {
      log.warn("auth.denied", { event: "auth.denied" });
      res
        .status(401)
        .set("WWW-Authenticate", 'Bearer realm="onenote-mcp", error="invalid_token"')
        .json({ error: "unauthorized", message: "A valid bearer token is required." });
      return;
    }
    next();
  };
}
