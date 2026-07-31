/**
 * Shared helpers for tool implementations: safe result formatting, untrusted-content
 * wrapping, and a wrapper that turns thrown HttpErrors into safe tool errors.
 */

import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { HttpError } from "../http.js";
import { log, newRequestId } from "../logging.js";
import { wrapUntrusted, wrapUntrustedFields } from "../security.js";
import type { OneNoteRecord } from "../onenoteClient.js";

/** Free-text fields on OneNote records that may carry authored content → treat as untrusted. */
const UNTRUSTED_FIELDS = ["title", "displayName", "content", "html", "preview"];

export function sanitizeRecord(rec: OneNoteRecord): OneNoteRecord {
  return wrapUntrustedFields(rec, UNTRUSTED_FIELDS);
}

export function sanitizeRecords(recs: OneNoteRecord[]): OneNoteRecord[] {
  return recs.map(sanitizeRecord);
}

/** Wrap a whole page-HTML string as untrusted content. */
export function sanitizeHtml(html: string): string {
  return wrapUntrusted(html);
}

/** Wrap a structured object into the standard text+structured tool result. */
export function ok(structured: Record<string, unknown>): CallToolResult {
  return {
    content: [{ type: "text", text: JSON.stringify(structured, null, 2) }],
    structuredContent: structured,
  };
}

/** Standard error result (safe, content-free message). */
export function errorResult(message: string): CallToolResult {
  return {
    content: [{ type: "text", text: JSON.stringify({ status: "error", message }, null, 2) }],
    structuredContent: { status: "error", message },
    isError: true,
  };
}

/**
 * Run a read/write tool body with a fresh request id, uniform logging, and safe
 * error conversion. Never leaks upstream bodies or content.
 */
export async function runTool(
  toolName: string,
  fn: (requestId: string) => Promise<CallToolResult>,
): Promise<CallToolResult> {
  const requestId = newRequestId();
  const started = Date.now();
  try {
    const result = await fn(requestId);
    log.info("tool.ok", { tool: toolName, requestId, latencyMs: Date.now() - started });
    return result;
  } catch (err) {
    const latencyMs = Date.now() - started;
    if (err instanceof HttpError) {
      log.warn("tool.http_error", { tool: toolName, requestId, status: err.status, latencyMs });
      return errorResult(err.safeMessage);
    }
    log.error("tool.error", { tool: toolName, requestId, latencyMs });
    return errorResult(
      "The tool encountered an unexpected error. See server logs (no content) for the request id: " + requestId,
    );
  }
}
