/**
 * Privacy-preserving logger + audit trail.
 *
 * RULES (enforced by only ever logging the fields below):
 *  - NEVER log request/response bodies, tool arguments, notebook/section/page titles,
 *    page HTML, access tokens, client secrets, or refresh tokens.
 *  - DO log: event name, tool name, outcome, HTTP status, latency, a random request id,
 *    and — for the audit trail only — the record TYPE and id a write targeted.
 */

type Level = "error" | "warn" | "info" | "debug";
const ORDER: Record<Level, number> = { error: 0, warn: 1, info: 2, debug: 3 };

let threshold: Level = "info";
export function setLogLevel(l: Level): void {
  threshold = l;
}

function emit(level: Level, event: string, fields: Record<string, unknown>): void {
  if (ORDER[level] > ORDER[threshold]) return;
  // Whitelist of safe field keys. Anything else is dropped defensively.
  const SAFE = new Set([
    "event",
    "tool",
    "outcome",
    "status",
    "latencyMs",
    "requestId",
    "recordType",
    "recordId",
    "confirmed",
    "writesEnabled",
    "reason",
    "attempt",
    "authMode",
  ]);
  const safe: Record<string, unknown> = { level, event, ts: new Date().toISOString() };
  for (const [k, v] of Object.entries(fields)) {
    if (SAFE.has(k)) safe[k] = v;
  }
  const line = JSON.stringify(safe);
  if (level === "error") console.error(line);
  else if (level === "warn") console.warn(line);
  else console.log(line);
}

export const log = {
  error: (event: string, f: Record<string, unknown> = {}) => emit("error", event, f),
  warn: (event: string, f: Record<string, unknown> = {}) => emit("warn", event, f),
  info: (event: string, f: Record<string, unknown> = {}) => emit("info", event, f),
  debug: (event: string, f: Record<string, unknown> = {}) => emit("debug", event, f),
};

/**
 * Audit event for every write attempt/outcome. Records only non-content identifiers.
 * Emitted at info level so it is retained by default.
 */
export function audit(entry: {
  tool: string;
  outcome: "preview" | "success" | "denied" | "error";
  recordType?: string;
  recordId?: string;
  confirmed: boolean;
  reason?: string;
}): void {
  emit("info", "audit.write", { ...entry, writesEnabled: true });
}

/** 16 hex chars, request correlation id. */
export function newRequestId(): string {
  const b = new Uint8Array(8);
  globalThis.crypto.getRandomValues(b);
  return Array.from(b, (x) => x.toString(16).padStart(2, "0")).join("");
}
