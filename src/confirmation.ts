/**
 * Explicit-confirmation gate for write operations.
 *
 * Every write tool takes a `confirm` boolean (default false). When false, the tool
 * returns a structured PREVIEW of exactly what it would change and performs NO write.
 * The caller (the model, on the user's explicit instruction) must re-invoke with
 * `confirm: true` to actually perform the change. This is layered on top of the
 * server-level ENABLE_WRITES flag (writes are not even registered when it is off).
 */

import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";

export const CONFIRM_FIELD_DESCRIPTION =
  "Must be set to true to actually perform this change. When false or omitted, the tool " +
  "returns a preview of the intended change and writes nothing. Only set true after the " +
  "user has explicitly approved the exact change shown in the preview.";

/** Build the standard "preview only, not yet performed" result. */
export function previewResult(action: string, target: Record<string, unknown>, payload: Record<string, unknown>): CallToolResult {
  const structured = {
    status: "confirmation_required" as const,
    action,
    target,
    proposed_change: payload,
    message:
      "No change was made. This is a preview. To apply it, call this tool again with confirm=true. " +
      "Confirm the details above with the user before doing so.",
  };
  return {
    content: [{ type: "text", text: JSON.stringify(structured, null, 2) }],
    structuredContent: structured,
  };
}

/** Build the standard "change applied" result. */
export function appliedResult(action: string, target: Record<string, unknown>, result: unknown): CallToolResult {
  const structured = {
    status: "applied" as const,
    action,
    target,
    result,
  };
  return {
    content: [{ type: "text", text: JSON.stringify(structured, null, 2) }],
    structuredContent: structured,
  };
}
