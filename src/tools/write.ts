/**
 * Write tools. Registered ONLY when ENABLE_WRITES=true. Each requires confirm=true;
 * otherwise it returns a dry-run preview and performs no change.
 *
 *   onenote_create_page(sectionId, title, htmlBody)  → create_page
 *   onenote_append_to_page(pageId, htmlBody)          → append_to_page
 *
 * These are the two primary OneNote write operations. Page content is HTML.
 *
 * SURGICAL EDITS (not shipped as a tool): to modify a *specific* existing element rather
 * than appending to the whole body, PATCH /pages/{id}/content with a command whose
 * `target` is that element's `data-id`. Read the page first with includeIds=true
 * (onenote_read_page) to obtain those data-id anchors. See onenoteClient.ts.
 */

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { buildPageHtml, type OneNoteClient } from "../onenoteClient.js";
import { ok, runTool } from "./helpers.js";
import { CONFIRM_FIELD_DESCRIPTION, previewResult, appliedResult } from "../confirmation.js";
import { audit } from "../logging.js";

const CREATE = { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true } as const;

/** Trim an HTML snippet for preview so we never echo an entire document into logs/output. */
function preview(html: string, max = 500): string {
  const s = html.trim();
  return s.length > max ? s.slice(0, max) + `… (${s.length} chars total)` : s;
}

export function registerWriteTools(server: McpServer, client: OneNoteClient): void {
  // create_page(sectionId, title, htmlBody)
  server.registerTool(
    "onenote_create_page",
    {
      title: "Create a OneNote page",
      description:
        "Create a new page in a section. The title becomes the page's <title>; htmlBody is the page's HTML content " +
        "(an HTML fragment — headings, paragraphs, lists, tables, etc.). Requires confirm=true to actually create; " +
        "otherwise returns a preview. Only create pages the user has explicitly approved.",
      inputSchema: {
        sectionId: z.string().min(1).describe("The section id to create the page in (from onenote_list_sections)."),
        title: z.string().min(1).describe("The page title."),
        htmlBody: z
          .string()
          .min(1)
          .describe("The page body as an HTML fragment, e.g. '<h1>Notes</h1><p>First line.</p>'."),
        confirm: z.boolean().optional().describe(CONFIRM_FIELD_DESCRIPTION),
      },
      annotations: CREATE,
    },
    async ({ sectionId, title, htmlBody, confirm }) =>
      runTool("onenote_create_page", async (requestId) => {
        const target = { sectionId };
        if (!confirm) {
          audit({ tool: "onenote_create_page", outcome: "preview", recordType: "page", confirmed: false });
          return previewResult("create_page", target, { title, htmlBody: preview(htmlBody) });
        }
        const fullHtml = buildPageHtml(title, htmlBody);
        const created = await client.createPage(sectionId, fullHtml, requestId);
        audit({
          tool: "onenote_create_page",
          outcome: "success",
          recordType: "page",
          recordId: created?.id,
          confirmed: true,
        });
        return appliedResult("create_page", target, {
          id: created?.id ?? null,
          title: (created?.title as string) ?? title,
          links: created?.links ?? null,
        });
      }),
  );

  // append_to_page(pageId, htmlBody)
  server.registerTool(
    "onenote_append_to_page",
    {
      title: "Append a block to a OneNote page",
      description:
        "Append an HTML block to the END of an existing page's body. htmlBody is an HTML fragment " +
        "(e.g. '<p>Follow-up: …</p>'). This does NOT surgically edit existing content — it only appends. " +
        "Requires confirm=true to actually append; otherwise returns a preview.",
      inputSchema: {
        pageId: z.string().min(1).describe("The page id to append to (from onenote_list_pages)."),
        htmlBody: z.string().min(1).describe("The HTML fragment to append at the end of the page body."),
        confirm: z.boolean().optional().describe(CONFIRM_FIELD_DESCRIPTION),
      },
      annotations: CREATE,
    },
    async ({ pageId, htmlBody, confirm }) =>
      runTool("onenote_append_to_page", async (requestId) => {
        const target = { pageId };
        if (!confirm) {
          audit({ tool: "onenote_append_to_page", outcome: "preview", recordType: "page", recordId: pageId, confirmed: false });
          return previewResult("append_to_page", target, { htmlBody: preview(htmlBody) });
        }
        await client.appendToPage(pageId, htmlBody, requestId);
        audit({ tool: "onenote_append_to_page", outcome: "success", recordType: "page", recordId: pageId, confirmed: true });
        return appliedResult("append_to_page", target, { appended: true });
      }),
  );
}
