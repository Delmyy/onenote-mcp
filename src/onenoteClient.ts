/**
 * Microsoft Graph OneNote API client (v1.0).
 *
 * Key facts encoded here (see https://learn.microsoft.com/graph/api/resources/onenote):
 *   - Base:   https://graph.microsoft.com/v1.0
 *   - Auth:   Authorization: Bearer <graph access token>   (from GraphAuthProvider)
 *   - Scope:  application  Notes.ReadWrite.All   OR   delegated  Notes.ReadWrite
 *   - Root:   OneNote lives under a user/group/site context:
 *               delegated + me  →  /me/onenote
 *               user            →  /users/{id}/onenote
 *               group           →  /groups/{id}/onenote
 *               site            →  /sites/{id}/onenote
 *             App-only (client credentials) has NO user context, so "/me" is invalid —
 *             application mode must target a specific user/group/site.
 *   - Content model: page bodies are **HTML**. The two primary write operations are
 *       • create page:  POST .../sections/{id}/pages   (Content-Type: text/html)
 *       • append block: PATCH .../pages/{id}/content   (JSON command array)
 *
 *   NOTE ON SURGICAL EDITS: append is a coarse whole-body operation. To edit a *specific*
 *   existing element (replace/insert next to one paragraph, image, etc.) you must PATCH
 *   `/pages/{id}/content` with a command whose `target` is the element's `data-id`. Those
 *   data-id values are only present if you read the page with `?includeIDs=true` first
 *   (see getPageContent). We deliberately expose only whole-body append here; targeted
 *   PATCH-by-data-id is a documented extension point, not a shipped tool.
 */

import type { HttpClient } from "./http.js";
import type { GraphAuthProvider } from "./graphAuth.js";
import type { TargetType } from "./config.js";

export interface OneNoteRecord {
  id: string;
  [attr: string]: unknown;
}

export interface ListResult {
  records: OneNoteRecord[];
  /** Graph's @odata.nextLink for pagination, if present. */
  nextLink?: string;
}

export interface PageContent {
  pageId: string;
  html: string;
}

/** Escape a string for safe interpolation into HTML text/attribute context. */
function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** Build a complete OneNote-compatible HTML document. Graph derives the title from <title>. */
export function buildPageHtml(title: string, htmlBody: string): string {
  return (
    `<!DOCTYPE html>\n<html lang="en">\n<head>\n` +
    `  <meta charset="utf-8" />\n` +
    `  <title>${escapeHtml(title)}</title>\n` +
    `</head>\n<body>\n${htmlBody}\n</body>\n</html>`
  );
}

export class OneNoteClient {
  private root: string;

  constructor(
    private http: HttpClient,
    private auth: GraphAuthProvider,
    apiBase: string,
    targetType: TargetType,
    targetId: string,
  ) {
    this.root = `${apiBase}${OneNoteClient.resourcePrefix(targetType, targetId)}/onenote`;
  }

  /** Compute the Graph resource prefix for the configured target. */
  private static resourcePrefix(targetType: TargetType, targetId: string): string {
    switch (targetType) {
      case "me":
        return "/me";
      case "user":
        return `/users/${encodeURIComponent(targetId)}`;
      case "group":
        return `/groups/${encodeURIComponent(targetId)}`;
      case "site":
        return `/sites/${encodeURIComponent(targetId)}`;
    }
  }

  private async authHeaders(json = false, requestId?: string): Promise<Record<string, string>> {
    const token = await this.auth.getToken(requestId);
    const h: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    };
    if (json) h["Content-Type"] = "application/json";
    return h;
  }

  private async getList(path: string, query: Record<string, string>, requestId?: string): Promise<ListResult> {
    const qs = new URLSearchParams(query).toString();
    const url = `${this.root}${path}${qs ? `?${qs}` : ""}`;
    const { data } = await this.http.requestJson<{ value?: unknown[]; "@odata.nextLink"?: string }>(url, {
      headers: await this.authHeaders(false, requestId),
      requestId,
    });
    const records = Array.isArray(data?.value) ? (data!.value as OneNoteRecord[]) : [];
    return { records, nextLink: data?.["@odata.nextLink"] };
  }

  private async getOne(path: string, query: Record<string, string>, requestId?: string): Promise<OneNoteRecord | null> {
    const qs = new URLSearchParams(query).toString();
    const url = `${this.root}${path}${qs ? `?${qs}` : ""}`;
    const { data } = await this.http.requestJson<OneNoteRecord>(url, {
      headers: await this.authHeaders(false, requestId),
      requestId,
    });
    return data ?? null;
  }

  // ── Reads ────────────────────────────────────────────────────────────────────

  /** List notebooks in the configured user/group/site. */
  listNotebooks(requestId?: string): Promise<ListResult> {
    return this.getList(
      "/notebooks",
      { $select: "id,displayName,isDefault,createdDateTime,lastModifiedDateTime,links", $orderby: "displayName" },
      requestId,
    );
  }

  /** List sections in a notebook. */
  listSections(notebookId: string, requestId?: string): Promise<ListResult> {
    return this.getList(
      `/notebooks/${encodeURIComponent(notebookId)}/sections`,
      { $select: "id,displayName,isDefault,createdDateTime,lastModifiedDateTime", $orderby: "displayName" },
      requestId,
    );
  }

  /** List pages in a section (newest first). */
  listPages(sectionId: string, top = 25, requestId?: string): Promise<ListResult> {
    return this.getList(
      `/sections/${encodeURIComponent(sectionId)}/pages`,
      {
        $select: "id,title,createdDateTime,lastModifiedDateTime,links,contentUrl,level,order",
        $orderby: "lastModifiedDateTime desc",
        $top: String(Math.min(Math.max(top, 1), 100)),
      },
      requestId,
    );
  }

  /** Get a page's metadata (title, timestamps, links). */
  getPageMeta(pageId: string, requestId?: string): Promise<OneNoteRecord | null> {
    return this.getOne(
      `/pages/${encodeURIComponent(pageId)}`,
      { $select: "id,title,createdDateTime,lastModifiedDateTime,links,contentUrl" },
      requestId,
    );
  }

  /**
   * Get a page's HTML content. `includeIds` adds `data-id` attributes to every element,
   * which are the anchors required for a later surgical PATCH (see class note).
   */
  async getPageContent(pageId: string, includeIds = true, requestId?: string): Promise<PageContent> {
    const qs = includeIds ? "?includeIDs=true" : "";
    const url = `${this.root}/pages/${encodeURIComponent(pageId)}/content${qs}`;
    const { data } = await this.http.requestText(url, {
      headers: {
        Authorization: `Bearer ${await this.auth.getToken(requestId)}`,
        Accept: "text/html",
      },
      requestId,
    });
    return { pageId, html: data };
  }

  // ── Writes ─────────────────────────────────────────────────────────────────────

  /**
   * Create a new page in a section from an HTML document.
   * Graph accepts the page as `text/html`; the <title> becomes the page title.
   * Returns the created page's metadata (including its new id).
   */
  async createPage(sectionId: string, fullHtml: string, requestId?: string): Promise<OneNoteRecord | null> {
    const token = await this.auth.getToken(requestId);
    const { data } = await this.http.requestJson<OneNoteRecord>(
      `${this.root}/sections/${encodeURIComponent(sectionId)}/pages`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          "Content-Type": "text/html",
        },
        body: fullHtml,
        requestId,
      },
    );
    return data ?? null;
  }

  /**
   * Append an HTML block to the end of an existing page's body.
   * Uses the OneNote content-update command array; returns nothing (Graph replies 204).
   * `htmlBody` is an HTML fragment (e.g. "<p>note</p>"), not a full document.
   */
  async appendToPage(pageId: string, htmlBody: string, requestId?: string): Promise<void> {
    const commands = [{ target: "body", action: "append", position: "after", content: htmlBody }];
    await this.http.requestJson<undefined>(`${this.root}/pages/${encodeURIComponent(pageId)}/content`, {
      method: "PATCH",
      headers: await this.authHeaders(true, requestId),
      body: JSON.stringify(commands),
      requestId,
    });
  }
}
