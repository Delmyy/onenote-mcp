/**
 * Hardened HTTP helper: per-request timeout, client-side rate limiting, and
 * bounded retries with backoff for 429/5xx. Returns a normalized result and
 * NEVER throws raw upstream bodies (which could contain content/PII) to callers.
 *
 * Exposes both requestJson (Graph metadata + token endpoint) and requestText
 * (OneNote page /content, which returns raw HTML rather than JSON).
 */

import { log } from "./logging.js";

export class HttpError extends Error {
  constructor(
    public status: number,
    /** Safe, content-free message suitable for returning to the model/user. */
    public safeMessage: string,
  ) {
    super(safeMessage);
    this.name = "HttpError";
  }
}

/** Simple token-bucket limiter shared across all requests from this process. */
class RateLimiter {
  private queue: Array<() => void> = [];
  private tokens: number;
  constructor(private rps: number) {
    this.tokens = rps;
    setInterval(() => {
      this.tokens = this.rps;
      while (this.tokens > 0 && this.queue.length > 0) {
        this.tokens--;
        this.queue.shift()!();
      }
    }, 1000).unref?.();
  }
  acquire(): Promise<void> {
    if (this.tokens > 0) {
      this.tokens--;
      return Promise.resolve();
    }
    return new Promise((resolve) => this.queue.push(resolve));
  }
}

export interface HttpClientOptions {
  timeoutMs: number;
  rateLimitRps: number;
  maxRetries?: number;
}

export interface RequestOptions {
  method?: string;
  headers?: Record<string, string>;
  /** Already-serialized body (JSON string, HTML string, or form body), or undefined. */
  body?: string;
  requestId?: string;
}

export interface HttpResult<T> {
  status: number;
  data: T;
}

export class HttpClient {
  private limiter: RateLimiter;
  private maxRetries: number;
  constructor(private opts: HttpClientOptions) {
    this.limiter = new RateLimiter(opts.rateLimitRps);
    this.maxRetries = opts.maxRetries ?? 3;
  }

  /** Core send loop with timeout + rate limit + bounded retry. Throws HttpError on !ok. */
  private async send(url: string, req: RequestOptions): Promise<Response> {
    const method = req.method ?? "GET";
    let attempt = 0;

    while (true) {
      attempt++;
      await this.limiter.acquire();
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), this.opts.timeoutMs);
      const started = Date.now();
      try {
        const res = await fetch(url, {
          method,
          headers: req.headers,
          body: req.body,
          signal: controller.signal,
        });
        const latencyMs = Date.now() - started;

        // Retry on 429 / 5xx (bounded).
        if ((res.status === 429 || res.status >= 500) && attempt <= this.maxRetries) {
          const retryAfter = Number(res.headers.get("retry-after")) || 0;
          const backoff = retryAfter > 0 ? retryAfter * 1000 : Math.min(2 ** attempt * 250, 4000);
          log.warn("http.retry", { status: res.status, attempt, latencyMs, requestId: req.requestId });
          await safeText(res);
          await sleep(backoff);
          continue;
        }

        if (!res.ok) {
          // Read the body to free the socket. Surface ONLY Microsoft-authored diagnostics
          // (error code/description + which host failed), length-capped — no user content.
          const detail = extractGraphError(await safeReadText(res));
          let host = "";
          try {
            host = new URL(url).host;
          } catch {
            /* ignore */
          }
          log.warn("http.error", { status: res.status, host, latencyMs, requestId: req.requestId, code: detail?.code });
          throw new HttpError(
            res.status,
            safeStatusMessage(res.status) +
              ` [${host}${detail ? ` ${detail.code}: ${detail.message}` : " (no error body)"}]`,
          );
        }

        log.debug("http.ok", { status: res.status, latencyMs, requestId: req.requestId });
        return res;
      } catch (err) {
        if (err instanceof HttpError) throw err;
        // Network / timeout / abort.
        const aborted = (err as Error)?.name === "AbortError";
        if (attempt <= this.maxRetries && (aborted || isTransient(err))) {
          log.warn("http.retry", { reason: aborted ? "timeout" : "network", attempt, requestId: req.requestId });
          await sleep(Math.min(2 ** attempt * 250, 4000));
          continue;
        }
        log.error("http.fail", { reason: aborted ? "timeout" : "network", requestId: req.requestId });
        throw new HttpError(0, aborted ? "Upstream request timed out." : "Upstream request failed.");
      } finally {
        clearTimeout(timer);
      }
    }
  }

  /** Perform a request and parse a JSON response (204 → undefined). */
  async requestJson<T = unknown>(url: string, req: RequestOptions = {}): Promise<HttpResult<T>> {
    const res = await this.send(url, req);
    const data = (res.status === 204 ? (undefined as T) : ((await res.json()) as T));
    return { status: res.status, data };
  }

  /** Perform a request and return the raw text body (used for OneNote page HTML content). */
  async requestText(url: string, req: RequestOptions = {}): Promise<HttpResult<string>> {
    const res = await this.send(url, req);
    const data = res.status === 204 ? "" : await res.text();
    return { status: res.status, data };
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

async function safeText(res: Response): Promise<void> {
  try {
    await res.text();
  } catch {
    /* ignore */
  }
}

async function safeReadText(res: Response): Promise<string> {
  try {
    return await res.text();
  } catch {
    return "";
  }
}

/**
 * Extract Microsoft-authored error diagnostics (safe fields only, length-capped).
 * Handles both shapes:
 *   - Graph:            { error: { code, message } }
 *   - OAuth token endpt:{ error: "invalid_client", error_description: "..." }
 * Falls back to a short snippet for non-JSON bodies.
 */
function extractGraphError(body: string): { code: string; message: string } | null {
  const clip = (s: unknown, n: number) => String(s ?? "").slice(0, n);
  try {
    const j = JSON.parse(body) as {
      error?: { code?: string; message?: string } | string;
      error_description?: string;
    };
    const e = j?.error;
    if (e && typeof e === "object" && (e.code || e.message)) {
      return { code: clip(e.code, 80), message: clip(e.message, 400) };
    }
    if (typeof e === "string") {
      return { code: clip(e, 80), message: clip(j.error_description, 400) };
    }
  } catch {
    /* non-JSON error body — fall through to snippet */
  }
  const trimmed = body.trim();
  return trimmed ? { code: "non-json", message: trimmed.slice(0, 200) } : null;
}

function isTransient(err: unknown): boolean {
  const code = (err as { cause?: { code?: string } })?.cause?.code;
  return code === "ECONNRESET" || code === "ETIMEDOUT" || code === "ECONNREFUSED" || code === "EAI_AGAIN";
}

/** Map status codes to safe, actionable, content-free messages. */
function safeStatusMessage(status: number): string {
  switch (status) {
    case 400:
      return "The request was rejected as invalid (400). Check the parameters and try again.";
    case 401:
      return "Authentication with Microsoft Graph failed (401). The token may be expired, missing a required permission, or admin consent may not be granted.";
    case 403:
      return "Microsoft Graph denied access (403). The app/user may lack the Notes.ReadWrite(.All) permission, or admin consent is missing.";
    case 404:
      return "The requested notebook, section, or page was not found (404).";
    case 409:
      return "The request conflicts with the current state of the resource (409).";
    case 422:
      return "Graph could not process the request (422). One or more values were not accepted.";
    case 429:
      return "Microsoft Graph is throttling requests (429). Try again shortly.";
    default:
      return status >= 500
        ? `Microsoft Graph had a server error (${status}). Try again later.`
        : `Microsoft Graph returned an unexpected status (${status}).`;
  }
}
