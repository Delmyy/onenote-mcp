/**
 * Microsoft Graph token provider (Azure AD / Entra ID).
 *
 * Supports two flows, selected by GRAPH_AUTH_MODE:
 *
 *  1. "application" — OAuth 2.0 **client credentials** (app-only).
 *     grant_type=client_credentials, scope=https://graph.microsoft.com/.default.
 *     Requires the app's Graph *application* permission `Notes.ReadWrite.All` with
 *     admin consent granted. There is NO user context, so callers must target a
 *     specific user/group/site (see onenoteClient resourceRoot). No refresh token is
 *     issued — you simply request a new token when the cached one nears expiry.
 *
 *  2. "delegated" — OAuth 2.0 **refresh_token** grant on behalf of a signed-in user.
 *     Requires the *delegated* permission `Notes.ReadWrite` plus `offline_access`.
 *     The long-lived refresh token is minted once via `npm run authorize` and stored
 *     as GRAPH_REFRESH_TOKEN. Access tokens are cached; when Entra rotates the refresh
 *     token on a refresh, we keep the new one in memory for the life of the process.
 *
 * In BOTH cases access tokens are short-lived (~60–90 min). This provider caches the
 * token in memory and transparently re-acquires it ~1 minute before expiry.
 *
 * Tokens, secrets, and refresh tokens are NEVER logged.
 */

import type { HttpClient } from "./http.js";
import { log } from "./logging.js";

interface TokenResponse {
  access_token?: string;
  expires_in?: number;
  refresh_token?: string;
  error?: string;
  error_description?: string;
}

export interface GraphAuthConfig {
  authMode: "application" | "delegated";
  tenantId: string;
  clientId: string;
  clientSecret: string;
  /** Delegated only. */
  delegatedScopes: string;
  /** Delegated only; may be rotated in memory as Entra returns new ones. */
  refreshToken: string;
}

export class GraphAuthProvider {
  private cachedToken: string | null = null;
  private expiresAtMs = 0;
  private refreshToken: string;
  /** Guards against a thundering herd of concurrent token requests. */
  private inflight: Promise<string> | null = null;

  constructor(
    private http: HttpClient,
    private cfg: GraphAuthConfig,
  ) {
    this.refreshToken = cfg.refreshToken;
  }

  private tokenUrl(): string {
    return `https://login.microsoftonline.com/${encodeURIComponent(this.cfg.tenantId)}/oauth2/v2.0/token`;
  }

  /** Return a valid access token, acquiring/refreshing as needed. */
  async getToken(requestId?: string): Promise<string> {
    const now = Date.now();
    // 60s safety margin so a token never expires mid-request.
    if (this.cachedToken && this.expiresAtMs > now + 60_000) return this.cachedToken;
    if (this.inflight) return this.inflight;

    this.inflight = this.acquire(requestId).finally(() => {
      this.inflight = null;
    });
    return this.inflight;
  }

  private async acquire(requestId?: string): Promise<string> {
    const params: Record<string, string> =
      this.cfg.authMode === "application"
        ? {
            client_id: this.cfg.clientId,
            client_secret: this.cfg.clientSecret,
            grant_type: "client_credentials",
            scope: "https://graph.microsoft.com/.default",
          }
        : {
            client_id: this.cfg.clientId,
            client_secret: this.cfg.clientSecret,
            grant_type: "refresh_token",
            refresh_token: this.mustRefreshToken(),
            scope: this.cfg.delegatedScopes,
          };

    const { data } = await this.http.requestJson<TokenResponse>(this.tokenUrl(), {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Accept: "application/json",
      },
      body: new URLSearchParams(params).toString(),
      requestId,
    });

    if (!data?.access_token) {
      // data.error/error_description come from Entra and are safe (no secrets), but keep the
      // surfaced message generic; the detailed reason is available in server logs by request id.
      log.error("graph.token_failed", { requestId, authMode: this.cfg.authMode, reason: data?.error });
      throw new Error(
        "Failed to obtain a Microsoft Graph access token. Verify the tenant, client id/secret, " +
          "the required permission, and (for application mode) that admin consent was granted.",
      );
    }

    this.cachedToken = data.access_token;
    this.expiresAtMs = Date.now() + (data.expires_in ?? 3600) * 1000;
    // Entra may rotate the refresh token on each delegated refresh — keep the newest.
    if (this.cfg.authMode === "delegated" && data.refresh_token) {
      this.refreshToken = data.refresh_token;
    }
    log.info("graph.token_ok", { requestId, authMode: this.cfg.authMode });
    return this.cachedToken;
  }

  private mustRefreshToken(): string {
    if (!this.refreshToken || this.refreshToken.startsWith("replace-with")) {
      throw new Error(
        "GRAPH_AUTH_MODE=delegated requires GRAPH_REFRESH_TOKEN. Mint one with `npm run authorize`.",
      );
    }
    return this.refreshToken;
  }
}
