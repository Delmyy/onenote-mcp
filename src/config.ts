/**
 * Environment configuration loader + validator.
 * Fails fast with a safe (secret-free) message if required values are missing.
 */

export type GraphAuthMode = "application" | "delegated";
export type TargetType = "me" | "user" | "group" | "site";

export interface Config {
  port: number;
  mcpPath: string;
  mcpAuthToken: string;
  graph: {
    apiBase: string;
    authMode: GraphAuthMode;
    tenantId: string;
    clientId: string;
    clientSecret: string;
    /** Delegated only: space-delimited scopes used with the refresh_token grant. */
    delegatedScopes: string;
    /** Delegated only: long-lived refresh token minted by `npm run authorize`. */
    refreshToken: string;
    /** Whose notebooks to act on. Application mode CANNOT use "me". */
    targetType: TargetType;
    /** User id/UPN, group id, or site id. Ignored when targetType === "me". */
    targetId: string;
  };
  enableWrites: boolean;
  httpTimeoutMs: number;
  rateLimitRps: number;
  logLevel: "error" | "warn" | "info" | "debug";
}

function req(name: string): string {
  const v = process.env[name];
  if (!v || v.trim() === "" || v.startsWith("replace-with")) {
    throw new Error(
      `Missing required environment variable: ${name}. ` +
        `Set it in your secret manager / .env (see .env.example). Never hardcode it.`,
    );
  }
  return v.trim();
}

function opt(name: string, fallback: string): string {
  const v = process.env[name];
  return v && v.trim() !== "" ? v.trim() : fallback;
}

function num(name: string, fallback: number): number {
  const v = process.env[name];
  if (!v) return fallback;
  const n = Number(v);
  if (!Number.isFinite(n) || n <= 0) {
    throw new Error(`Environment variable ${name} must be a positive number.`);
  }
  return n;
}

let cached: Config | null = null;

export function loadConfig(): Config {
  if (cached) return cached;

  const enableWrites = opt("ENABLE_WRITES", "false").toLowerCase() === "true";
  const logLevel = opt("LOG_LEVEL", "info") as Config["logLevel"];
  if (!["error", "warn", "info", "debug"].includes(logLevel)) {
    throw new Error(`LOG_LEVEL must be one of error|warn|info|debug`);
  }

  const authMode = opt("GRAPH_AUTH_MODE", "application").toLowerCase() as GraphAuthMode;
  if (authMode !== "application" && authMode !== "delegated") {
    throw new Error(`GRAPH_AUTH_MODE must be "application" or "delegated".`);
  }

  const targetType = opt("ONENOTE_TARGET_TYPE", authMode === "delegated" ? "me" : "user").toLowerCase() as TargetType;
  if (!["me", "user", "group", "site"].includes(targetType)) {
    throw new Error(`ONENOTE_TARGET_TYPE must be one of me|user|group|site.`);
  }

  // App-only (client credentials) has no signed-in user, so "/me" is invalid on Graph.
  // Application mode MUST target a specific user, group, or site.
  if (authMode === "application" && targetType === "me") {
    throw new Error(
      `GRAPH_AUTH_MODE=application cannot use ONENOTE_TARGET_TYPE=me. ` +
        `App-only tokens have no user context — set ONENOTE_TARGET_TYPE to user|group|site and ONENOTE_TARGET_ID.`,
    );
  }
  const targetId = targetType === "me" ? "" : req("ONENOTE_TARGET_ID");

  cached = {
    port: num("PORT", 8082),
    mcpPath: opt("MCP_PATH", "/mcp"),
    mcpAuthToken: req("MCP_AUTH_TOKEN"),
    graph: {
      apiBase: opt("GRAPH_API_BASE", "https://graph.microsoft.com/v1.0").replace(/\/$/, ""),
      authMode,
      tenantId: req("GRAPH_TENANT_ID"),
      clientId: req("GRAPH_CLIENT_ID"),
      clientSecret: req("GRAPH_CLIENT_SECRET"),
      delegatedScopes: opt("GRAPH_DELEGATED_SCOPES", "offline_access Notes.ReadWrite User.Read"),
      // Refresh token is only required in delegated mode.
      refreshToken: authMode === "delegated" ? req("GRAPH_REFRESH_TOKEN") : opt("GRAPH_REFRESH_TOKEN", ""),
      targetType,
      targetId,
    },
    enableWrites,
    httpTimeoutMs: num("HTTP_TIMEOUT_MS", 20000),
    rateLimitRps: num("RATE_LIMIT_RPS", 5),
    logLevel,
  };
  return cached;
}

/** Test helper: clear the memoized config so tests can re-load with different env. */
export function _resetConfigForTests(): void {
  cached = null;
}
