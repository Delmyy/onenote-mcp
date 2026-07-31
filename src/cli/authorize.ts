/**
 * One-time Microsoft Graph (Entra ID) delegated authorization helper.
 *
 * Runs the OAuth 2.0 authorization-code flow WITH PKCE to mint a long-lived refresh
 * token for delegated mode (GRAPH_AUTH_MODE=delegated). The refresh token is then stored
 * in your secret manager as GRAPH_REFRESH_TOKEN. This script never writes it to disk.
 *
 * You only need this for DELEGATED mode. Application (client-credentials) mode needs no
 * user sign-in and therefore no refresh token.
 *
 * Usage (PowerShell):
 *   $env:GRAPH_TENANT_ID="..."; $env:GRAPH_CLIENT_ID="..."; $env:GRAPH_CLIENT_SECRET="..."
 *   npm run authorize
 *
 * Usage (bash):
 *   GRAPH_TENANT_ID=... GRAPH_CLIENT_ID=... GRAPH_CLIENT_SECRET=... npm run authorize
 *
 * Notes:
 *   - The app registration must have a redirect URI of http://localhost:8082/oauth/callback
 *     registered under platform "Web" (or "Mobile & desktop"), and the delegated permission
 *     Notes.ReadWrite + offline_access.
 *   - A public client can omit the secret; a confidential (Web) client includes it (we send it
 *     when present — Entra accepts client_secret alongside PKCE for confidential clients).
 */

import http from "node:http";
import { randomBytes, createHash } from "node:crypto";

function reqEnv(name: string): string {
  const v = process.env[name];
  if (!v || v.startsWith("replace-with")) {
    console.error(`Missing env ${name}. Set it before running (see .env.example).`);
    process.exit(1);
  }
  return v;
}

function base64url(buf: Buffer): string {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

async function main(): Promise<void> {
  const tenantId = reqEnv("GRAPH_TENANT_ID");
  const clientId = reqEnv("GRAPH_CLIENT_ID");
  const clientSecret = process.env.GRAPH_CLIENT_SECRET ?? "";
  const scopes = process.env.GRAPH_DELEGATED_SCOPES ?? "offline_access Notes.ReadWrite User.Read";
  const redirectUri = process.env.GRAPH_REDIRECT_URI ?? "http://localhost:8082/oauth/callback";

  const url = new URL(redirectUri);
  const port = Number(url.port || 8082);
  const callbackPath = url.pathname || "/oauth/callback";

  const authorizeUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const state = base64url(randomBytes(16));
  const codeVerifier = base64url(randomBytes(32));
  const codeChallenge = base64url(createHash("sha256").update(codeVerifier).digest());

  const authUrl = new URL(authorizeUrl);
  authUrl.searchParams.set("client_id", clientId);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", redirectUri);
  authUrl.searchParams.set("response_mode", "query");
  authUrl.searchParams.set("scope", scopes);
  authUrl.searchParams.set("state", state);
  authUrl.searchParams.set("code_challenge", codeChallenge);
  authUrl.searchParams.set("code_challenge_method", "S256");

  console.log("\n1) Open this URL in your browser and sign in / approve access:\n");
  console.log("   " + authUrl.toString() + "\n");
  console.log(`2) Waiting for the redirect to ${redirectUri} ...\n`);

  const code = await new Promise<string>((resolve, reject) => {
    const server = http.createServer((rq, rs) => {
      const reqUrl = new URL(rq.url ?? "/", `http://localhost:${port}`);
      if (reqUrl.pathname !== callbackPath) {
        rs.writeHead(404).end("Not found");
        return;
      }
      const returnedState = reqUrl.searchParams.get("state");
      const returnedCode = reqUrl.searchParams.get("code");
      const err = reqUrl.searchParams.get("error");
      if (err) {
        rs.writeHead(400).end(`Authorization error: ${err}`);
        server.close();
        reject(new Error(err));
        return;
      }
      if (returnedState !== state) {
        rs.writeHead(400).end("State mismatch — possible CSRF. Aborted.");
        server.close();
        reject(new Error("State mismatch"));
        return;
      }
      if (!returnedCode) {
        rs.writeHead(400).end("No authorization code returned.");
        server.close();
        reject(new Error("No code"));
        return;
      }
      rs.writeHead(200, { "Content-Type": "text/html" }).end(
        "<h3>Authorized. You can close this tab and return to the terminal.</h3>",
      );
      server.close();
      resolve(returnedCode);
    });
    server.listen(port);
    server.on("error", reject);
  });

  console.log("3) Exchanging the code for tokens ...\n");
  const form: Record<string, string> = {
    client_id: clientId,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    scope: scopes,
    code_verifier: codeVerifier,
  };
  if (clientSecret) form.client_secret = clientSecret;

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
    body: new URLSearchParams(form).toString(),
  });
  if (!res.ok) {
    console.error(`Token exchange failed (HTTP ${res.status}). Check client id/secret, redirect URI, and scopes.`);
    process.exit(1);
  }
  const json = (await res.json()) as { refresh_token?: string; access_token?: string };
  if (!json.refresh_token) {
    console.error(
      "Token exchange succeeded but no refresh_token was returned. Ensure 'offline_access' is in GRAPH_DELEGATED_SCOPES.",
    );
    process.exit(1);
  }

  console.log("✅ Success. Store this value in your secret manager as GRAPH_REFRESH_TOKEN:\n");
  console.log("   " + json.refresh_token + "\n");
  console.log("Security notes:");
  console.log("  • This refresh token can mint access tokens for the OneNote scopes — treat it like a password.");
  console.log("  • Do NOT paste it into source, chat, or a committed .env. Use Azure Key Vault / your secret manager.");
  console.log("  • Clear your terminal scrollback after copying it.\n");
}

main().catch((err) => {
  console.error("authorize failed:", (err as Error).message);
  process.exit(1);
});
