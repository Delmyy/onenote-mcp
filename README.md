# OneNote MCP server

A secure remote **MCP server** around the **Microsoft Graph OneNote API** (v1.0).
Serves an MCP endpoint over **Streamable HTTP** that claude.ai (as a custom connector),
`claude mcp`, ChatGPT, or the MCP Inspector can register. Modeled on the sibling
`lawmatics-mcp` and deploys the same way (Azure Container App).

- **Read-only by default.** Write tools (`create_page`, `append_to_page`) are registered only
  when `ENABLE_WRITES=true`, and each additionally requires an explicit `confirm: true` (a
  dry-run preview is returned otherwise).
- **No secrets in source.** All credentials come from env / Azure Key Vault.
- **No content in logs.** Titles, page HTML, and tokens are never logged. See [`src/logging.ts`](src/logging.ts).

## Tools

| Tool (registered name) | Requested name | Purpose |
|---|---|---|
| `onenote_list_notebooks` | `list_notebooks` | List notebooks for the configured user/group/site |
| `onenote_list_sections` | `list_sections(notebookId)` | List sections in a notebook |
| `onenote_list_pages` | `list_pages(sectionId)` | List pages in a section (newest first) |
| `onenote_read_page` | `read_page(pageId)` | Return a page's HTML/text content |
| `onenote_create_page` | `create_page(sectionId, title, htmlBody)` | Create a page (write; gated) |
| `onenote_append_to_page` | `append_to_page(pageId, htmlBody)` | Append an HTML block to a page (write; gated) |

Tools are prefixed with `onenote_` for discoverability (house convention), matching how
`lawmatics-mcp` prefixes its tools. Page content is **HTML** throughout.

> **Surgical edits.** `append_to_page` is a coarse whole-body append. To change a *specific*
> existing element, you must `PATCH /onenote/pages/{id}/content` with a command whose `target`
> is that element's `data-id`. Those anchors only exist if you first read the page with
> `includeIds=true` (`onenote_read_page` does this by default). This is documented in
> [`src/onenoteClient.ts`](src/onenoteClient.ts) as an extension point; no destructive
> edit/replace tool is shipped by design.

---

## (a) Azure app registration + admin consent

You need **one** app registration. Choose **application** (app-only, client credentials) or
**delegated** (on behalf of a signed-in user). Application mode is the primary/default flow.

### 1. Register the app
1. **Entra admin center** → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name: `onenote-mcp`. Supported account types: **Single tenant** (typical). Register.
3. From **Overview**, copy the **Directory (tenant) ID** → `GRAPH_TENANT_ID` and the
   **Application (client) ID** → `GRAPH_CLIENT_ID`.

### 2. Add a client secret
1. **Certificates & secrets** → **Client secrets** → **New client secret** → set an expiry.
2. Copy the secret **Value** (not the Secret ID) → `GRAPH_CLIENT_SECRET`. Store it in Key Vault.

### 3. Add the Graph permission + grant admin consent

**Application mode (client credentials — default):**
1. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**.
2. Add **`Notes.ReadWrite.All`**.
3. Click **Grant admin consent for &lt;tenant&gt;** (a Global Admin must do this). Status must show a green check.
4. Set `GRAPH_AUTH_MODE=application`, `ONENOTE_TARGET_TYPE=user`, and `ONENOTE_TARGET_ID=<UPN or object id>`
   (app-only has no "me" — you must name whose notebooks to act on; `group`/`site` also work).

**Delegated mode (on behalf of a user):**
1. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
2. Add **`Notes.ReadWrite`** and **`offline_access`** (and `User.Read`). Grant admin consent (or let the
   user consent at first sign-in).
3. **Authentication** → **Add a platform** → **Web** → redirect URI `http://localhost:8082/oauth/callback`
   (add your production callback too if you script re-auth there). Save.
4. Set `GRAPH_AUTH_MODE=delegated` and mint a refresh token (see below); `ONENOTE_TARGET_TYPE` defaults to `me`.

### 4. (Delegated only) mint the refresh token — one time
Access tokens are short-lived; `offline_access` yields a refresh token the server caches and rotates.

```bash
# set GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET first (see .env.example)
npm run authorize
```

Sign in and approve; copy the printed **refresh token** into your secret manager as
`GRAPH_REFRESH_TOKEN`. The script uses PKCE and never writes the token to disk.

---

## (b) Run locally

> **Windows/OneDrive note:** don't `npm install` inside a OneDrive-synced folder. Copy the project
> to a local path first (e.g. `C:\dev\onenote-mcp`), install/build there, and keep only source in OneDrive.

```bash
cp .env.example .env       # fill via your secret manager
npm install
npm run build
npm start                  # serves POST http://localhost:8082/mcp   (health: GET /healthz)
```

Quick smoke test with the MCP Inspector:

```bash
npx @modelcontextprotocol/inspector
# Transport: Streamable HTTP · URL: http://localhost:8082/mcp
# Header:    Authorization: Bearer <MCP_AUTH_TOKEN>
```

Required env: `MCP_AUTH_TOKEN`, `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET`,
plus `ONENOTE_TARGET_ID` (application mode) or `GRAPH_REFRESH_TOKEN` (delegated mode).
Writes are off unless `ENABLE_WRITES=true`.

---

## (c) Deploy as an Azure Container App

Same pattern as `lawmatics-mcp`. Secrets live in Container Apps secrets (optionally sourced from
Azure Key Vault); nothing is baked into the image.

```bash
# variables
RG=ctg-mcp-rg
LOC=eastus
ACR=ctgmcpacr                      # must be globally unique
ENV=ctg-mcp-env
APP=onenote-mcp

# 0) resource group
az group create -n $RG -l $LOC

# 1) build & push the image with ACR (no local Docker needed)
az acr create -n $ACR -g $RG --sku Basic --admin-enabled true
az acr build -r $ACR -t onenote-mcp:latest .

# 2) Container Apps environment
az containerapp env create -n $ENV -g $RG -l $LOC

# 3) create the app (ingress 8082) with secrets
az containerapp create \
  -n $APP -g $RG --environment $ENV \
  --image $ACR.azurecr.io/onenote-mcp:latest \
  --registry-server $ACR.azurecr.io \
  --target-port 8082 --ingress external \
  --min-replicas 1 --max-replicas 3 \
  --secrets \
     mcp-auth-token=$MCP_AUTH_TOKEN \
     graph-client-secret=$GRAPH_CLIENT_SECRET \
  --env-vars \
     PORT=8082 \
     MCP_PATH=/mcp \
     MCP_AUTH_TOKEN=secretref:mcp-auth-token \
     GRAPH_AUTH_MODE=application \
     GRAPH_TENANT_ID=$GRAPH_TENANT_ID \
     GRAPH_CLIENT_ID=$GRAPH_CLIENT_ID \
     GRAPH_CLIENT_SECRET=secretref:graph-client-secret \
     ONENOTE_TARGET_TYPE=user \
     ONENOTE_TARGET_ID=$ONENOTE_TARGET_ID \
     ENABLE_WRITES=false

# 4) get the public URL — your MCP endpoint is https://<fqdn>/mcp
az containerapp show -n $APP -g $RG --query properties.configuration.ingress.fqdn -o tsv
```

To enable writes later: `az containerapp update -n $APP -g $RG --set-env-vars ENABLE_WRITES=true`.

**Key Vault (recommended):** store `MCP_AUTH_TOKEN`, `GRAPH_CLIENT_SECRET`, and (delegated)
`GRAPH_REFRESH_TOKEN` in Key Vault, give the app a managed identity with **Key Vault Secrets User**,
and reference them via `--secrets mcp-auth-token=keyvaultref:<secret-uri>,identityref:<identity-id>`.

---

## (d) Connect it

### As a claude.ai connector
1. Deploy (above) and confirm `GET https://<fqdn>/healthz` returns `{"status":"ok"}`.
2. In **claude.ai → Settings → Connectors → Add custom connector**:
   - **URL:** `https://<fqdn>/mcp`
   - **Authentication:** the connector must send `Authorization: Bearer <MCP_AUTH_TOKEN>`.
3. Save; the six OneNote tools appear (writes only if `ENABLE_WRITES=true`).

### Via `claude mcp` (Claude Code)
```bash
claude mcp add onenote --transport http https://<fqdn>/mcp \
  --header "Authorization: Bearer <MCP_AUTH_TOKEN>"
```
Or local dev against `http://localhost:8082/mcp`. Verify with `claude mcp list`, then `/mcp` in-session.

---

## Graph facts encoded here

- Base `https://graph.microsoft.com/v1.0`; bearer auth; OData `$select`/`$orderby`/`$top`;
  list envelope `{ value: [...], "@odata.nextLink" }`.
- OneNote lives under a user/group/site context: `/me/onenote` (delegated) or
  `/users/{id}/onenote` (application). **App-only cannot use `/me`.**
- Page content is HTML. Create = `POST .../sections/{id}/pages` (`Content-Type: text/html`, title
  from `<title>`); append = `PATCH .../pages/{id}/content` with a JSON command array; read =
  `GET .../pages/{id}/content?includeIDs=true` (returns HTML with `data-id` anchors).
- Client-credentials access tokens (~60–90 min) are cached in memory and re-acquired ~1 min before
  expiry; delegated refresh tokens are used and rotated in memory. See [`src/graphAuth.ts`](src/graphAuth.ts).
