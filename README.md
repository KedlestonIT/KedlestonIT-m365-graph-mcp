# Kedleston Microsoft 365 Graph MCP

Remote Model Context Protocol server for Microsoft 365, hosted on Cloudflare Workers and backed by Microsoft Graph.

## Security position

The server is deliberately deployed read-only first. Set `WRITE_ENABLED=true` only after read tests pass. Requests require an `MCP_API_KEY`, tenant access is pinned with `ALLOWED_TENANT_ID`, Graph paths are restricted to approved roots, and secrets are stored with Wrangler rather than committed.

This project uses Microsoft Graph application permissions. It does not grant access by itself: a Microsoft Entra administrator must approve each required permission. Microsoft Graph does not expose every Microsoft 365 administrative feature, so some workloads may need workload-specific APIs or PowerShell.

## 1. Create the Entra application

Create an app registration named `Kedleston M365 Graph MCP` in the Kedleston tenant. Record the tenant ID and application/client ID. Create a client secret initially; for production, plan to replace it with a certificate or an Entra workload identity design.

Add only the Microsoft Graph **Application** permissions required by the enabled tools, then grant admin consent. Suggested starting set:

- `User.Read.All`
- `Group.Read.All`
- `Sites.Read.All`
- `Files.Read.All`
- `Mail.Read`
- `Calendars.Read`
- `Team.ReadBasic.All`

For controlled write operations add only what is needed, for example:

- `Mail.Send`
- `Calendars.ReadWrite`
- `Sites.ReadWrite.All`
- `Files.ReadWrite.All`
- `Group.ReadWrite.All`

Avoid broad directory write permissions unless there is a separately approved use case. Use Exchange Online application access policies or Application RBAC to restrict which mailboxes the app can access.

## 2. Configure Cloudflare

Install dependencies and authenticate Wrangler:

```bash
npm install
npx wrangler login
```

Edit `wrangler.toml` and replace `REPLACE_WITH_TENANT_ID`.

Store all credentials as Cloudflare secrets:

```bash
npx wrangler secret put TENANT_ID
npx wrangler secret put CLIENT_ID
npx wrangler secret put CLIENT_SECRET
npx wrangler secret put MCP_API_KEY
```

Generate the API key with at least 32 random bytes, for example:

```bash
openssl rand -base64 48
```

## 3. Deploy read-only

```bash
npm run typecheck
npm run deploy
```

Test:

```bash
curl https://<scrubbed-url-domain>.<your-subdomain>.workers.dev/health
```

The MCP URL is the Worker URL. Configure the MCP client to send `Authorization: Bearer <MCP_API_KEY>`.

## 4. Enable write access

After testing read operations and reviewing Entra audit logs:

```bash
npx wrangler deploy --var WRITE_ENABLED:true --var ALLOWED_TENANT_ID:<tenant-id> --var GRAPH_BASE_URL:https://<scrubbed-url-domain>/v1.0
```

Prefer a separate production Worker and Entra application for write access. Apply Cloudflare Access in front of the Worker and restrict access to approved Kedleston administrators.

## Available tools

- `graph_read`
- `graph_write`
- `list_users`
- `list_sites`
- `list_user_messages`
- `send_mail`
- `list_calendar_events`

`graph_read` and `graph_write` provide broad Graph coverage while retaining tenant pinning, path validation, permission enforcement and a write kill switch.

## Recommended production controls

1. Separate read-only and write-enabled Entra applications.
2. Restrict Exchange mailbox scope with Application RBAC.
3. Use `Sites.Selected` rather than `Sites.ReadWrite.All` where practical.
4. Put Cloudflare Access in front of the Worker and require Entra ID sign-in.
5. Keep `WRITE_ENABLED=false` by default.
6. Review Entra sign-in logs, Microsoft Purview audit logs and Cloudflare Worker logs.
7. Rotate the client secret and MCP API key at least every 90 days until certificate-based authentication is implemented.
8. Add per-tool role checks and an immutable audit destination before enabling high-impact writes.
