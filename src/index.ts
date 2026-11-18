import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createMcpHandler } from "agents/mcp";
import { z } from "zod";

interface Env {
  TENANT_ID: string;
  CLIENT_ID: string;
  CLIENT_SECRET: string;
  GRAPH_BASE_URL: string;
  WRITE_ENABLED: string;
  ALLOWED_TENANT_ID: string;
  MCP_API_KEY: string;
}

const SAFE_ROOTS = new Set([
  "users", "groups", "applications", "servicePrincipals", "organization",
  "sites", "drives", "teams", "chats", "reports", "security", "auditLogs",
  "identity", "deviceManagement", "education", "communications", "places",
  "planner", "solutions", "admin", "directory", "domains", "subscriptions"
]);

const WRITE_METHODS = new Set(["POST", "PATCH", "PUT", "DELETE"]);
let cachedToken: { value: string; expiresAt: number } | undefined;

function jsonText(value: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(value, null, 2) }] };
}

function validatePath(path: string): string {
  const clean = path.trim().replace(/^\/+/, "");
  if (!clean || clean.includes("..") || /^https?:\/\//i.test(clean)) throw new Error("A relative Microsoft Graph path is required.");
  const root = clean.split(/[/?(]/, 1)[0];
  if (!SAFE_ROOTS.has(root) && root !== "me") throw new Error(`Graph root '${root}' is not allowed.`);
  return clean;
}

async function getToken(env: Env): Promise<string> {
  if (env.ALLOWED_TENANT_ID !== "REPLACE_WITH_TENANT_ID" && env.TENANT_ID !== env.ALLOWED_TENANT_ID) {
    throw new Error("TENANT_ID does not match ALLOWED_TENANT_ID.");
  }
  const now = Math.floor(Date.now() / 1000);
  if (cachedToken && cachedToken.expiresAt > now + 120) return cachedToken.value;
  const body = new URLSearchParams({
    client_id: env.CLIENT_ID,
    client_secret: env.CLIENT_SECRET,
    scope: "https://<scrubbed-url-domain>/.default",
    grant_type: "client_credentials"
  });
  const response = await fetch(`https://<scrubbed-url-domain>/$%7BencodeURIComponent(env.TENANT_ID)}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body
  });
  const data = await response.json() as { access_token?: string; expires_in?: number; error_description?: string };
  if (!response.ok || !data.access_token) throw new Error(data.error_description || `Token request failed (${response.status}).`);
  cachedToken = { value: data.access_token, expiresAt: now + (data.expires_in || 3600) };
  return data.access_token;
}

async function graph(env: Env, method: string, path: string, body?: unknown, headers?: Record<string, string>) {
  const upper = method.toUpperCase();
  if (WRITE_METHODS.has(upper) && env.WRITE_ENABLED !== "true") throw new Error("Write operations are disabled. Set WRITE_ENABLED=true after testing read-only access.");
  const token = await getToken(env);
  const clean = validatePath(path);
  const response = await fetch(`${env.GRAPH_BASE_URL.replace(/\/$/, "")}/${clean}`, {
    method: upper,
    headers: {
      authorization: `Bearer ${token}`,
      accept: "application/json",
      ...(body !== undefined ? { "content-type": "application/json" } : {}),
      ...(headers || {})
    },
    body: body === undefined ? undefined : JSON.stringify(body)
  });
  const text = await response.text();
  let data: unknown = text;
  try { data = text ? JSON.parse(text) : { status: response.status }; } catch {}
  if (!response.ok) throw new Error(`Microsoft Graph ${response.status}: ${typeof data === "string" ? data : JSON.stringify(data)}`);
  return data;
}

function buildServer(env: Env) {
  const server = new McpServer({ name: "Kedleston Microsoft 365 Graph MCP", version: "1.0.0" });

  server.tool("graph_read", "Read data from an allowed Microsoft Graph v1.0 endpoint.", {
    path: z.string().describe("Relative Graph path, for example users?$top=10 or sites/root"),
    consistencyLevel: z.enum(["eventual"]).optional()
  }, async ({ path, consistencyLevel }) => {
    try { return jsonText(await graph(env, "GET", path, undefined, consistencyLevel ? { ConsistencyLevel: consistencyLevel } : undefined)); }
    catch (error) { return { isError: true, ...jsonText({ error: error instanceof Error ? error.message : String(error) }) }; }
  });

  server.tool("graph_write", "Create, update or delete data through an allowed Microsoft Graph v1.0 endpoint. Disabled unless WRITE_ENABLED=true.", {
    method: z.enum(["POST", "PATCH", "PUT", "DELETE"]),
    path: z.string(),
    body: z.unknown().optional()
  }, async ({ method, path, body }) => {
    try { return jsonText(await graph(env, method, path, body)); }
    catch (error) { return { isError: true, ...jsonText({ error: error instanceof Error ? error.message : String(error) }) }; }
  });

  server.tool("list_users", "List Microsoft 365 users.", {
    top: z.number().int().min(1).max(100).default(25),
    select: z.string().default("id,displayName,userPrincipalName,mail,accountEnabled")
  }, async ({ top, select }) => jsonText(await graph(env, "GET", `users?$top=${top}&$select=${encodeURIComponent(select)}`)));

  server.tool("list_sites", "Search SharePoint sites by keyword.", { query: z.string().min(1) },
    async ({ query }) => jsonText(await graph(env, "GET", `sites?search=${encodeURIComponent(query)}`)));

  server.tool("list_user_messages", "List messages in a user's mailbox. Requires Mail.Read application permission.", {
    userId: z.string(), top: z.number().int().min(1).max(50).default(20)
  }, async ({ userId, top }) => jsonText(await graph(env, "GET", `users/${encodeURIComponent(userId)}/messages?$top=${top}&$select=id,subject,from,receivedDateTime,isRead&$orderby=receivedDateTime desc`)));

  server.tool("send_mail", "Send an email as a specified user. Requires Mail.Send application permission and WRITE_ENABLED=true.", {
    userId: z.string(), to: z.array(z.string().email()).min(1), subject: z.string(), body: z.string(), contentType: z.enum(["Text", "HTML"]).default("HTML"), saveToSentItems: z.boolean().default(true)
  }, async ({ userId, to, subject, body, contentType, saveToSentItems }) => jsonText(await graph(env, "POST", `users/${encodeURIComponent(userId)}/sendMail`, {
    message: { subject, body: { contentType, content: body }, toRecipients: to.map(address => ({ emailAddress: { address } })) }, saveToSentItems
  })));

  server.tool("list_calendar_events", "List calendar events for a user.", {
    userId: z.string(), start: z.string(), end: z.string(), top: z.number().int().min(1).max(100).default(50)
  }, async ({ userId, start, end, top }) => jsonText(await graph(env, "GET", `users/${encodeURIComponent(userId)}/calendarView?startDateTime=${encodeURIComponent(start)}&endDateTime=${encodeURIComponent(end)}&$top=${top}&$select=id,subject,start,end,organizer,location`)));

  return server;
}

export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    if (new URL(request.url).pathname === "/health") return Response.json({ status: "ok", service: "kedleston-m365-graph-mcp", writeEnabled: env.WRITE_ENABLED === "true" });
    const supplied = request.headers.get("authorization")?.replace(/^Bearer\s+/i, "");
    if (!env.MCP_API_KEY || supplied !== env.MCP_API_KEY) return new Response("Unauthorized", { status: 401, headers: { "WWW-Authenticate": "Bearer" } });
    const server = buildServer(env);
    return createMcpHandler(server)(request, env, ctx);
  }
};
