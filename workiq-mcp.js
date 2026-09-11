// Work IQ remote MCP client (Streamable HTTP transport).
//
// Differences from the A2A client:
//   - requires an `initialize` handshake before any other call
//   - session is carried in the Mcp-Session-Id response header
//   - responses may be JSON or an SSE stream (text/event-stream)
//   - tools are discovered with tools/list and invoked with tools/call

const MCP_PROTOCOL_VERSION = "2025-06-18";

let mcpSessionId = null;

function mcpUrl() {
  return CONFIG.mcpEndpoint || "https://workiq.svc.cloud.microsoft/mcp";
}

// The endpoint may answer with plain JSON or an SSE stream. SSE frames look
// like "event: message\ndata: {...}\n\n", so pull the last data: line out.
function parseMcpBody(contentType, text) {
  if (contentType && contentType.includes("text/event-stream")) {
    const dataLines = text
      .split(/\r?\n/)
      .filter(l => l.startsWith("data:"))
      .map(l => l.slice(5).trim())
      .filter(Boolean);
    const last = dataLines[dataLines.length - 1];
    if (!last) return { _raw: text };
    try { return JSON.parse(last); } catch { return { _raw: text }; }
  }
  try { return JSON.parse(text); } catch { return { _raw: text }; }
}

async function mcpRpc(method, params = undefined, { notification = false } = {}) {
  const token = await getWorkIqToken();

  const body = { jsonrpc: "2.0", method };
  if (params !== undefined) body.params = params;
  if (!notification) body.id = crypto.randomUUID();

  const headers = {
    "Authorization": `Bearer ${token}`,
    "Content-Type": "application/json",
    "Accept": "application/json, text/event-stream",
    "MCP-Protocol-Version": MCP_PROTOCOL_VERSION,
  };
  if (mcpSessionId) headers["Mcp-Session-Id"] = mcpSessionId;

  const response = await fetch(mcpUrl(), {
    method: "POST",
    headers,
    body: JSON.stringify(body),
  });

  // The server assigns a session id on initialize; reuse it from then on.
  const sid = response.headers.get("mcp-session-id");
  if (sid) mcpSessionId = sid;

  const requestId = response.headers.get("request-id");
  const contentType = response.headers.get("content-type") || "";
  const text = await response.text();

  if (notification) return { status: response.status, requestId };

  const json = parseMcpBody(contentType, text);

  if (!response.ok) {
    throw new Error(
      `HTTP ${response.status} (request-id ${requestId})\n\n` +
      JSON.stringify(json, null, 2)
    );
  }
  if (json.error) {
    throw new Error(`MCP error (request-id ${requestId}):\n\n` +
      JSON.stringify(json.error, null, 2));
  }

  return { result: json.result, requestId, status: response.status };
}

// Handshake. Must succeed before tools/list or tools/call.
async function mcpInitialize() {
  mcpSessionId = null;

  const { result, requestId, status } = await mcpRpc("initialize", {
    protocolVersion: MCP_PROTOCOL_VERSION,
    capabilities: {},
    clientInfo: { name: "weekly-status-agent", version: "1.0.0" },
  });

  // Tell the server the handshake is complete (notification, no response body).
  try { await mcpRpc("notifications/initialized", {}, { notification: true }); }
  catch { /* non-fatal */ }

  return { result, requestId, status, sessionId: mcpSessionId };
}

async function mcpListTools() {
  const { result, requestId } = await mcpRpc("tools/list", {});
  return { tools: result?.tools ?? [], requestId };
}

async function mcpCallTool(name, args = {}) {
  const { result, requestId } = await mcpRpc("tools/call", {
    name,
    arguments: args,
  });

  const text = (result?.content ?? [])
    .filter(c => c.type === "text")
    .map(c => c.text)
    .join("\n\n");

  return { text, raw: result, requestId };
}
