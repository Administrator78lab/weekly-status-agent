// Minimal Work IQ A2A client.
// Protocol: JSON-RPC 2.0 over HTTPS POST, method name in the body not the URL.
// The agent id goes in the URL PATH - posting to the bare /a2a/ root is rejected
// with 400 AuthenticationError. Captured from Microsoft's own CLI.

const DEFAULT_AGENT = "bizchat-as-gpt-scenario";

async function askWorkIq(question, contextId = null, agentId = DEFAULT_AGENT) {
  const token = await getWorkIqToken();

  const message = {
    role: "ROLE_USER",
    messageId: crypto.randomUUID(),
    parts: [{ text: question }],
    metadata: {
      Location: {
        TimeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        CountryOrRegion: "IN",
      },
    },
  };

  // contextId threads follow-up turns onto the same conversation.
  if (contextId) message.contextId = contextId;

  const body = {
    jsonrpc: "2.0",
    id: crypto.randomUUID(),
    method: "SendMessage",
    params: { message },
  };

  const url = CONFIG.a2aEndpoint + agentId + "/";

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json",
      "A2A-Version": "1.0",
    },
    body: JSON.stringify(body),
  });

  const requestId = response.headers.get("request-id");
  const json = await response.json();

  if (!response.ok) {
    throw new Error(`HTTP ${response.status} (request-id ${requestId}): ${JSON.stringify(json, null, 2)}`);
  }
  if (json.error) {
    throw new Error(`A2A error: ${JSON.stringify(json.error, null, 2)}`);
  }

  const task = json.result?.task;
  const text = task?.artifacts
    ?.flatMap(a => a.parts ?? [])
    .map(p => p.text)
    .filter(Boolean)
    .join("\n\n");

  return { text, contextId: task?.contextId, requestId, raw: json };
}

// Lists agents available to the signed-in user. Useful for discovering agent
// ids to put in the A2A URL path, including declarative agents you publish.
async function listWorkIqAgents() {
  const token = await getWorkIqToken();
  const url = CONFIG.a2aEndpoint.replace(/\/a2a\/$/, "/a2a/agents");

  const response = await fetch(url, {
    headers: { "Authorization": `Bearer ${token}`, "A2A-Version": "1.0" },
  });

  const requestId = response.headers.get("request-id");
  const text = await response.text();
  let out;
  try { out = JSON.stringify(JSON.parse(text), null, 2); } catch { out = text; }
  return `HTTP ${response.status} (request-id ${requestId})\n\n${out}`;
}
