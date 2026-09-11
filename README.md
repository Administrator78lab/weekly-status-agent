# Weekly Status Agent

A static single-page web app that reads Microsoft 365 data by calling the
**Microsoft Work IQ API** directly from the browser — no backend, no Azure
hosting, no build step.

It talks to Work IQ over the **remote MCP (Model Context Protocol)** endpoint,
which works from browser JavaScript because the service sends
`Access-Control-Allow-Origin: *`.

> **Working:** Entra sign-in, token acquisition, MCP handshake, `tools/list`,
> and the `fetch` tool returning real mailbox data.
> **Not working:** the `ask` tool returns a service-side internal error
> (see [Findings](#findings)).

## How it works

```
Browser (this app)
  ├─ MSAL.js ─────────> login.microsoftonline.com     (Entra sign-in, PKCE)
  └─ fetch(POST) ─────> workiq.svc.cloud.microsoft/mcp
                             │
                             └─ on-behalf-of ──> Microsoft Graph
                                                 (mail, calendar, files, chat)
```

There is no server component. A static host serves the files; the access token
and all Microsoft 365 data stay in the browser.

## Files

| File | Purpose |
| --- | --- |
| `index.html` | UI and event wiring |
| `config.js` | Public configuration, placeholder identifiers |
| `config.local.js` | Real tenant values — **gitignored**, create your own |
| `auth.js` | MSAL sign-in and token acquisition |
| `workiq-mcp.js` | Work IQ MCP client (Streamable HTTP) |
| `workiq.js` | Work IQ A2A client (kept for comparison, see Findings) |
| `msal-browser.min.js` | MSAL 3.28.1, vendored locally |

## Setup

### 1. Register the app in Microsoft Entra ID

Create an app registration with platform **Single-page application** and these
redirect URIs:

```
http://localhost:5500/
https://<your-account>.github.io/<your-repo>/
```

Single tenant (`AzureADMyOrg`) is sufficient.

### 2. Provision Work IQ in the tenant

Work IQ has no service principal in a tenant by default, so it does not appear
under *APIs my organization uses* until you create one:

```powershell
az login --tenant <tenant-id> --allow-no-subscriptions
az ad sp create --id fdcc1f02-fc51-4226-8753-f668596af7f7
```

Then grant tenant-wide consent for the Work IQ CLI app:

```
https://login.microsoftonline.com/<tenant-id>/adminconsent?client_id=ba081686-5d24-4bc6-a0d6-d034ecffed87
```

### 3. Grant API permissions

Add these **delegated** permissions to your app registration and click
**Grant admin consent**:

- Work IQ: `WorkIQAgent.Ask`
- Microsoft Graph: `Mail.Read`, `Sites.Read.All`, `People.Read.All`,
  `Chat.Read`, `ChannelMessage.Read.All`,
  `OnlineMeetingTranscript.Read.All`, `ExternalItem.Read.All`

The Graph permissions are required because Work IQ uses the **on-behalf-of**
flow: it exchanges your token for a Graph token to read data as you. Entra only
issues that second token if the downstream scopes were consented *for the
calling client app*. Partial consent is rejected — the set is validated as a
whole.

### 4. Local configuration

```js
// config.local.js  (gitignored)
CONFIG.clientId = "<your-app-client-id>";
CONFIG.tenantId = "<your-tenant-id>";
CONFIG.workIqScope = `${CONFIG.workIqResourceAppId}/WorkIQAgent.Ask`;
```

### 5. Run

```powershell
npx http-server -p 5500 -c-1
```

Open `http://localhost:5500/`. The port must match the registered redirect URI.
MSAL cannot run from a `file://` URL, because OAuth requires an origin.

## Using the MCP endpoint

MCP is stateful: `initialize` first, then other calls.

```http
POST https://workiq.svc.cloud.microsoft/mcp
Authorization: Bearer <token>
Content-Type: application/json
Accept: application/json, text/event-stream
MCP-Protocol-Version: 2025-06-18

{"jsonrpc":"2.0","id":"1","method":"initialize",
 "params":{"protocolVersion":"2025-06-18","capabilities":{},
           "clientInfo":{"name":"weekly-status-agent","version":"1.0.0"}}}
```

The server replies with its own info (`WorkIQ.MCP.Server`) and may return a
`Mcp-Session-Id` header to send on later requests. Responses can be plain JSON
or an SSE stream, so clients must handle both content types.

### Tools exposed

`tools/list` returns 11 tools:

| Tool | Purpose |
| --- | --- |
| `ask` | Natural-language question over M365 data; supports `agentId`, `conversationId`, `fileUrls`, `timeZone` |
| `fetch` | Read entities by Graph-style path, e.g. `/me/messages?$select=subject,from&$top=10` |
| `fetch_blob` | Binary content up to 4 MB, base64 encoded |
| `search_paths` | Discover available entity paths |
| `get_schema` | OpenAPI schema for an operation |
| `call_function` | GET-style functions such as `delta`, `reminderView` |
| `create_entity` / `update_entity` / `delete_entity` | Write operations |
| `do_action` | Actions such as `/me/sendMail`, copy, move |
| `list_agents` | Copilot agent ids for `ask`'s `agentId` parameter |

Example structured read:

```json
{"jsonrpc":"2.0","id":"2","method":"tools/call",
 "params":{"name":"fetch",
           "arguments":{"entityUrls":["/me/messages?$select=subject,from&$top=10"]}}}
```

## Findings

Notes from getting this working, since the documentation does not cover them.

### MCP works from a custom app; A2A did not

The project initially targeted the **A2A** endpoint, following the documented
example. Every call returned:

```
400 AuthenticationError: "Error authenticating with resource"
```

That error persisted through every documented remedy: full admin consent for all
seven Graph scopes, multitenant `signInAudience`, the `/common` authority, and
appending the agent id to the path. The token itself was always valid — correct
`aud` and `scp` claims.

Switching the **same token and the same app registration** to the remote MCP
endpoint worked immediately. So the failure was specific to how the A2A endpoint
was being addressed, not to permissions and not to the client application
identity.

An earlier version of this README concluded that Work IQ only accepts
Microsoft's own preauthorized client applications. **That conclusion was wrong**
and has been removed. The `AADSTS65002` preauthorization error that suggested it
came from requesting a Work IQ token with the *Azure CLI*, which is a
first-party-to-first-party restriction and does not apply to third-party app
registrations.

### The A2A endpoint expects an agent id in the path

Captured from Microsoft's Work IQ CLI by pointing its configurable endpoint at a
local listener:

```http
POST https://workiq.svc.cloud.microsoft/a2a/bizchat-as-gpt-scenario/
A2A-Version: 1.0
User-Agent: WorkIQ-SDK/1.0

{"jsonrpc":"2.0","id":"<guid>","method":"SendMessage",
 "params":{"message":{"role":"ROLE_USER","messageId":"<guid>",
           "parts":[{"text":"..."}],
           "metadata":{"Location":{"TimeZone":"Asia/Kolkata",
                                   "CountryOrRegion":"IN"}}}}}
```

The published example posts to the bare `/a2a/` root. Adding the agent id did
not by itself resolve the 400 in this project, but it is what the working client
sends.

### The `ask` tool currently fails server-side

`tools/call` with `ask` is accepted, allocates a `conversationId`, and then
returns:

```json
{"response":null,"conversationId":"…",
 "error":"An internal error occurred while asking question. Please try again later."}
```

This happens with only the required `question` argument, so it is not a
malformed request. The `fetch` tool, which reads Microsoft Graph without
involving the Copilot model, works correctly against the same session — which
isolates the fault to Work IQ's answering backend.

## Licensing

Work IQ API usage is **metered and billed via Copilot Credits**, unlike
declarative agents used inside Microsoft channels (Teams, Outlook, the Microsoft
365 Copilot app), which are included in the Microsoft 365 Copilot licence. Set
up usage-based billing in the Microsoft 365 admin center under
**Copilot > Cost management**.

See the [Work IQ API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/api-overview).

## A note on secrets

There are none in this repository, by design. A single-page app cannot keep a
secret — its JavaScript is readable by anyone — which is why it uses the
authorization code flow with PKCE rather than a client secret. Security comes
from the Entra redirect-URI allow-list and user sign-in, not from concealment.

Tenant and client identifiers live in `config.local.js`, which is gitignored, so
the published configuration contains placeholders only. Clone this repo and it
will load but not sign in until you supply your own values.
