# Weekly Status Agent

A single-page web app that calls the **Microsoft Work IQ API** directly from the
browser, intended to generate a weekly status report from Microsoft 365 data.

Built as a learning exercise: no backend, no Azure hosting, no bundler — just
static files. It also documents two undocumented behaviours of the Work IQ
preview API (see **Findings**).

> **Status: does not work end to end.** Authentication succeeds and the request
> is correctly formed, but the Work IQ service rejects calls from custom app
> registrations. See [Findings](#findings). The repository is kept as a
> reference for the auth flow and the A2A request format.

## Architecture

```
Browser (this app)
  ├─ MSAL.js ──────────> login.microsoftonline.com   (Entra sign-in, PKCE)
  └─ fetch(POST) ──────> workiq.svc.cloud.microsoft/a2a/<agent-id>/
                              │
                              └─ on-behalf-of ──> Microsoft Graph
                                                  (mail, calendar, files, chat)
```

There is no server. A static host serves the files; the access token and all
Microsoft 365 data stay in the browser.

## Files

| File | Purpose |
| --- | --- |
| `index.html` | UI and event wiring |
| `config.js` | Public configuration with placeholder identifiers |
| `config.local.js` | Real tenant values — **gitignored**, create your own |
| `auth.js` | MSAL sign-in and token acquisition |
| `workiq.js` | Work IQ A2A client (JSON-RPC 2.0) |
| `msal-browser.min.js` | MSAL 3.28.1, vendored locally |

## Setup

1. Copy the placeholder config and fill in your own values:

   ```js
   // config.local.js  (gitignored)
   CONFIG.clientId = "<your-entra-app-client-id>";
   CONFIG.tenantId = "<your-tenant-id>";
   CONFIG.workIqScope = `${CONFIG.workIqResourceAppId}/WorkIQAgent.Ask`;
   ```

2. Register a **single-page application** in Microsoft Entra ID with redirect
   URIs for your deployed origin and `http://localhost:5500/`.

3. Provision the Work IQ service principals — they are not present in a tenant
   by default:

   ```powershell
   az ad sp create --id fdcc1f02-fc51-4226-8753-f668596af7f7
   ```

   Then grant tenant-wide consent for the Work IQ CLI app:

   ```text
   https://login.microsoftonline.com/<tenant-id>/adminconsent?client_id=ba081686-5d24-4bc6-a0d6-d034ecffed87
   ```

4. Grant your app these delegated permissions and admin-consent them:

   - Work IQ: `WorkIQAgent.Ask`
   - Microsoft Graph: `Mail.Read`, `Sites.Read.All`, `People.Read.All`,
     `Chat.Read`, `ChannelMessage.Read.All`,
     `OnlineMeetingTranscript.Read.All`, `ExternalItem.Read.All`

   Work IQ uses the on-behalf-of flow: it exchanges the caller's token for a
   Microsoft Graph token to read data as the signed-in user. Entra only issues
   that second token if the downstream scopes were consented **for the calling
   client app**. Partial consent is rejected — the chain is validated as a set.

## Running locally

```powershell
npx http-server -p 5500 -c-1
```

Open `http://localhost:5500/`. The port must match a registered redirect URI.
MSAL cannot run from a `file://` URL, because OAuth requires an origin.

## Findings

Two behaviours of the Work IQ preview API that the documentation does not cover,
both established by capturing the traffic of Microsoft's own Work IQ CLI.

### 1. The agent id belongs in the URL path

The published example shows a POST to the bare `/a2a/` root. In practice the
CLI addresses a specific agent by appending its id to the path:

```http
POST https://workiq.svc.cloud.microsoft/a2a/bizchat-as-gpt-scenario/
Authorization: Bearer <token>
A2A-Version: 1.0
User-Agent: WorkIQ-SDK/1.0

{
  "jsonrpc": "2.0",
  "id": "<guid>",
  "method": "SendMessage",
  "params": {
    "message": {
      "role": "ROLE_USER",
      "messageId": "<guid>",
      "parts": [{ "text": "Summarise my recent emails" }],
      "metadata": {
        "Location": { "TimeZone": "Asia/Kolkata", "CountryOrRegion": "IN" }
      }
    }
  }
}
```

Posting to `/a2a/` without an agent id returns
`400 AuthenticationError: "Error authenticating with resource"`, which is
misleading — nothing is wrong with the token.

Agent ids can be listed with `workiq agents list`. Declarative agents appear as
A2A-addressable endpoints with ids of the form `P_<guid>.declarativeAgent`.

Multi-turn conversations thread by passing the previous response's `contextId`.
Omitting the `A2A-Version` header defaults to A2A v0.3.

### 2. Work IQ requires client preauthorization

Even with a correctly formed request, a valid token (`aud` = the Work IQ
resource, `scp` = `WorkIQAgent.Ask`), full admin consent for every required
scope, and the agent id in the path, calls from a custom app registration are
still rejected with `400 AuthenticationError`.

Requesting a Work IQ token from a different client makes the reason explicit:

```text
AADSTS65002: Consent between first party application '04b07795-…' and first
party resource 'fdcc1f02-…' must be configured via preauthorization —
applications owned and operated by Microsoft must get approval from the API
owner before requesting tokens for that API.
```

Work IQ maintains an allow-list of preauthorized client application ids.
Microsoft's Work IQ CLI (`ba081686-…`) is on it; Azure CLI and custom app
registrations are not. Entra will issue a token to a consented third-party app,
but the service then rejects the `appid` claim.

This is not a configuration problem and cannot be resolved from the tenant side.
It is consistent with the preview
[Terms of Use](https://learn.microsoft.com/en-us/legal/work-iq-apis/terms-of-use),
which state that Microsoft sets and enforces limits on API use at its sole
discretion.

**For custom tooling today, use the Work IQ CLI as a local MCP stdio server
(`workiq mcp`)**, which runs under Microsoft's preauthorized client id.

## Licensing

Work IQ API usage is **metered and billed via Copilot Credits**, unlike
declarative agents used inside Microsoft channels (Teams, Outlook, the Microsoft
365 Copilot app), which are included in the Microsoft 365 Copilot licence.
Usage-based billing is configured in the Microsoft 365 admin center under
**Copilot > Cost management**.

See the [Work IQ API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/api-overview).

## A note on secrets

There are no secrets in this repository, by design. A single-page app cannot
keep a secret — its JavaScript is fully readable — which is why it uses the
authorization code flow with PKCE rather than a client secret. Security comes
from the Entra redirect-URI allow-list and user sign-in, not from concealment.

Tenant and client identifiers are kept in `config.local.js`, which is gitignored,
so the published configuration contains placeholders only.
