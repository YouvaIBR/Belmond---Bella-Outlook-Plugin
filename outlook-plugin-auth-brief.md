# Outlook Plugin — Entra ID JWT Auth Strategy
## Implementation Brief for Claude Code

---

## Overview

Secure an Outlook add-in that calls a n8n webhook (triggering a MCP agent), without storing any credentials in the plugin code or manifest. Authentication is handled via Microsoft Entra ID: the plugin retrieves a signed JWT token from the logged-in user's Microsoft 365 session and sends it with every request. n8n verifies the token before passing the request to the MCP agent.

---

## Architecture

```
Outlook Plugin (MSAL)
      |
      | Bearer JWT token (Entra ID signed)
      v
n8n Webhook
      |
      | Verify signature via Microsoft JWKS endpoint
      |--- invalid --> 401 Unauthorized
      |--- valid   -->
      v
MCP Agent (n8n workflow)
      |
      v
Response --> Outlook Plugin
```

---

## Stack

| Layer | Technology |
|---|---|
| Plugin framework | Office.js (Outlook Add-in) |
| Auth library | `@azure/msal-browser` (NAA — Nested App Authentication) |
| Identity provider | Microsoft Entra ID (formerly Azure AD) |
| Backend / orchestration | n8n (self-hosted or cloud) |
| Token verification | n8n built-in JWT credential (RS256 + PEM public key) |
| Agent | MCP agent node in n8n |

---

## What needs to be set up

### 1. Azure — App Registration

In **Entra ID → App registrations → New registration**:

- Name: `outlook-plugin-mcp` (or similar)
- Supported account types: single tenant or multitenant depending on client
- Redirect URI: add `brk-multihub://auth` (required for NAA/SSO in Office)
- For Outlook web fallback, also add an SPA redirect URI pointing to the add-in's HTML page

In **API permissions**, add:
- `openid`
- `profile`
- `email`

Note down:
- `tenant_id`
- `client_id`

No client secret needed — the plugin is a public client.

---

### 2. Outlook Plugin — MSAL integration

Install the dependency:

```bash
npm install @azure/msal-browser
```

Initialize MSAL with NAA (Nested App Authentication):

```javascript
import { createNestablePublicClientApplication } from "@azure/msal-browser";

const msalInstance = await createNestablePublicClientApplication({
  auth: {
    clientId: "YOUR_CLIENT_ID",
    authority: "https://login.microsoftonline.com/YOUR_TENANT_ID"
  }
});
```

Retrieve the token before each request:

```javascript
async function getToken() {
  const account = msalInstance.getAllAccounts()[0];

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: ["openid", "profile", "email"],
      account
    });
    return result.accessToken;
  } catch (e) {
    // Fallback: prompt user to sign in
    const result = await msalInstance.acquireTokenPopup({
      scopes: ["openid", "profile", "email"]
    });
    return result.accessToken;
  }
}
```

Send the token with every request to n8n:

```javascript
async function callAgent(payload) {
  const token = await getToken();

  const response = await fetch("https://your-n8n-instance.com/webhook/my-agent", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  return response.json();
}
```

---

### 3. n8n — Webhook + JWT verification

**Option A — built-in JWT auth (simplest)**

In n8n, go to **Credentials → Add → JWT**:
- Key type: `PEM Key`
- Public key: paste Entra ID's public key in PEM format (exported from the JWKS endpoint at `https://login.microsoftonline.com/{tenant_id}/discovery/v2.0/keys`)
- Algorithm: `RS256`

In the **Webhook node**, set:
- Authentication: `JWT`
- Credential: the JWT credential created above

n8n will automatically reject requests with an invalid or expired token.

> ⚠️ Important: n8n's built-in JWT auth verifies the signature but does NOT automatically enforce claims like `aud` (audience) or `iss` (issuer). Add a Code node after the webhook to check them explicitly (see below).

**Option B — manual verification via Code node (more control)**

Add a **Code node** immediately after the Webhook node:

```javascript
const jwksClient = require('jwks-rsa');
const jwt = require('jsonwebtoken');

const token = $input.first().json.headers.authorization?.replace('Bearer ', '');
if (!token) throw new Error('Missing token');

const client = jwksClient({
  jwksUri: `https://login.microsoftonline.com/YOUR_TENANT_ID/discovery/v2.0/keys`
});

const decoded = jwt.decode(token, { complete: true });
const key = await client.getSigningKey(decoded.header.kid);

const verified = jwt.verify(token, key.getPublicKey(), {
  algorithms: ['RS256'],
  audience: 'YOUR_CLIENT_ID',
  issuer: `https://login.microsoftonline.com/YOUR_TENANT_ID/v2.0`
});

return [{ json: { valid: true, user: verified.preferred_username, userId: verified.oid } }];
```

Add an **If node** after:
- If `valid === true` → continue to MCP agent
- Else → **Respond to Webhook** with status `401`

---

### 4. n8n — MCP Agent node

Connect the verified output to the existing MCP agent node. The `user` and `userId` fields from the verified token are available as context if needed (e.g. for logging or personalisation).

---

## Token lifecycle — nothing to manage

| Event | What happens |
|---|---|
| Token valid (~1h) | `acquireTokenSilent` returns cached token instantly |
| Token expired | MSAL silently fetches a new one using the refresh token |
| Refresh token expired (~90 days of inactivity) | `acquireTokenSilent` throws → fallback to `acquireTokenPopup` |
| Admin revokes session | Next silent renewal fails → user must sign in again |

---

## Known limitations

**Outlook Mobile**: NAA (`acquireTokenSilent`) currently fails on Outlook mobile due to a known Microsoft bug. `acquireTokenPopup` also fails on mobile. Desktop and Outlook Web work correctly. If mobile support is required, a fallback strategy (e.g. a dedicated login screen) will need to be discussed.

**New Outlook on Windows + corporate proxy / Okta**: Some tenants using Okta or authentication proxies experience NAA failures in New Outlook for Windows. Classic Outlook and OWA are unaffected. Test on the client's environment early.

**Conditional Access policies**: The "approved client app" CA grant is deprecated and will be retired in March 2026. If the client's tenant uses this policy, it must be migrated to the Application Protection Policy grant for NAA to work.

---

## CLAUDE.md for this project

Create a `CLAUDE.md` at the root of the plugin project with the following content:

```markdown
# Outlook Plugin — Auth context

## What this project does
Outlook add-in that calls a n8n webhook secured via Microsoft Entra ID JWT.
No credentials are stored in the code or manifest.

## Auth flow
1. MSAL (NAA) retrieves a JWT from the logged-in Microsoft 365 user silently
2. JWT sent as Bearer header to n8n webhook
3. n8n verifies signature + claims (aud, iss) before passing to MCP agent

## Key files
- `src/auth/msal.js` — MSAL initialisation and token retrieval
- `src/api/agent.js` — fetch wrapper that attaches the Bearer token
- `manifest.xml` — Office add-in manifest (do not add secrets here)

## Environment variables (never hardcoded)
- VITE_CLIENT_ID — Entra ID app client ID
- VITE_TENANT_ID — Entra ID tenant ID
- VITE_N8N_WEBHOOK_URL — n8n webhook endpoint

## Constraints
- Never hardcode credentials, tokens, or secrets anywhere in the codebase
- Always use acquireTokenSilent first, fall back to acquireTokenPopup on error
- MSAL must be initialised with createNestablePublicClientApplication (NAA), not PublicClientApplication
- The Webhook URL must come from an env variable, never be inlined

## Testing
- Unit tests: Jest
- Auth flow can be tested against a dev App Registration in Entra ID
- n8n webhook URL for dev is set in .env.local (gitignored)
```

---

## Checklist before handoff

- [ ] App Registration created in Entra ID with correct redirect URIs
- [ ] `tenant_id` and `client_id` stored in environment variables
- [ ] MSAL initialised with `createNestablePublicClientApplication`
- [ ] `acquireTokenSilent` + `acquireTokenPopup` fallback implemented
- [ ] n8n JWT credential configured with Entra ID public key (RS256)
- [ ] Claims (`aud`, `iss`) verified in n8n Code node or built-in auth
- [ ] 401 response wired up for invalid tokens
- [ ] Tested on Outlook Desktop and Outlook Web
- [ ] Mobile limitation documented and flagged to client
