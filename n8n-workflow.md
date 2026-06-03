# n8n Workflow — Bella Outlook Webhook

## Architecture

```
[Webhook OPTIONS] → [Respond OPTIONS 204]   ← CORS preflight

[Webhook POST] → [Fetch JWKS] → [Validate JWT (JWKS)] → [JWT Valid?] → true  → [HTTP Request: Bella] → [Respond 200]
                                                                       └── false → [Reject 401/403]
```

The webhook accepts all requests without built-in auth. A dedicated `Webhook OPTIONS` node handles CORS preflight (browsers send this before every cross-origin POST). The main `Webhook POST` node fetches Microsoft's JWKS endpoint dynamically to find the public key matching the token's `kid`, then verifies the signature using the Web Crypto API. This handles key rotation automatically — no manual PEM updates needed.

> **Why not use n8n's built-in JWT auth?** It only supports a single static PEM key. Microsoft publishes 6 rotating keys and can use any of them to sign a token. A static key causes intermittent 403 errors when Microsoft switches to a different key.

---

## Step 1 — Expose an API scope in Azure

Before generating a token with the correct audience, you need to expose a scope on the app registration.

1. **Entra ID** → **App registrations** → `bella-outlook-addin`
2. **Expose an API** → if no Application ID URI is set, click **Set** and accept `api://<YOUR_CLIENT_ID>`
3. **Add a scope**:
   - **Scope name**: `access_as_application`
   - **Who can consent**: `Admins only`
   - **Admin consent display name**: `Access Bella`
   - **State**: `Enabled`

---

## Step 2 — Import the workflow

Import `Bella Outlook Plugin.json` into n8n. The workflow includes a Code node that validates the JWT dynamically — no PEM key setup required. The JWKS endpoint is called at runtime to find the correct key for each token.

After importing, update the hardcoded `CLIENT_ID` in the **Validate JWT (JWKS)** Code node to match your Azure `VITE_CLIENT_ID`.

---

## Step 3 — Configure the Webhook POST node

- **HTTP Method**: `POST`
- **Path**: `bella-outlook`
- **Authentication**: `None` (JWT validation is handled by the Code node)
- **Response Mode**: `Using 'Respond to Webhook' Node`
- **Allowed Origins (CORS)**: `https://bel-prd-vision-prj.web.app`

---

## Step 4 — Configure the Webhook OPTIONS node

This node handles CORS preflight requests sent automatically by browsers before cross-origin POSTs.

- **HTTP Method**: `OPTIONS`
- **Path**: `bella-outlook`
- **Authentication**: `None`
- **Response Mode**: `Using 'Respond to Webhook' Node`

Connected to **Respond OPTIONS**, which returns:

- **Response Code**: `204`
- **Headers**:
  - `Access-Control-Allow-Origin`: `https://bel-prd-vision-prj.web.app`
  - `Access-Control-Allow-Methods`: `POST, OPTIONS`
  - `Access-Control-Allow-Headers`: `Authorization, Content-Type`

> **Why this is required**: browsers send a preflight OPTIONS request before every cross-origin POST. Without a node listening on OPTIONS at the same path, the browser never sends the actual request.

---

## Step 5 — HTTP Request node (call Bella)

- **Method**: `POST`
- **URL**: `https://n8n-985953980051.europe-west1.run.app/webhook/bella/concierge`
- **Authentication**: `Generic Credential Type` → `Header Auth` (credential name: `bella n8n auth`)
- **Additional header**:
  - Name: `Cookie`
  - Value: session cookie (stored in the `bella n8n auth` credential)
- **Body** (JSON):

```json
{
  "query": "={{ $('Validate JWT (JWKS)').item.json.body.emailSubject + ' — ' + $('Validate JWT (JWKS)').item.json.body.emailBody }}",
  "product_code": "LRS",
  "vision": false
}
```

> `product_code` is currently hardcoded to `"LRS"`. To make it dynamic, replace with `={{ $('Validate JWT (JWKS)').item.json.body.productCode }}`.

---

## Step 6 — Respond to Webhook node (success)

- **Response Code**: `200`
- **Respond With**: `JSON`
- **Body**: `={{ JSON.stringify($json) }}`
- **Headers**:
  - `Access-Control-Allow-Origin`: `https://bel-prd-vision-prj.web.app`

---

## Step 7 — Reject node (auth failure)

- **Response Code**: `401` (or `403` for audience mismatch)
- **Respond With**: `JSON`
- **Body**: `={{ JSON.stringify({ error: $json.error }) }}`
- **Headers**:
  - `Access-Control-Allow-Origin`: `https://bel-prd-vision-prj.web.app`

---

## Testing

### Generate a token

```bash
curl -X POST https://login.microsoftonline.com/YOUR_TENANT_ID/oauth2/v2.0/token \
  -d "grant_type=client_credentials" \
  -d "client_id=YOUR_CLIENT_ID" \
  -d "client_secret=YOUR_SECRET" \
  -d "scope=api://YOUR_CLIENT_ID/.default"
```

### Send a test request

```bash
curl -X POST https://youvaibrahim.app.n8n.cloud/webhook-test/bella-outlook \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer ACCESS_TOKEN" \
  -d '{
    "emailBody": "Hello, I would like to book a table at your restaurant.",
    "emailSubject": "Restaurant reservation",
    "emailFrom": "client@example.com",
    "productCode": "LRS"
  }'
```

Use `/webhook-test/` during development (workflow does not need to be active). Use `/webhook/` in production (workflow must be active).

---

## Updating the CORS origin (Firebase migration)

All three `Respond` nodes have a hardcoded `Access-Control-Allow-Origin` header, as does the `Webhook POST` node's **Allowed Origins** field.

---

## Troubleshooting

| Error                                  | Cause                                          | Fix                                                                                                         |
| -------------------------------------- | ---------------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `Invalid signature` (intermittent 403) | Was using static PEM — Microsoft rotated key   | Use the JWKS-based workflow (dynamic key lookup, no manual PEM needed)                                      |
| `Key not found for kid: ...`           | Token signed with unknown key                  | Microsoft occasionally publishes new keys before retiring old ones — wait a few minutes and retry           |
| `Token expired`                        | JWT older than ~1h                             | MSAL should refresh automatically; if not, re-authenticate                                                  |
| `Invalid audience`                     | Token `aud` doesn't match CLIENT_ID            | Check the `CLIENT_ID` constant in the **Validate JWT (JWKS)** Code node matches your Azure app registration |
| `401` on Bella call                    | Cookie credential expired                      | Update the `Cookie` header value in the **HTTP Request** node                                               |
| `ERR_FAILED` / CORS error in browser   | Missing or wrong `Access-Control-Allow-Origin` | Update all CORS headers and Allowed Origins to the current hosting domain                                   |
| Preflight blocked                      | OPTIONS node missing or path mismatch          | Ensure `Webhook OPTIONS` uses the exact same path as `Webhook POST` (`bella-outlook`)                       |
