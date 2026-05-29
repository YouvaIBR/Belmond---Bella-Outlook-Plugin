# n8n Workflow — Bella Outlook Webhook

## Architecture

```
[Webhook] → [Code: Validate JWT via JWKS] → [If: JWT valid?] → [HTTP Request: Bella] → [Respond 200]
                                                            └── false → [Respond 401]
```

The webhook accepts all requests without built-in auth. A Code node fetches Microsoft's JWKS endpoint dynamically to find the public key matching the token's `kid`, then verifies the signature using the Web Crypto API. This handles key rotation automatically — no manual PEM updates needed.

> **Why not use n8n's built-in JWT auth?** It only supports a single static PEM key. Microsoft publishes 6 rotating keys and can use any of them to sign a token. A static key causes intermittent 403 errors when Microsoft switches to a different key.

---

## Step 1 — Expose an API scope in Azure

Before generating a token with the correct audience, you need to expose a scope on the app registration.

1. **Entra ID** → **App registrations** → `bella-outlook-addin`
2. **Expose an API** → if no Application ID URI is set, click **Set** and accept `api://YOUR_CLIENT_ID`
3. **Add a scope**:
   - **Scope name**: `access_as_application`
   - **Who can consent**: `Admins only`
   - **Admin consent display name**: `Access Bella`
   - **State**: `Enabled`

---

## Step 2 — Import the workflow

Import `Bella Outlook Plugin.json` into n8n. The workflow includes a Code node that validates the JWT dynamically — no PEM key setup required. The JWKS endpoint is called at runtime to find the correct key for each token.

---

## Step 3 — Configure the Webhook node

- **HTTP Method**: `POST`
- **Path**: `bella-outlook`
- **Authentication**: `None` (JWT validation is handled by the Code node)
- **Response Mode**: `Using 'Respond to Webhook' Node`
- **Allowed Origins (CORS)**: `https://youvaibr.github.io`

> **CORS is required**: the add-in runs on `https://youvaibr.github.io` and browsers block cross-origin requests unless the webhook explicitly allows that origin. Without this setting, the request reaches n8n but the browser discards the response (`ERR_FAILED`).

---

## Step 5 — HTTP Request node (call Bella)

- **Method**: `POST`
- **URL**: `https://n8n-985953980051.europe-west1.run.app/webhook/bella/concierge`
- **Authentication**: `Generic Credential Type` → `Header Auth`
  - Name: `Authorization`
  - Value: `Basic TOKEN==`
- **Body** (JSON):
```json
{
  "query": "={{ $json.body.emailSubject + ' — ' + $json.body.emailBody }}",
  "product_code": "={{ $json.body.productCode }}",
  "vision": false
}
```

---

## Step 6 — Respond to Webhook node

- **Response Code**: `200`
- **Respond With**: `JSON`
- **Body**: `={{ JSON.stringify($json) }}`

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

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `Invalid signature` (intermittent 403) | Was using static PEM — Microsoft rotated key | Use the JWKS-based workflow (dynamic key lookup, no manual PEM needed) |
| `Key not found for kid: ...` | Token signed with unknown key | Microsoft occasionally publishes new keys before retiring old ones — wait a few minutes and retry |
| `Token expired` | JWT older than ~1h | MSAL should refresh automatically; if not, re-authenticate |
| `Invalid audience` | Token `aud` doesn't match CLIENT_ID | Check `VITE_CLIENT_ID` matches the Azure app registration |
| `401` on Bella call | Basic auth token wrong | Check the `Authorization` header value in the HTTP Request node |
| `ERR_FAILED` / CORS error | Missing `Access-Control-Allow-Origin` header | Set **Allowed Origins** to `https://youvaibr.github.io` in the Webhook node options |
