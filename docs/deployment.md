# Deployment Guide — Bella Outlook Add-in

This guide covers everything needed to deploy the Bella Outlook add-in from scratch: Azure configuration, GitLab setup, and installing the add-in in Outlook.

---

## Overview

```
GitLab repo (source)
      │
      │  git push → GitLab CI/CD
      ▼
Firebase Hosting (https://<FIREBASE_DOMAIN>)
      │
      │  manifest.xml references this URL
      ▼
Microsoft 365 Admin Center (add-in deployment)
      │
      ▼
Outlook (Desktop + Web)
```

---

## Part 1 — Azure App Registration

### 1.1 Create the registration

1. Go to [portal.azure.com](https://portal.azure.com)
2. **Entra ID** → **App registrations** → **New registration**
3. Fill in:
   - **Name**: `bella-outlook-addin`
   - **Supported account types**: Single tenant
   - **Redirect URI**: Single-page application (SPA) → `brk-multihub://<FIREBASE_DOMAIN>`
4. Click **Register**

### 1.2 Note the IDs

On the **Overview** page, copy:

- **Application (client) ID** → used as `VITE_CLIENT_ID`
- **Directory (tenant) ID** → used as `VITE_TENANT_ID`

### 1.3 Add redirect URIs

**Authentication** → existing SPA platform → **Add URI**, add all of these:

| URI                                          | Purpose                                       |
| -------------------------------------------- | --------------------------------------------- |
| `brk-multihub://<FIREBASE_DOMAIN>`           | NAA broker (standard)                         |
| `brk-<TENANT_BROKER_ID>://<FIREBASE_DOMAIN>` | NAA broker (tenant-specific — see note below) |
| `https://<FIREBASE_DOMAIN>/index.html`       | Outlook Web                                   |

> **Important — Tenant-specific broker ID**: Do not add this URI yet — you won't have the `TENANT_BROKER_ID` until after the first authentication attempt. The flow is: deploy the add-in → open it in Outlook → if you get the error `Invalid Reply Address. Reply Address must have scheme brk-XXXX://`, copy the `brk-XXXX` value from the failing request URL in DevTools, then come back here and add `brk-XXXX://<FIREBASE_DOMAIN>` as a redirect URI.

> **Set `accessTokenAcceptedVersion` to 2**: In Azure → **Manifest (JSON)**, set `"accessTokenAcceptedVersion": 2`. This is required to allow custom Application ID URIs with external domains (e.g. Firebase Hosting). If the field is not visible in the editor, do a hard refresh (Cmd+Shift+R on Mac, Ctrl+Shift+R on Windows) with cache cleared — Azure Portal sometimes requires a cache clear to display all manifest fields.

### 1.4 Add API permissions

**API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**:

- `openid`
- `profile`
- `email`
- `User.Read`

Then click **Grant admin consent for [your tenant]**.

### 1.5 Expose an API

**Expose an API**:

1. **Application ID URI** — set to:

   ```
   api://<FIREBASE_DOMAIN>/<YOUR_CLIENT_ID>
   ```

   > This requires `accessTokenAcceptedVersion: 2` (see step 1.3). If your tenant policy blocks custom domains, use `api://<YOUR_CLIENT_ID>` and update the scope in `src/auth/msal.ts` accordingly.

2. **Add a scope**:

| Field                      | Value                                             |
| -------------------------- | ------------------------------------------------- |
| Scope name                 | `access_as_user`                                  |
| Who can consent            | `Admins and users`                                |
| Admin consent display name | `Access Bella as user`                            |
| Admin consent description  | `Allows the add-in to access the API as the user` |
| User consent display name  | `Access Bella as user`                            |
| User consent description   | `Allows the add-in to access the API as you`      |
| State                      | `Enabled`                                         |

---

## Part 2 — Firebase Hosting Setup

### 2.1 Create the Firebase project

1. Go to [console.firebase.google.com](https://console.firebase.google.com)
2. Click **Add project** → give it a name (e.g. `bella-belmond`)
3. Once created, go to **Hosting** → **Get started** and follow the steps
4. Note your project ID and hosting domain (e.g. `bella-belmond.web.app`)

### 2.2 Add a `firebase.json` at the project root

```json
{
  "hosting": {
    "public": "dist",
    "ignore": ["firebase.json", "**/.*", "**/node_modules/**"],
    "rewrites": [{ "source": "**", "destination": "/index.html" }]
  }
}
```

### 2.3 Authentication — Workload Identity Federation

CI authentication to Firebase is handled via WIF — no static token or service account key required. GCP is already configured. Contact Daniel González (daniel.gonzalez@belmond.com) for any IAM changes.

| Resource | Value |
| -------- | ----- |
| Service account | `bel-prd-bapi-sac-fb-hosting@bel-prd-bapi-prj.iam.gserviceaccount.com` |
| WIF provider | `projects/835091172967/locations/global/workloadIdentityPools/bel-bapi-shared-pool/providers/bel-bapi-gitlab` |
| Firebase project | `bel-prd-bapi-prj` |

### 2.4 Add GitLab CI/CD variables

**GitLab** → **Settings** → **CI/CD** → **Variables** → **Add variable**:

| Variable name           | Value                                          |
| ----------------------- | ---------------------------------------------- |
| `VITE_CLIENT_ID`        | Azure Application (client) ID                  |
| `VITE_TENANT_ID`        | Azure Directory (tenant) ID                    |
| `VITE_N8N_WEBHOOK_URL`  | n8n webhook URL                                |
| `FIREBASE_PROJECT_ID`   | Firebase project ID (provided by client)       |
| `GCP_WIF_PROVIDER_URL`  | Full WIF provider URL — `https://iam.googleapis.com/...` (provided by client) |
| `GCP_WIF_PROVIDER_PATH` | WIF provider path without scheme (provided by client) |
| `GCP_SERVICE_ACCOUNT`   | GCP service account email (provided by client) |

> No Firebase token or key needed — authentication is handled automatically by WIF in the CI job.

### 2.5 Deploy

Every push to `main` triggers the pipeline automatically:

```bash
git push origin main
# GitLab CI: install → lint → test → build → deploy to Firebase Hosting
```

The add-in will be live at:

```
https://<FIREBASE_DOMAIN>/index.html
```

---

## Part 3 — Update manifest.xml

Make sure `manifest.xml` contains:

```xml
<WebApplicationInfo>
  <Id>YOUR_CLIENT_ID</Id>
  <Resource>api://<FIREBASE_DOMAIN>/YOUR_CLIENT_ID</Resource>
  <Scopes>
    <Scope>openid</Scope>
    <Scope>profile</Scope>
    <Scope>email</Scope>
  </Scopes>
</WebApplicationInfo>
```

---

## Part 4 — Deploy the Add-in in Outlook (Admin Center)

> This requires admin access to Microsoft 365. The deploying account must have the **Exchange Administrator** or **Global Administrator** role.

### 4.1 Access the Add-ins portal

Go directly to:

```
https://admin.microsoft.com/AdminPortal/Home#/Settings/AddIns
```

> Use this exact URL — do **not** use the "Integrated Apps" portal (`#/Settings/IntegratedApps`), which does not accept XML manifests.

### 4.2 Deploy the add-in

1. Click **+ Deploy Add-In**
2. Click **Next**
3. Select **"I have the manifest file (.xml)"** → click **Choose File** → upload `manifest.xml`
4. Click **Upload**
5. Select who gets the add-in:
   - **Just me** (for testing)
   - **Specific users/groups**
   - **Everyone**
6. In **Deployment method** leave on **Fixed**
7. Click **Deploy**

> The add-in can take **up to 72 hours** to appear in Outlook. To force a refresh, relaunch Outlook or use a private browsing window.

---

## Part 5 — Sideload for Testing (without admin)

If you cannot deploy via the admin center, sideload directly for your own account:

1. Open this URL in your browser (while logged into Outlook Web):
   ```
   https://aka.ms/olksideload
   ```
2. Click **My add-ins** → **Add a custom add-in** → **Add from file**
3. Upload `manifest.xml`
4. Click **Install**

> Note: Sideloading may be blocked by your organization's IT policy.

---

## Troubleshooting

| Symptom                                                      | Cause                                    | Fix                                                                                                                                      |
| ------------------------------------------------------------ | ---------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| `Signing in…` stuck forever                                  | MSAL auth failing — Office.js not loaded | Check that assets load correctly (no 404s in DevTools console)                                                                           |
| `Invalid Reply Address. Must have scheme brk-XXXX://`        | Tenant uses a custom broker scheme       | Add `brk-XXXX://<FIREBASE_DOMAIN>` as SPA redirect URI in Azure (where XXXX is the `brk_client_id` from the failing request in DevTools) |
| `Manifest file validation has failed`                        | Wrong portal used                        | Use `https://admin.microsoft.com/AdminPortal/Home#/Settings/AddIns`, not "Integrated Apps"                                               |
| `Wrong Package. Your package does not match submission type` | Uploaded XML to Teams/M365 Apps portal   | Use the Add-ins portal link above                                                                                                        |
| `This operation was unsuccessful — eligibility requirements` | Admin account lacks required role        | Assign **Exchange Administrator** or **Global Administrator** role                                                                       |
| Add-in not showing after deployment                          | Propagation delay                        | Wait up to 72h, or sideload via `https://aka.ms/olksideload`                                                                             |
| `401` from n8n webhook                                       | JWT token audience mismatch              | Check n8n JWT credential — the PEM key must match the current Microsoft signing key (run `get_pem_from_kid.sh`)                          |
| `Access denied` on the add-in page                           | User not authorized                      | Check Azure → Enterprise applications → assign users                                                                                     |
| `Failed to load resource: 404` on assets                     | Wrong Firebase deploy or `dist` path     | Check `firebase.json` points to `"public": "dist"` and that the build completed successfully                                             |
