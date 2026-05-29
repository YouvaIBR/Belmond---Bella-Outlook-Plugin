# Deployment Guide — Bella Outlook Add-in

This guide covers everything needed to deploy the Bella Outlook add-in from scratch: Azure configuration, GitHub setup, and installing the add-in in Outlook.

---

## Overview

```
GitHub repo (source)
      │
      │  git push → GitHub Actions
      ▼
GitHub Pages (https://youvaibr.github.io/Belmond---Bella-Outlook-Plugin)
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
   - **Redirect URI**: Single-page application (SPA) → `brk-multihub://youvaibr.github.io`
4. Click **Register**

### 1.2 Note the IDs

On the **Overview** page, copy:
- **Application (client) ID** → used as `VITE_CLIENT_ID`
- **Directory (tenant) ID** → used as `VITE_TENANT_ID`

### 1.3 Add redirect URIs

**Authentication** → existing SPA platform → **Add URI**, add all of these:

| URI | Purpose |
|-----|---------|
| `brk-multihub://youvaibr.github.io` | NAA broker (standard) |
| `brk-<TENANT_BROKER_ID>://youvaibr.github.io` | NAA broker (tenant-specific — see note below) |
| `https://youvaibr.github.io/Belmond---Bella-Outlook-Plugin/index.html` | Outlook Web |

> **Important — Tenant-specific broker ID**: Some Microsoft 365 tenants use a custom broker scheme instead of `brk-multihub`. If you get the error `Invalid Reply Address. Reply Address must have scheme brk-XXXX://`, add that exact scheme with your domain as a redirect URI. The broker ID (`brk_client_id`) is visible in the failing request URL in DevTools.

> **Set `accessTokenAcceptedVersion` to 2**: In Azure → **Manifest (JSON)**, set `"accessTokenAcceptedVersion": 2`. This is required to allow custom Application ID URIs with external domains (e.g. GitHub Pages).

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
   api://youvaibr.github.io/Belmond---Bella-Outlook-Plugin/<YOUR_CLIENT_ID>
   ```
   > This requires `accessTokenAcceptedVersion: 2` (see step 1.3). If your tenant policy blocks custom domains, use `api://<YOUR_CLIENT_ID>` and update the scope in `src/auth/msal.ts` accordingly.

2. **Add a scope**:

| Field | Value |
|-------|-------|
| Scope name | `access_as_user` |
| Who can consent | `Admins and users` |
| Admin consent display name | `Access Bella as user` |
| Admin consent description | `Allows the add-in to access the API as the user` |
| User consent display name | `Access Bella as user` |
| User consent description | `Allows the add-in to access the API as you` |
| State | `Enabled` |

---

## Part 2 — GitHub Repository Setup

### 2.1 Enable GitHub Pages

1. **Settings** → **Pages**
2. **Source** → `GitHub Actions`
3. Make sure the repository is **public** (GitHub Pages is free only on public repos)

### 2.2 Add GitHub Secrets

**Settings** → **Secrets and variables** → **Actions** → **New repository secret**:

| Secret name | Value |
|-------------|-------|
| `VITE_CLIENT_ID` | Azure Application (client) ID |
| `VITE_TENANT_ID` | Azure Directory (tenant) ID |
| `VITE_N8N_WEBHOOK_URL` | n8n webhook URL |

### 2.3 Deploy

Every push to `main` triggers the workflow automatically:

```bash
git push origin main
# GitHub Actions: lint → test → build → deploy to GitHub Pages
```

The add-in will be live at:
```
https://youvaibr.github.io/Belmond---Bella-Outlook-Plugin/index.html
```

---

## Part 3 — Update manifest.xml

Make sure `manifest.xml` contains:

```xml
<WebApplicationInfo>
  <Id>YOUR_CLIENT_ID</Id>
  <Resource>api://youvaibr.github.io/Belmond---Bella-Outlook-Plugin/YOUR_CLIENT_ID</Resource>
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
6. Click **Deploy**

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

| Symptom | Cause | Fix |
|---------|-------|-----|
| `Signing in…` stuck forever | MSAL auth failing — Office.js not loaded | Check that assets load correctly (no 404s in DevTools console) and that `base` path is set in `vite.config.ts` |
| `Invalid Reply Address. Must have scheme brk-XXXX://` | Tenant uses a custom broker scheme | Add `brk-XXXX://youvaibr.github.io` as SPA redirect URI in Azure (where XXXX is the `brk_client_id` from the failing request in DevTools) |
| `Manifest file validation has failed` | Wrong portal used | Use `https://admin.microsoft.com/AdminPortal/Home#/Settings/AddIns`, not "Integrated Apps" |
| `Wrong Package. Your package does not match submission type` | Uploaded XML to Teams/M365 Apps portal | Use the Add-ins portal link above |
| `This operation was unsuccessful — eligibility requirements` | Admin account lacks required role | Assign **Exchange Administrator** or **Global Administrator** role |
| Add-in not showing after deployment | Propagation delay | Wait up to 72h, or sideload via `https://aka.ms/olksideload` |
| `401` from n8n webhook | JWT token audience mismatch | Check n8n JWT credential — the PEM key must match the current Microsoft signing key (run `get_pem_from_kid.sh`) |
| `Access denied` on the add-in page | User not authorized | Check Azure → Enterprise applications → assign users |
| `Failed to load resource: 404` on assets | Wrong Vite base path | Ensure `vite.config.ts` sets `base: "/Belmond---Bella-Outlook-Plugin/"` for production |
