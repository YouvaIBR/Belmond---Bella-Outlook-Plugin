# Bella — Belmond Outlook Add-in

Outlook add-in that lets customer service agents generate AI-powered email replies via **Bella**, Belmond's conversational AI (n8n MCP agent).

## How it works

```
Outlook (MSAL NAA)
      │
      │  Bearer JWT (Entra ID signed)
      ▼
n8n Webhook
      │  Verify signature via Microsoft JWKS
      │── invalid ──► 401
      │── valid   ──►
      ▼
MCP Agent (Bella)
      │
      ▼
Response ──► Outlook add-in
```

The plugin authenticates the logged-in Microsoft 365 user silently via MSAL (Nested App Authentication), then sends the JWT directly to the n8n webhook. **No intermediate backend. No credentials in the code.**

---

## Prerequisites

- Node.js 20+
- An **Entra ID App Registration** (see [deployment guide](docs/deployment.md))
- An **n8n** instance with the Bella webhook configured (see [n8n workflow guide](n8n-workflow.md))
- A **public GitHub repository** with GitHub Pages enabled

---

## Local development

### 1. Clone and install

```bash
git clone https://github.com/YouvaIBR/Belmond---Bella-Outlook-Plugin.git
cd Belmond---Bella-Outlook-Plugin
npm install
```

### 2. Configure environment variables

```bash
cp .env.example .env
```

Fill in `.env`:

```env
VITE_CLIENT_ID=<your Entra ID app client ID>
VITE_TENANT_ID=<your Entra ID tenant ID>
VITE_N8N_WEBHOOK_URL=https://your-n8n-instance.com/webhook/bella-outlook
```

### 3. Start the dev server

```bash
npm run dev
```

The add-in is served at `http://localhost:3000`.

> For Outlook to load the add-in, the dev server must be accessible via HTTPS. Use a tunnel such as [ngrok](https://ngrok.com) (`ngrok http 3000`) and update the manifest URLs accordingly.

---

## Available scripts

| Command | Description |
|---|---|
| `npm run dev` | Start Vite dev server (port 3000) |
| `npm run build` | Type-check + production build → `dist/` |
| `npm run test` | Run Vitest unit tests |
| `npm run test:coverage` | Tests with coverage report |
| `npm run lint` | Biome linter |
| `npm run lint:fix` | Auto-fix lint issues |
| `npm run format` | Prettier formatter |

---

## Project structure

```
src/
├── types/index.ts        Types: MailItem, AgentRequest/Response, UIPanel
├── auth/msal.ts          MSAL NAA — acquireTokenSilent → popup fallback
├── api/agent.ts          callAgent() — fetch to n8n with Bearer JWT
├── services/office.ts    Office.js abstractions (getMailItem, insertDraftReply)
├── ui/
│   ├── panels.ts         Show/hide panels, error display, product selector
│   └── draft.ts          HTML sanitizer + DOM rendering
├── main.ts               Office.onReady entry point
└── tests/
    ├── agent.test.ts
    ├── panels.test.ts
    └── draft.test.ts
```

---

## Documentation

| Guide | Description |
|-------|-------------|
| [Deployment Guide](docs/deployment.md) | Complete setup: Azure App Registration, GitHub Pages, admin center deployment, sideloading, and troubleshooting |
| [n8n Workflow Guide](n8n-workflow.md) | n8n webhook configuration, JWT key setup, and testing |

---

## Deployment (GitHub Pages)

Deployment is automated via GitHub Actions on every push to `main`.

### One-time GitHub setup

1. **Settings → Pages** → Source: `GitHub Actions`
2. Make the repository **public**
3. **Settings → Secrets → Actions** → add three secrets:

| Secret | Value |
|---|---|
| `VITE_CLIENT_ID` | Entra ID client ID |
| `VITE_TENANT_ID` | Entra ID tenant ID |
| `VITE_N8N_WEBHOOK_URL` | n8n webhook URL |

### Deploy

```bash
git push origin main
# GitHub Actions: lint → test → build → deploy
```

The add-in will be live at `https://youvaibr.github.io/Belmond---Bella-Outlook-Plugin`.

---

## Installing the add-in in Outlook

### Via Microsoft 365 Admin Center (recommended for production)

Go to `https://admin.microsoft.com/AdminPortal/Home#/Settings/AddIns` and upload `manifest.xml`.

See the full step-by-step in the [Deployment Guide](docs/deployment.md#part-4--deploy-the-add-in-in-outlook-admin-center).

### Via sideload (for testing, no admin required)

Open `https://aka.ms/olksideload` in your browser while logged into Outlook Web, then upload `manifest.xml`.

See [Part 5 of the Deployment Guide](docs/deployment.md#part-5--sideload-for-testing-without-admin).

---

## Known limitations

| Limitation | Detail |
|---|---|
| **Outlook Mobile** | MSAL NAA (`acquireTokenSilent`) fails on mobile — known Microsoft bug. Desktop and OWA work correctly. |
| **New Outlook + corporate proxy / Okta** | Some tenants experience NAA failures. Test on the target environment early. |
| **Conditional Access** | The "approved client app" CA grant is deprecated (retiring March 2026). Tenants using it must migrate to the Application Protection Policy grant. |
| **Key rotation** | Microsoft rotates JWT signing keys every few weeks. Re-run `get_pem_from_kid.sh` and update the n8n credential when this happens. |
