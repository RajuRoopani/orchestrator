# 🧠 Claude Orchestrator

A local AI orchestration dashboard built on the **Anthropic Claude API**. It lets you describe high-level tasks and automatically breaks them into a hierarchical execution plan, runs sub-agents in parallel, and surfaces real-time results — all through a clean dark-glass web UI.

It also ships with a first-class **Microsoft ICM (Incident Management) Dashboard** and a **DRI Investigation** workflow for on-call engineers.

---

## Screenshots

### Plan & Chat
![Plan view](docs/screenshots/chat.png)

### ICM Dashboard
![ICM Dashboard](docs/screenshots/icm.png)

### DRI Investigation
![DRI Investigation](docs/screenshots/dri.png)

---

## Features

| Feature | Description |
|---------|-------------|
| 🧠 **AI Orchestration** | Describe a task → Claude generates a hierarchical plan → sub-agents execute each step |
| ⚡ **Real-time Dashboard** | Live WebSocket updates as agents work through plan steps |
| 🔥 **ICM Dashboard** | Browse active Microsoft ICM incidents with Fluent Glass UI — severity badges, stats bar, team filters |
| 🚨 **DRI Investigation** | One-command `/DRI [incident]` workflow — 7-step automated triage using Kusto, ADO, Geneva log links, and related ICMs |
| 📊 **Hourly Activity Feed** | Auto-generated summary of your last hour of coding activity (claude-mem + browser context), refreshable on demand |
| 📋 **History** | Persistent session history stored locally at `~/.orchestrator/history/` |
| 🎨 **Fluent Glass UI** | Deep blue/purple glassmorphism theme — gradient severity badges, glowing stat dots, smooth transitions |

---

## Tech Stack

- **Backend:** Node.js + Express + TypeScript
- **AI:** Claude CLI (`claude -p`) — no API key required, uses your Claude Code session
- **Real-time:** WebSockets (`ws`)
- **Frontend:** Vanilla HTML/CSS/JS (zero framework, zero build step)
- **ICM API:** `https://prod.microsofticm.com/api2/incidentapi` (token auto-loaded from ambient-mcp)
- **Kusto:** `icmcluster.kusto.windows.net / IcmDataWarehouse` via REST API (ambient-mcp token)
- **ADO MCP:** `@azure-devops/mcp` with `az cli` auth (`rroopani@microsoft.com`)

---

## Prerequisites

- Node.js 18+
- Claude Code CLI (`claude`) authenticated and in your `PATH`
- (Optional) `az cli` logged into the MSFT corp tenant for ADO MCP tools:
  ```bash
  az login --tenant 72f988bf-86f1-41af-91ab-2d7cd011db47 --allow-no-subscriptions
  ```
- (Optional) Microsoft ICM bearer token — auto-loaded from `ambient-mcp` if available

---

## Setup

### 1. Clone

```bash
git clone https://github.com/RajuRoopani/orchestrator.git
cd orchestrator
```

### 2. Install dependencies

```bash
npm install
```

### 3. Configure environment

```bash
cp .env.example .env
```

Edit `.env` — no API key needed, the server uses the `claude` CLI:

```env
PORT=3333          # optional, default 3333
```

### 4. Build & start

```bash
npm run dev        # build TypeScript + start server
```

Or separately:

```bash
npm run build      # tsc → dist/
npm start          # node dist/index.js
```

Open **http://localhost:3333** in your browser.

---

## Usage

### AI Plan & Execute

1. Type a task description in the chat panel, e.g.:
   > *"Create a REST API in Python with FastAPI that manages a todo list with SQLite"*
2. Click **Generate Plan** — Claude breaks it into a hierarchical execution plan.
3. Watch the **Dashboard** tab as sub-agents execute each step in real time.

### DRI Investigation

Type in the chat:
```
/DRI Bot fails to respond in tenant xyz-corp, getting 403 errors since 14:00 UTC
```

This kicks off a 7-step automated DRI investigation:

| Step | What it does |
|------|-------------|
| 1 | **Load ICM** — queries `IcmDataWarehouse` via Kusto REST API |
| 2 | **Extract identifiers** — TenantId, BotId, ThreadId, CorrelationId |
| 3 | **Kusto triage** — tenant anomalies, ACL issues, hot shard analysis |
| 4 | **ADO work items** — related bugs, hotfixes, TSGs in O365Exchange |
| 5 | **Geneva log URLs** — pre-built deep-links for LogMessage / IncomingRequest / OutgoingRequest |
| 6 | **Related ICMs & docs** — similar incidents + learn.microsoft.com resources |
| 7 | **Compile report** — structured JSON + HTML report at `/api/dri/:icmId/report` |

> **ADO auth:** Requires `az cli` logged into the MSFT corp tenant (see Prerequisites). Run `az login --tenant 72f988bf-86f1-41af-91ab-2d7cd011db47` once.

### ICM Dashboard

The token is **auto-loaded** from `ambient-mcp` if you have the browser context MCP running. If not:

1. Go to the **ICM Dashboard** tab.
2. Click **⚙ Token** in the top-right.
3. Paste your `Authorization` header from `portal.microsofticm.com`
   (DevTools → Network → any API request → copy the `Authorization` value).
4. Select your team and click **Save & Load**.

The dashboard shows:
- Severity 1/2/3/Active counts in the stats bar
- Per-incident severity badges, status pills, flags (Outage / CRI)
- Direct **🚨 Investigate** button to launch a DRI workflow for any incident

> The ambient-mcp token refreshes automatically every 10 minutes in the background.

### Hourly Activity Feed

The **Dashboard** tab includes an auto-generated summary of your last hour of activity — what you built, browsed, and focused on — powered by `claude-mem` and `ambient-mcp`.

- Generates automatically 15s after server start, then every hour
- Click **↻ Generate Now** to refresh on demand
- Shows a loading state and surfaces errors if generation fails

---

## Project Structure

```
orchestrator/
├── src/
│   ├── index.ts            # Express server, REST API, ICM proxy, WebSocket setup
│   ├── orchestrator.ts     # AI plan generation using Claude
│   ├── executor.ts         # Plan execution engine (runs sub-agents)
│   ├── dri-orchestrator.ts # DRI investigation workflow
│   ├── ws-manager.ts       # WebSocket connection manager
│   └── models.ts           # Shared TypeScript types
├── public/
│   ├── index.html          # Single-page app shell
│   ├── app.js              # Frontend logic (tabs, ICM, DRI, WebSocket client)
│   └── styles.css          # Fluent Glass dark theme
├── commands/
│   └── DRI.md              # DRI slash command definition
├── docs/
│   └── screenshots/        # UI screenshots
├── .env.example
├── package.json
└── tsconfig.json
```

---

## API Reference

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/api/plan` | Generate a Claude execution plan |
| `POST` | `/api/execute` | Execute a plan |
| `POST` | `/api/dri` | Start a DRI investigation (auto-executes) |
| `GET`  | `/api/dri/:icmId/report` | Render full HTML DRI report |
| `POST` | `/api/icm/token` | Set ICM bearer token + team ID |
| `GET`  | `/api/icm/active` | Fetch active ICM incidents (cached 5 min) |
| `POST` | `/api/icm/refresh` | Bust the ICM cache |
| `GET`  | `/api/ambient/icm-token` | Auto-load ICM token from ambient-mcp |
| `GET`  | `/api/adx/icm/:icmId` | Query IcmDataWarehouse via Kusto REST |
| `GET`  | `/api/activity-summary` | Get latest + history of activity summaries |
| `POST` | `/api/activity-summary/generate` | Trigger on-demand activity summary |
| `GET`  | `/api/history` | List past DRI sessions |
| `WS`   | `ws://localhost:3333` | Real-time execution + activity updates |

---

## Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `PORT` | ❌ | `3333` | HTTP server port |

> No API key needed — all LLM calls use the `claude` CLI which authenticates via your Claude Code session.

---

## Development

```bash
npm run build       # compile TypeScript
npm run type-check  # type-check without emitting
npm run dev         # build + run (one command)
```

To watch for changes during development, use `tsc --watch` in one terminal and `node dist/index.js` in another, or add `nodemon` to the dev dependencies.

---

## License

MIT
