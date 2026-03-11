import 'dotenv/config';
import express from 'express';
import { createServer } from 'http';
import { WebSocketServer } from 'ws';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import { WSManager } from './ws-manager';
import { Orchestrator } from './orchestrator';
import { Executor } from './executor';
import { createDriPlan } from './dri-orchestrator';
import { Plan } from './models';

const PORT = parseInt(process.env.PORT ?? '3333', 10);
const HISTORY_DIR = path.join(os.homedir(), '.orchestrator', 'history');
const CHAT_FILE   = path.join(os.homedir(), '.orchestrator', 'chat.json');
const SUMMARY_FILE = path.join(os.homedir(), '.orchestrator', 'activity-summaries.json');
fs.mkdirSync(HISTORY_DIR, { recursive: true });

/* ─── ICM cache ──────────────────────────────────────────────────────────────── */
interface IcmEntry { IncidentId: number; Title: string; Severity: number; Status: string; OwningTeamName: string; OwningTenantName: string; ContactAlias: string; CreateDate: string; MitigateDate: string | null; ImpactStartTime: string | null; IsCustomerImpacting: boolean; IsOutage: boolean; HitCount: number; OwningTeamId: number; Keywords: string; }
let icmCache: { data: IcmEntry[]; fetchedAt: number } | null = null;
const ICM_CACHE_TTL = 5 * 60 * 1000; // 5 minutes
let icmTeamId: number | null = null;

const ICM_API_BASE = 'https://prod.microsofticm.com/api2/incidentapi';
let icmBearerToken: string | null = null;

function aliasFromJwt(token: string): string | null {
  try {
    const payload = JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString('utf-8')) as Record<string, string>;
    const upn = payload['upn'] || payload['unique_name'] || payload['preferred_username'] || payload['email'] || '';
    return upn.split('@')[0] || null;
  } catch { return null; }
}

async function fetchActiveIcms(): Promise<IcmEntry[]> {
  if (icmCache && Date.now() - icmCache.fetchedAt < ICM_CACHE_TTL) return icmCache.data;
  if (!icmBearerToken) throw new Error('NO_TOKEN');

  const alias = aliasFromJwt(icmBearerToken);
  console.log(`[ICM] alias from token: ${alias}, teamId: ${icmTeamId}`);

  const https = require('https') as typeof import('https');
  const filterParts = ['ParentId eq null', "State ne 'RESOLVED'", "State ne 'CLOSED'"];
  if (icmTeamId) filterParts.unshift(`OwningTeamId eq ${icmTeamId}`);
  const params = new URLSearchParams({
    '$filter': filterParts.join(' and '),
    '$orderby': 'Severity asc,CreatedDate desc',
    '$top': '100',
    '$select': 'Id,Severity,State,Title,CreatedDate,OwningTenantName,OwningTeamName,ContactAlias,NotificationStatus,HitCount,OwningServiceId,OwningTeamId,ParentId,IsCustomerImpacting,IsNoise,IsOutage,MitigateData,ImpactStartTime',
    '$expand': 'CustomFields,AlertSource',
  });
  const url = `${ICM_API_BASE}/incidents?${params}`;

  return new Promise((resolve, reject) => {
    const req = https.request(url, {
      method: 'GET',
      headers: {
        'Authorization': icmBearerToken!.startsWith('Bearer ') ? icmBearerToken! : `Bearer ${icmBearerToken}`,
        'Accept': 'application/json',
      },
    }, (res) => {
      const chunks: Buffer[] = [];
      res.on('data', (c: Buffer) => chunks.push(c));
      res.on('end', () => {
        const body = Buffer.concat(chunks).toString('utf-8');
        console.log(`[ICM] URL: ${url}`);
        console.log(`[ICM] HTTP ${res.statusCode}, body snippet: ${body.slice(0, 300)}`);
        if (res.statusCode === 401) return reject(new Error('TOKEN_EXPIRED'));
        if (res.statusCode === 403) return reject(new Error('TOKEN_FORBIDDEN'));
        if ((res.statusCode ?? 0) >= 400) return reject(new Error(`ICM HTTP ${res.statusCode}: ${body.slice(0, 200)}`));
        try {
          const json = JSON.parse(body) as { value: IcmEntry[] };
          const rows = (json.value ?? []).map((item: unknown) => { const r = item as Record<string, unknown>;
            const state = String(r['State'] ?? '');
            const normalizedStatus = state === 'ACTIVE' ? 'Active' : state === 'MITIGATED' ? 'Mitigated' : state === 'CLOSED' ? 'Closed' : state;
            const mitigateData = r['MitigateData'] as Record<string, unknown> | null;
            return ({
              IncidentId:         r['Id'] as number,
              Title:              r['Title'] as string,
              Severity:           r['Severity'] as number,
              Status:             normalizedStatus,
              OwningTeamName:     (r['OwningTeamName'] ?? '') as string,
              OwningTenantName:   (r['OwningTenantName'] ?? '') as string,
              ContactAlias:       (r['ContactAlias'] ?? '') as string,
              OwningTeamId:       (r['OwningTeamId'] ?? 0) as number,
              CreateDate:         (r['CreatedDate'] ?? '') as string,
              MitigateDate:       (mitigateData?.['Date'] ?? null) as string | null,
              ImpactStartTime:    (r['ImpactStartTime'] ?? null) as string | null,
              IsCustomerImpacting:(r['IsCustomerImpacting'] ?? false) as boolean,
              IsOutage:           (r['IsOutage'] ?? false) as boolean,
              HitCount:           (r['HitCount'] ?? 0) as number,
              Keywords:           '',
            }); });
          console.log(`[ICM] parsed ${rows.length} rows`);
          icmCache = { data: rows, fetchedAt: Date.now() };
          resolve(rows);
        } catch (e) {
          reject(new Error('Parse error: ' + (e as Error).message));
        }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

/* ─── Ambient-MCP token reader ───────────────────────────────────────────────── */
const AMBIENT_STORE = path.join(os.homedir(), '.claude', 'browser-context-mcp', 'store.json');

const ICM_DOMAINS = [
  'prod.microsofticm.com',
  'portal.microsofticm.com',
  'upsapi.prod.microsofticm.com',
  'oncallapi.prod.microsofticm.com',
];

interface AmbientToken {
  token: string;
  domain: string;
  status: string;
  capturedAt: string;
}

function readAmbientIcmToken(): AmbientToken | null {
  try {
    if (!fs.existsSync(AMBIENT_STORE)) return null;
    const store = JSON.parse(fs.readFileSync(AMBIENT_STORE, 'utf-8')) as {
      tokens?: Record<string, { headers?: Record<string, string>; status?: string; capturedAt?: string }>;
    };
    const tokens = store.tokens ?? {};
    for (const domain of ICM_DOMAINS) {
      const entry = tokens[domain];
      if (!entry) continue;
      const auth = entry.headers?.['Authorization'] ?? entry.headers?.['authorization'] ?? '';
      if (!auth) continue;
      const raw = auth.startsWith('Bearer ') ? auth.slice(7) : auth;
      if (!raw) continue;
      return { token: raw, domain, status: entry.status ?? 'unknown', capturedAt: entry.capturedAt ?? '' };
    }
    return null;
  } catch { return null; }
}

/* ─── Ambient-MCP Kusto token reader ────────────────────────────────────────── */
const KUSTO_DOMAIN = 'icmcluster.kusto.windows.net';

function readAmbientKustoToken(): AmbientToken | null {
  try {
    if (!fs.existsSync(AMBIENT_STORE)) return null;
    const store = JSON.parse(fs.readFileSync(AMBIENT_STORE, 'utf-8')) as {
      tokens?: Record<string, { headers?: Record<string, string>; status?: string; capturedAt?: string }>;
    };
    const entry = (store.tokens ?? {})[KUSTO_DOMAIN];
    if (!entry) return null;
    const auth = entry.headers?.['Authorization'] ?? entry.headers?.['authorization'] ?? '';
    if (!auth) return null;
    const raw = auth.startsWith('Bearer ') ? auth.slice(7) : auth;
    if (!raw) return null;
    return { token: raw, domain: KUSTO_DOMAIN, status: entry.status ?? 'unknown', capturedAt: entry.capturedAt ?? '' };
  } catch { return null; }
}

type KustoV2Frame =
  | { FrameType: 'DataTable'; TableName: string; Columns: { ColumnName: string }[]; Rows: unknown[][] }
  | { FrameType: string };

/** Execute a Kusto query against IcmDataWarehouse and return rows as objects */
async function runKustoQuery(kql: string): Promise<Record<string, unknown>[]> {
  const raw = readAmbientKustoToken();
  if (!raw) throw new Error('No Kusto token found in ambient-mcp store for icmcluster.kusto.windows.net');

  // Read extra headers the browser used (x-ms-user-id, x-ms-app) for auth context
  const storeEntry = (() => {
    try {
      const store = JSON.parse(fs.readFileSync(AMBIENT_STORE, 'utf-8')) as {
        tokens?: Record<string, { headers?: Record<string, string> }>;
      };
      return store.tokens?.[KUSTO_DOMAIN]?.headers ?? {};
    } catch { return {} as Record<string, string>; }
  })();

  const headers: Record<string, string> = {
    'Authorization': `Bearer ${raw.token}`,
    'Content-Type': 'application/json',
    'Accept': 'application/json',
  };
  if (storeEntry['x-ms-user-id']) headers['x-ms-user-id'] = storeEntry['x-ms-user-id'];
  if (storeEntry['x-ms-app']) headers['x-ms-app'] = storeEntry['x-ms-app'];

  const resp = await fetch('https://icmcluster.kusto.windows.net/v2/rest/query', {
    method: 'POST',
    headers,
    body: JSON.stringify({ db: 'IcmDataWarehouse', csl: kql }),
  });

  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`Kusto query failed (${resp.status}): ${body.slice(0, 300)}`);
  }

  // v2 REST API returns a list of frames; the data is in the PrimaryResult frame
  const frames = await resp.json() as KustoV2Frame[];
  const dataFrame = frames.find(
    (f): f is KustoV2Frame & { FrameType: 'DataTable'; TableName: string; Columns: { ColumnName: string }[]; Rows: unknown[][] } =>
      f.FrameType === 'DataTable' && (f as { TableName?: string }).TableName === 'PrimaryResult'
  ) ?? frames.find(
    (f): f is KustoV2Frame & { FrameType: 'DataTable'; Columns: { ColumnName: string }[]; Rows: unknown[][] } =>
      f.FrameType === 'DataTable' && Array.isArray((f as { Rows?: unknown }).Rows)
  );

  if (!dataFrame || !('Columns' in dataFrame)) return [];
  const cols = dataFrame.Columns.map(c => c.ColumnName);
  return dataFrame.Rows.map(row => Object.fromEntries(cols.map((c, i) => [c, row[i]])));
}

/* ─── Chat persistence ───────────────────────────────────────────────────────── */
function loadChat(): { role: string; text: string }[] {
  try { if (fs.existsSync(CHAT_FILE)) return JSON.parse(fs.readFileSync(CHAT_FILE, 'utf-8')); }
  catch {}
  return [];
}
function saveChat(messages: { role: string; text: string }[]) {
  fs.writeFileSync(CHAT_FILE, JSON.stringify(messages.slice(-200), null, 2));
}

/* ─── Activity Summary ───────────────────────────────────────────────────────── */
interface ActivitySummary {
  generatedAt: string;
  periodLabel: string;
  headline: string;
  sections: { title: string; icon: string; items: string[] }[];
  insights: string[];
  browsers: { title: string; url: string; time: string }[];
}
let latestSummary: ActivitySummary | null = null;

function queryClaudeMem(sql: string): string {
  const dbPath = path.join(os.homedir(), '.claude-mem', 'claude-mem.db');
  if (!fs.existsSync(dbPath)) return '';
  try {
    const { execSync } = require('child_process') as typeof import('child_process');
    return execSync(`sqlite3 "${dbPath}" "${sql.replace(/"/g, '\\"')}"`, { encoding: 'utf-8', timeout: 5000 }) as string;
  } catch { return ''; }
}

function getRecentObservations(hoursBack = 1): string {
  const since = Date.now() - hoursBack * 3600000;
  return queryClaudeMem(
    `SELECT project, type, title, substr(subtitle,1,120), created_at FROM observations WHERE created_at_epoch > ${since} ORDER BY created_at_epoch DESC LIMIT 30`
  ) || 'No recent observations.';
}

function getRecentSessions(hoursBack = 1): string {
  const since = Date.now() - hoursBack * 3600000;
  return queryClaudeMem(
    `SELECT project, request, completed, created_at FROM session_summaries WHERE created_at_epoch > ${since} ORDER BY created_at_epoch DESC LIMIT 10`
  ) || '';
}

function getBrowserContext(): { summary: string; visits: { title: string; url: string; time: string }[] } {
  const storePath = path.join(os.homedir(), '.claude', 'browser-context-mcp', 'store.json');
  try {
    if (!fs.existsSync(storePath)) return { summary: 'No browser context.', visits: [] };
    const store = JSON.parse(fs.readFileSync(storePath, 'utf-8')) as {
      visits?: { title?: string; url?: string; visitedAt?: string }[];
    };
    const visits = (store.visits ?? []).slice(0, 20).map((v) => ({
      title: v.title ?? '(no title)',
      url: v.url ?? '',
      time: v.visitedAt ? new Date(v.visitedAt).toLocaleTimeString() : '',
    }));
    const lines = visits.map((v) => `${v.time} — ${v.title} (${v.url})`).join('\n');
    return { summary: lines || 'No visits recorded.', visits: visits.slice(0, 5) };
  } catch {
    return { summary: 'Could not read browser context.', visits: [] };
  }
}

async function generateActivitySummary(): Promise<void> {
  console.log('[Activity] Generating hourly summary…');
  try {
    const { spawn } = require('child_process') as typeof import('child_process');

    const observations = getRecentObservations(1);
    const sessions     = getRecentSessions(1);
    const browser      = getBrowserContext();
    const now          = new Date();
    const h = now.getHours();
    const periodLabel  = `${h}:00 – ${h}:59`;

    const prompt = `You are summarising a developer's last hour of activity for their personal dashboard.

<observations>${observations}</observations>
<sessions>${sessions || 'None'}</sessions>
<browser_visits>${browser.summary}</browser_visits>

Return ONLY valid JSON (no markdown fences) with this exact shape:
{
  "headline": "One punchy sentence under 80 chars summarising the hour",
  "sections": [
    { "title": "Built / Changed",      "icon": "🛠",  "items": ["…"] },
    { "title": "Browsed & Researched", "icon": "🌐",  "items": ["…"] },
    { "title": "Focus Patterns",       "icon": "🎯",  "items": ["…"] }
  ],
  "insights": ["Actionable insight 1", "Insight 2", "Insight 3"]
}
Rules: ≤4 items/section, ≤3 insights, ≤80 chars/item, mention specific file/project names.`;

    const text = await new Promise<string>((resolve, reject) => {
      const env = { ...process.env };
      delete env['CLAUDECODE'];
      delete env['CLAUDE_CODE_ENTRYPOINT'];
      delete env['ANTHROPIC_API_KEY'];

      const proc = spawn('claude', [
        '-p', prompt,
        '--model', 'claude-haiku-4-5-20251001',
        '--output-format', 'stream-json',
        '--verbose',
      ], { env, stdio: ['ignore', 'pipe', 'pipe'] });

      let fullText = '';
      let buf = '';
      let errOut = '';
      proc.stdout.on('data', (d: Buffer) => {
        buf += d.toString();
        const lines = buf.split('\n');
        buf = lines.pop() ?? '';
        for (const line of lines) {
          if (!line.trim()) continue;
          try {
            const ev = JSON.parse(line) as { type: string; message?: { content?: { type: string; text?: string }[] } };
            if (ev.type === 'assistant' && Array.isArray(ev.message?.content)) {
              fullText += ev.message!.content!.filter((c) => c.type === 'text').map((c) => c.text ?? '').join('');
            }
          } catch { /* ignore */ }
        }
      });
      proc.stderr.on('data', (d: Buffer) => { errOut += d.toString(); });
      const timer = setTimeout(() => { proc.kill(); reject(new Error('Activity generation timed out')); }, 90000);
      proc.on('close', (code: number) => {
        clearTimeout(timer);
        if (code !== 0) reject(new Error(`claude exited ${code}: ${errOut.slice(0, 200)}`));
        else resolve(fullText.trim().replace(/^```json\s*/i, '').replace(/```\s*$/, ''));
      });
    });

    const parsed = JSON.parse(text) as Omit<ActivitySummary, 'generatedAt' | 'periodLabel' | 'browsers'>;

    const summary: ActivitySummary = {
      generatedAt: now.toISOString(),
      periodLabel,
      headline:  parsed.headline,
      sections:  parsed.sections,
      insights:  parsed.insights,
      browsers:  browser.visits,
    };

    latestSummary = summary;

    // Persist last 24 summaries
    try {
      const existing: ActivitySummary[] = fs.existsSync(SUMMARY_FILE)
        ? JSON.parse(fs.readFileSync(SUMMARY_FILE, 'utf-8')) : [];
      fs.writeFileSync(SUMMARY_FILE, JSON.stringify([summary, ...existing].slice(0, 24), null, 2));
    } catch {}

    wsManager.broadcast({ type: 'activity_summary', summary });
    console.log('[Activity] Summary broadcast — headline:', summary.headline);
  } catch (err) {
    const msg = (err as Error).message;
    console.error('[Activity] Generation failed:', msg);
    wsManager.broadcast({ type: 'activity_error', message: msg } as never);
  }
}

function saveHistory(entry: Record<string, unknown>) {
  const file = path.join(HISTORY_DIR, `${entry.planId ?? Date.now()}.json`);
  fs.writeFileSync(file, JSON.stringify(entry, null, 2));
}

function loadHistory(): Record<string, unknown>[] {
  return fs.readdirSync(HISTORY_DIR)
    .filter((f) => f.endsWith('.json'))
    .map((f) => {
      try { return JSON.parse(fs.readFileSync(path.join(HISTORY_DIR, f), 'utf-8')); }
      catch { return null; }
    })
    .filter(Boolean)
    .sort((a: any, b: any) => (b.startedAt ?? 0) - (a.startedAt ?? 0));
}

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'public')));

const httpServer = createServer(app);
const wss = new WebSocketServer({ server: httpServer });

const wsManager = new WSManager(wss);
const orchestrator = new Orchestrator(wsManager);
const executor = new Executor(wsManager);

// Multi-plan state
const plans = new Map<string, Plan>();
const executingPlans = new Set<string>();
let latestPlanId: string | null = null;

// ─── Routes ───────────────────────────────────────────────────────────────────

app.post('/api/plan', async (req, res) => {
  const { message } = req.body as { message?: string };
  if (!message?.trim()) {
    res.status(400).json({ error: 'message is required' });
    return;
  }

  try {
    const plan = await orchestrator.generatePlan(message.trim());
    plans.set(plan.id, plan);
    latestPlanId = plan.id;
    wsManager.broadcast({ type: 'plan_ready', plan });
    res.json({ success: true, planId: plan.id, taskCount: plan.tasks.length });
  } catch (err) {
    const msg = (err as Error).message;
    console.error('[API] Plan generation failed:', msg);
    wsManager.broadcast({ type: 'error', message: `Plan generation failed: ${msg}` });
    res.status(500).json({ error: msg });
  }
});

app.post('/api/execute', async (req, res) => {
  const { planId } = req.body as { planId?: string };
  const id = planId ?? latestPlanId;

  if (!id || !plans.has(id)) {
    res.status(400).json({ error: 'Plan not found. Generate a plan first.' });
    return;
  }
  if (executingPlans.has(id)) {
    res.status(409).json({ error: 'This plan is already executing.' });
    return;
  }

  const plan = plans.get(id)!;
  executingPlans.add(id);
  res.json({ success: true, message: 'Execution started', planId: id });

  executor
    .execute(plan)
    .catch((err) => {
      console.error('[API] Execution error:', err);
      wsManager.broadcast({ type: 'error', message: `Execution error: ${(err as Error).message}` });
    })
    .finally(() => {
      executingPlans.delete(id);
    });
});

app.post('/api/dri', async (req, res) => {
  const { message } = req.body as { message?: string };
  if (!message?.trim()) {
    res.status(400).json({ error: 'message is required' });
    return;
  }

  const plan = createDriPlan(message.trim(), PORT);
  plans.set(plan.id, plan);
  latestPlanId = plan.id;

  // Broadcast plan immediately so the UI can render the DRI tab
  wsManager.broadcast({ type: 'plan_ready', plan });
  res.json({ success: true, planId: plan.id, taskCount: plan.tasks.length });

  // Auto-execute DRI plans — no user approval step needed
  if (executingPlans.has(plan.id)) return;
  executingPlans.add(plan.id);
  const driStartedAt = Date.now();
  executor
    .execute(plan)
    .then((stats) => {
      saveHistory({
        planId: plan.id,
        icmId: plan.icmId ?? 'UNKNOWN',
        title: plan.title,
        description: plan.description,
        type: 'dri',
        startedAt: driStartedAt,
        completedAt: Date.now(),
        durationMs: stats.durationMs,
        totalTasks: stats.totalTasks,
        completedTasks: stats.completedTasks,
        failedTasks: stats.failedTasks,
        reportPath: `/tmp/icm_${plan.icmId}_report.html`,
      });
    })
    .catch((err) => {
      console.error('[API/DRI] Execution error:', err);
      wsManager.broadcast({ type: 'error', message: `DRI execution error: ${(err as Error).message}` });
    })
    .finally(() => {
      executingPlans.delete(plan.id);
    });
});

app.post('/api/icm/token', (req, res) => {
  const { token, teamId } = req.body as { token?: string; teamId?: number | string };
  if (!token?.trim()) {
    icmBearerToken = null;
    icmCache = null;
    icmTeamId = null;
    res.json({ success: true, cleared: true });
    return;
  }
  icmBearerToken = token.trim().replace(/^Bearer\s+/i, '');
  icmTeamId = teamId ? Number(teamId) : null;
  icmCache = null;
  const alias = aliasFromJwt(icmBearerToken);
  res.json({ success: true, alias, teamId: icmTeamId });
});

app.get('/api/icm/active', async (_req, res) => {
  try {
    const data = await fetchActiveIcms();
    res.json({ success: true, data, fetchedAt: icmCache?.fetchedAt });
  } catch (err) {
    const msg = (err as Error).message;
    console.error('[ICM] fetch failed:', msg);
    const status = msg === 'NO_TOKEN' ? 401 : msg === 'TOKEN_EXPIRED' ? 401 : 500;
    res.status(status).json({ success: false, error: msg });
  }
});

app.post('/api/icm/refresh', (_req, res) => {
  icmCache = null;
  res.json({ success: true });
});

app.get('/api/history', (_req, res) => {
  res.json(loadHistory());
});

app.delete('/api/history/:planId', (req, res) => {
  const file = path.join(HISTORY_DIR, `${req.params.planId}.json`);
  if (fs.existsSync(file)) { fs.unlinkSync(file); res.json({ success: true }); }
  else res.status(404).json({ error: 'Not found' });
});

app.get('/api/dri/:icmId/report', (req, res) => {
  const { icmId } = req.params;

  // Try new JSON-based report first
  const wsBase = path.join(os.homedir(), 'Work6', 'experiments', 'orchestrator-workspaces');
  const jsonReports = fs.existsSync(wsBase)
    ? fs.readdirSync(wsBase)
        .filter((d) => d.startsWith('dri-'))
        .map((d) => path.join(wsBase, d, 'report.json'))
        .filter((p) => fs.existsSync(p))
        .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs)
    : [];

  const matchingJson = jsonReports.find((p) => {
    try { return JSON.parse(fs.readFileSync(p, 'utf-8')).icmId === icmId; } catch { return false; }
  });

  if (matchingJson) {
    try {
      const d = JSON.parse(fs.readFileSync(matchingJson, 'utf-8')) as Record<string, unknown>;
      const sev = Number(d.severity) || 0;
      const sevColor = ['','#c50f1f','#da3b01','#c19c00','#0078d4'][sev] ?? '#0078d4';
      const ids = (d.identifiers as Record<string,string>) ?? {};
      const identifierRows = Object.entries(ids).filter(([,v]) => v)
        .map(([k,v]) => `<tr><td class="id-label">${k}</td><td><div class="id-value-wrap"><span class="id-value">${v}</span><button class="copy-btn" onclick="copyId(this,'${v.replace(/'/g,"\\'")}')">Copy</button></div></td></tr>`).join('');
      const nextSteps = ((d.nextSteps as {title:string;detail:string;priority:string}[]) ?? [])
        .map((s,i) => `<li class="action-item"><div class="action-num pri-${s.priority}">${i+1}</div><div><div class="action-title">${s.title}</div><div class="action-body">${s.detail}</div></div></li>`).join('');
      const evidence = ((d.evidenceChecklist as {label:string;status:string;detail:string}[]) ?? [])
        .map((e) => `<li class="evidence-item"><div class="evidence-icon ${e.status}">${e.status==='pass'?'✓':e.status==='fail'?'✗':e.status==='warn'?'!':'?'}</div><div><div class="evidence-label">${e.label}</div><div class="evidence-detail">${e.detail}</div></div></li>`).join('');
      const gUrls = (d.genevaUrls as Record<string,string>) ?? {};
      const genevaHtml = Object.entries(gUrls).filter(([,v])=>v)
        .map(([k,v]) => `<a class="geneva-link ${k.toLowerCase().includes('log')?'log':k.toLowerCase().includes('in')?'incoming':'outgoing'}" href="${v}" target="_blank">📊 ${k}</a>`).join('');
      const discussions = ((d.discussions as {time:string;author:string;text:string}[]) ?? []).slice(0,20)
        .map((e) => `<div class="timeline-item"><div class="timeline-dot"></div><div class="timeline-time">${e.time}</div><div class="timeline-author">${e.author}</div><div class="timeline-content">${e.text}</div></div>`).join('');
      const relatedRows = ((d.relatedIcms as {id:string;title:string;severity:number;status:string;createDate:string}[]) ?? [])
        .map((r) => `<tr><td><a class="icm-link" href="https://portal.microsofticm.com/imp/v5/incidents/details/${r.id}/summary" target="_blank">${r.id}</a></td><td>${r.title}</td><td>Sev ${r.severity}</td><td>${r.status}</td><td>${r.createDate}</td></tr>`).join('');
      const docs = ((d.docs as {title:string;url:string;description:string}[]) ?? [])
        .map((doc) => `<li class="doc-item"><div><a class="doc-link" href="${doc.url}" target="_blank">${doc.title}</a><div class="doc-desc">${doc.description}</div></div></li>`).join('');

      const html = `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>ICM ${icmId} — DRI Report</title>
<style>:root{--ms-blue:#0078d4;--danger:#c50f1f;--warning:#c19c00;--success:#107c10;--surface:#fff;--surface3:#f0f2f5;--border:#e1e4e8;--text-primary:#1b1b1b;--text-secondary:#616161;--text-muted:#8a8a8a;--radius:8px;--radius-lg:12px;--shadow:0 2px 8px rgba(0,0,0,.08)}
*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--surface3);color:var(--text-primary);line-height:1.5;font-size:14px}
.header{background:#1b1b1b;color:#fff;padding:0 32px;height:56px;display:flex;align-items:center;gap:16px;position:sticky;top:0;z-index:100}
.header-icm{font-size:14px;font-weight:700;color:#60cdff}.hs{flex:1}.header-meta{font-size:12px;color:rgba(255,255,255,.5)}
.print-btn{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:#fff;padding:5px 14px;border-radius:4px;font-size:12px;cursor:pointer}
.page{max-width:1280px;margin:0 auto;padding:24px 24px 48px}
.hero{background:var(--surface);border-radius:var(--radius-lg);box-shadow:var(--shadow);padding:24px 28px;margin-bottom:20px;border-top:4px solid ${sevColor}}
.hero-title{font-size:20px;font-weight:700;line-height:1.3;margin-bottom:8px}
.badge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:100px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-right:6px}
.sev1,.sev2{color:#c50f1f;background:#fde7e9}.sev3{color:#c19c00;background:#fff8e1}.sev4{color:#0078d4;background:#e8f4fd}
.hero-meta-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:16px;margin-top:16px;padding-top:16px;border-top:1px solid var(--border)}
.meta-item label{display:block;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--text-muted);margin-bottom:3px}.meta-item span{font-size:13px;font-weight:500}
.rca-box{background:linear-gradient(135deg,#fff8e1,#fde7e9);border:1px solid #f0c060;border-left:4px solid var(--warning);border-radius:var(--radius);padding:16px 18px;margin-bottom:20px}
.rca-box h3{font-size:13px;font-weight:700;color:var(--warning);margin-bottom:8px}.rca-box p{font-size:13px;line-height:1.6}
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}@media(max-width:900px){.grid-2{grid-template-columns:1fr}}
.card{background:var(--surface);border-radius:var(--radius-lg);box-shadow:var(--shadow);overflow:hidden}
.card-header{padding:14px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}.card-header h2{font-size:13px;font-weight:700;flex:1}
.card-header-icon{width:28px;height:28px;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:15px}
.card-body{padding:16px 20px}
.id-table{width:100%;border-collapse:collapse}.id-table tr:not(:last-child) td{border-bottom:1px solid var(--border)}.id-table td{padding:8px 0;vertical-align:middle}
.id-label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);white-space:nowrap;padding-right:16px;width:140px}
.id-value-wrap{display:flex;align-items:center;gap:8px}.id-value{font-family:monospace;font-size:12px;word-break:break-all}
.copy-btn{flex-shrink:0;background:var(--surface3);border:1px solid var(--border);border-radius:4px;padding:2px 8px;font-size:11px;cursor:pointer;color:var(--text-secondary)}.copy-btn.copied{background:var(--success);color:#fff;border-color:var(--success)}
.evidence-list,.action-list,.doc-list{list-style:none}
.evidence-item,.action-item,.doc-item{display:flex;align-items:flex-start;gap:12px;padding:10px 0;border-bottom:1px solid var(--border)}.evidence-item:last-child,.action-item:last-child,.doc-item:last-child{border-bottom:none}
.evidence-icon{width:20px;height:20px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;flex-shrink:0;font-weight:700}
.evidence-icon.pass{background:#e6f4ea;color:var(--success)}.evidence-icon.fail{background:#fde7e9;color:var(--danger)}.evidence-icon.warn{background:#fff8e1;color:var(--warning)}.evidence-icon.unknown{background:var(--surface3);color:var(--text-muted)}
.evidence-label{font-size:13px;font-weight:600}.evidence-detail{font-size:12px;color:var(--text-secondary);margin-top:2px}
.action-num{min-width:24px;height:24px;background:var(--ms-blue);color:#fff;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0}
.action-num.pri-critical{background:var(--danger)}.action-num.pri-high{background:#da3b01}.action-num.pri-medium{background:var(--warning)}
.action-title{font-size:13px;font-weight:600}.action-body{font-size:12px;color:var(--text-secondary);margin-top:3px}
.geneva-links{display:flex;flex-wrap:wrap;gap:8px;padding:8px 0}
.geneva-link{display:inline-flex;align-items:center;gap:6px;padding:6px 14px;border-radius:4px;font-size:12px;font-weight:600;text-decoration:none;border:1px solid}
.geneva-link.log{background:#1e3a5f;color:#60cdff;border-color:#1e4a7a}.geneva-link.incoming{background:#1e3b2a;color:#6ccb7e;border-color:#1e5c30}.geneva-link.outgoing{background:#3b2a1e;color:#e8a86e;border-color:#5c3a1e}
.timeline{padding-left:24px;position:relative}.timeline::before{content:'';position:absolute;left:8px;top:0;bottom:0;width:2px;background:var(--border)}
.timeline-item{position:relative;margin-bottom:16px}.timeline-dot{position:absolute;left:-20px;top:4px;width:10px;height:10px;border-radius:50%;background:var(--ms-blue)}
.timeline-time{font-size:11px;color:var(--text-muted);font-family:monospace}.timeline-author{font-size:11px;font-weight:700;color:var(--ms-blue)}.timeline-content{font-size:13px;margin-top:2px}
.icm-table{width:100%;border-collapse:collapse}.icm-table th{background:var(--surface3);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);padding:8px 12px;text-align:left;border-bottom:2px solid var(--border)}.icm-table td{padding:10px 12px;border-bottom:1px solid var(--border);font-size:13px}
.icm-link{color:var(--ms-blue);text-decoration:none;font-weight:600}.doc-link{color:var(--ms-blue);font-size:13px;font-weight:500;text-decoration:none}.doc-desc{font-size:12px;color:var(--text-secondary);margin-top:2px}
.mb-20{margin-bottom:20px}.report-footer{margin-top:32px;padding-top:16px;border-top:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;font-size:11px;color:var(--text-muted)}
@media print{body{background:#fff}.header{background:#1b1b1b!important;print-color-adjust:exact}.print-btn{display:none}}</style></head>
<body>
<div class="header">
  <svg width="22" height="22" viewBox="0 0 24 24"><rect x="2" y="2" width="9" height="9" fill="#f25022"/><rect x="13" y="2" width="9" height="9" fill="#7fba00"/><rect x="2" y="13" width="9" height="9" fill="#00a4ef"/><rect x="13" y="13" width="9" height="9" fill="#ffb900"/></svg>
  <span style="font-size:15px;font-weight:600">Teams DRI</span>
  <span style="color:rgba(255,255,255,.3)">|</span>
  <span class="header-icm">ICM ${icmId}</span>
  <div class="hs"></div>
  <span class="header-meta">${new Date().toLocaleString()}</span>
  <button class="print-btn" onclick="window.print()">⎙ Export PDF</button>
</div>
<div class="page">
  <div class="hero">
    <div><span class="badge sev${sev}">SEV ${sev}</span><span class="badge" style="background:#fde7e9;color:#c50f1f">${String(d.status||'Unknown')}</span></div>
    <div class="hero-title">${String(d.title||'Unknown Incident')}</div>
    <div class="hero-meta-grid">
      <div class="meta-item"><label>ICM ID</label><span style="font-family:monospace">${icmId}</span></div>
      <div class="meta-item"><label>Created</label><span>${String(d.createDate||'N/A')}</span></div>
      <div class="meta-item"><label>Mitigated</label><span>${String(d.mitigateDate||'N/A')}</span></div>
      <div class="meta-item"><label>Owning Team</label><span>${String(d.owningTeam||'N/A')}</span></div>
      <div class="meta-item"><label>Forest / Region</label><span>${String(d.forest||'N/A')} / ${String(d.region||'N/A')}</span></div>
    </div>
  </div>
  <div class="rca-box"><h3>⚡ Root Cause Hypothesis</h3><p>${String(d.rootCauseHypothesis||'Under investigation')}</p></div>
  <div class="grid-2 mb-20">
    <div class="card"><div class="card-header"><div class="card-header-icon" style="background:#e8f4fd">🔑</div><h2>Resource Identifiers</h2></div><div class="card-body"><table class="id-table">${identifierRows}</table></div></div>
    <div class="card"><div class="card-header"><div class="card-header-icon" style="background:#e6f4ea">✅</div><h2>Evidence Checklist</h2></div><div class="card-body"><ul class="evidence-list">${evidence}</ul></div></div>
  </div>
  <div class="grid-2 mb-20">
    <div class="card"><div class="card-header"><div class="card-header-icon" style="background:#fff8e1">⚡</div><h2>Next Steps</h2></div><div class="card-body"><ol class="action-list">${nextSteps}</ol></div></div>
    <div class="card"><div class="card-header"><div class="card-header-icon" style="background:#1e1e1e;color:#60cdff">📊</div><h2>Geneva Log Links</h2></div><div class="card-body"><div class="geneva-links">${genevaHtml}</div></div></div>
  </div>
  <div class="grid-2 mb-20">
    <div class="card"><div class="card-header"><div class="card-header-icon" style="background:#f0f2f5">💬</div><h2>Discussion Timeline</h2></div><div class="card-body" style="max-height:360px;overflow-y:auto"><div class="timeline">${discussions}</div></div></div>
    <div class="card"><div class="card-header"><div class="card-header-icon" style="background:#fde7e9">🔗</div><h2>Related ICMs</h2></div><div class="card-body" style="padding:0"><table class="icm-table"><thead><tr><th>ICM</th><th>Title</th><th>Sev</th><th>Status</th><th>Date</th></tr></thead><tbody>${relatedRows}</tbody></table></div></div>
  </div>
  <div class="card mb-20"><div class="card-header"><div class="card-header-icon" style="background:#e8f4fd">📚</div><h2>Documentation &amp; TSGs</h2></div><div class="card-body"><ul class="doc-list">${docs}</ul></div></div>
  <div class="report-footer"><span>ICM <strong>${icmId}</strong> · Teams DRI Report</span><span>Generated by Claude DRI Agent · ${new Date().toLocaleString()}</span></div>
</div>
<script>function copyId(btn,text){navigator.clipboard.writeText(text).then(()=>{btn.textContent='Copied!';btn.classList.add('copied');setTimeout(()=>{btn.textContent='Copy';btn.classList.remove('copied')},2000)});}</script>
</body></html>`;

      res.setHeader('Content-Type', 'text/html');
      res.send(html);
      return;
    } catch (e) {
      console.error('[Report] Failed to render JSON report:', e);
    }
  }

  // Legacy: try old static HTML file
  const legacyPath = `/tmp/icm_${icmId}_report.html`;
  if (fs.existsSync(legacyPath)) {
    res.setHeader('Content-Type', 'text/html');
    res.send(fs.readFileSync(legacyPath, 'utf-8'));
    return;
  }

  res.status(404).send('<p style="font-family:system-ui;padding:32px;color:#666">Report not yet generated. Run a DRI investigation first.</p>');
});

app.post('/api/task/:planId/:taskId/cancel', (req, res) => {
  const { planId, taskId } = req.params;
  const cancelled = executor.cancelTask(planId, taskId);
  res.json(cancelled
    ? { success: true }
    : { error: `Task ${taskId} not found or not running` }
  );
});

app.post('/api/plan/:planId/cancel', (req, res) => {
  const { planId } = req.params;
  executor.cancelPlan(planId);
  executingPlans.delete(planId);
  res.json({ success: true });
});

app.get('/api/status', (_req, res) => {
  res.json({
    plans: [...plans.values()].map((p) => ({
      id: p.id,
      title: p.title,
      executing: executingPlans.has(p.id),
      tasks: p.tasks.map((t) => ({
        id: t.id,
        title: t.title,
        status: t.status,
        complexity: t.complexity,
      })),
    })),
    wsClients: wsManager.clientCount,
  });
});

app.get('/health', (_req, res) => {
  res.json({ status: 'ok', uptime: process.uptime() });
});

// ─── Ambient-MCP auto-token endpoint ─────────────────────────────────────────

app.get('/api/ambient/icm-token', (_req, res) => {
  const entry = readAmbientIcmToken();
  if (!entry) {
    res.status(404).json({ found: false, message: 'No ICM token found in ambient-mcp store.' });
    return;
  }
  // Auto-apply to server state so the caller doesn't need a second round-trip
  icmBearerToken = entry.token;
  icmCache = null;
  const alias = aliasFromJwt(entry.token);
  res.json({ found: true, token: entry.token, alias, domain: entry.domain, status: entry.status, capturedAt: entry.capturedAt });
});

// ─── ADX / Kusto ICM query endpoint ──────────────────────────────────────────

app.get('/api/adx/icm/:icmId', async (req, res) => {
  const { icmId } = req.params;
  if (!icmId || !/^\d+$/.test(icmId)) {
    res.status(400).json({ error: 'Invalid ICM ID — must be numeric' });
    return;
  }
  try {
    const kql = `IncidentDescriptions | where IncidentId == ${icmId}`;
    const rows = await runKustoQuery(kql);
    res.json({ icmId, query: kql, rows, rowCount: rows.length });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    res.status(500).json({ error: message });
  }
});

// ─── Microsoft Graph proxy ────────────────────────────────────────────────────
// Uses az cli token so sub-agents can call Graph without managing auth themselves
// Usage: GET /api/graph?path=/me  or  /api/graph?path=/me/photo/$value

async function getAzCliToken(resource: string): Promise<string> {
  const { execSync } = require('child_process') as typeof import('child_process');
  const raw = execSync(
    `az account get-access-token --resource ${resource} --query accessToken -o tsv`,
    { encoding: 'utf-8', timeout: 15000 }
  ).trim();
  if (!raw) throw new Error('az cli returned empty token — run: az login --tenant 72f988bf-86f1-41af-91ab-2d7cd011db47');
  return raw;
}

app.get('/api/graph', async (req, res) => {
  const graphPath = (req.query['path'] as string) || '/me';
  if (!graphPath.startsWith('/')) {
    res.status(400).json({ error: 'path must start with /' });
    return;
  }
  try {
    const token = await getAzCliToken('https://graph.microsoft.com');
    const url = `https://graph.microsoft.com/v1.0${graphPath}`;
    const resp = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
    const body = await resp.text();
    res.status(resp.status).setHeader('Content-Type', resp.headers.get('content-type') ?? 'application/json').send(body);
  } catch (err) {
    res.status(500).json({ error: (err as Error).message });
  }
});

// ─── Chat persistence endpoints ───────────────────────────────────────────────

app.get('/api/chat', (_req, res) => {
  res.json(loadChat());
});

app.post('/api/chat', (req, res) => {
  const { messages } = req.body as { messages?: { role: string; text: string }[] };
  if (!Array.isArray(messages)) { res.status(400).json({ error: 'messages array required' }); return; }
  saveChat(messages);
  res.json({ success: true });
});

// ─── Activity summary endpoints ───────────────────────────────────────────────

app.get('/api/activity-summary', (_req, res) => {
  try {
    const history: ActivitySummary[] = fs.existsSync(SUMMARY_FILE)
      ? JSON.parse(fs.readFileSync(SUMMARY_FILE, 'utf-8')) : [];
    res.json({ latest: latestSummary, history });
  } catch {
    res.json({ latest: latestSummary, history: [] });
  }
});

app.post('/api/activity-summary/generate', (_req, res) => {
  res.json({ success: true, message: 'Generation started' });
  generateActivitySummary().catch(console.error);
});

// ─── Start ────────────────────────────────────────────────────────────────────

httpServer.listen(PORT, () => {
  console.log(`🤖 Claude Orchestrator running at http://localhost:${PORT}`);
  console.log(`   WebSocket ready on ws://localhost:${PORT}`);

  // Auto-load ICM token from ambient-mcp on startup
  const ambientEntry = readAmbientIcmToken();
  if (ambientEntry) {
    icmBearerToken = ambientEntry.token;
    console.log(`[Ambient] ICM token auto-loaded from ${ambientEntry.domain} (${ambientEntry.status})`);
  }
  // Refresh ambient token every 10 minutes silently
  setInterval(() => {
    const fresh = readAmbientIcmToken();
    if (fresh && fresh.token !== icmBearerToken) {
      icmBearerToken = fresh.token;
      icmCache = null;
      console.log(`[Ambient] ICM token refreshed from ${fresh.domain}`);
      wsManager.broadcast({ type: 'ambient_token_refreshed', domain: fresh.domain, capturedAt: fresh.capturedAt } as never);
    }
  }, 10 * 60 * 1000);

  // Generate first activity summary 15s after startup, then every hour
  setTimeout(() => generateActivitySummary().catch(console.error), 15000);
  setInterval(() => generateActivitySummary().catch(console.error), 60 * 60 * 1000);
});

process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully...');
  httpServer.close(() => process.exit(0));
});
