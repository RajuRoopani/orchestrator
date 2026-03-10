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

  const plan = createDriPlan(message.trim());
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
  const reportPath = `/tmp/icm_${icmId}_report.html`;
  if (!fs.existsSync(reportPath)) {
    res.status(404).send('<p>Report not yet generated.</p>');
    return;
  }
  res.setHeader('Content-Type', 'text/html');
  res.send(fs.readFileSync(reportPath, 'utf-8'));
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

// ─── Start ────────────────────────────────────────────────────────────────────

httpServer.listen(PORT, () => {
  console.log(`🤖 Claude Orchestrator running at http://localhost:${PORT}`);
  console.log(`   WebSocket ready on ws://localhost:${PORT}`);
});

process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully...');
  httpServer.close(() => process.exit(0));
});
