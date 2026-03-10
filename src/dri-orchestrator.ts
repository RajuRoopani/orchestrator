import { v4 as uuidv4 } from 'uuid';
import * as os from 'os';
import { Plan } from './models';

/** Pull a bare ICM number out of free text, e.g. "ICM 750444345" or "750444345" */
function extractIcmId(text: string): string | null {
  const m = text.match(/\b(\d{7,12})\b/);
  return m ? m[1] : null;
}

export function createDriPlan(incidentDescription: string): Plan {
  const planId   = uuidv4();
  const icmId    = extractIcmId(incidentDescription) ?? 'UNKNOWN';
  const incident = incidentDescription.slice(0, 2000);

  // All steps share one workspace directory so they can exchange JSON files
  const sharedWs = `${os.homedir()}/Work6/experiments/orchestrator-workspaces/dri-${planId}`;

  const ctx = `
**ICM ID:** ${icmId}
**Incident:** ${incident}
**Shared workspace:** ${sharedWs}
  `.trim();

  return {
    id:                 planId,
    type:               'dri',
    icmId,
    title:              icmId !== 'UNKNOWN' ? `DRI: ICM ${icmId}` : 'DRI Investigation',
    description:        incidentDescription.slice(0, 120) + (incidentDescription.length > 120 ? '…' : ''),
    estimated_duration: '10–15 minutes',
    createdAt:          Date.now(),
    tasks: [

      // ── STEP 1 ── Load ICM (sequential, no deps) ──────────────────────────
      {
        id:           'step_1_load_incident',
        title:        'Load ICM Incident Details',
        description:  'Query IcmDataWarehouse for incident info and discussion timeline',
        complexity:   'low',
        dependencies: [],
        parallel_ok:  false,
        allowAllTools: true,
        status:  'pending',
        output:  '',
        workspace: sharedWs,
        claude_prompt: `You are a DRI (Designated Responsible Individual) engineer on Microsoft Teams / Copilot Studio on-call.

${ctx}

## Step 1: Load Incident from Kusto (IcmDataWarehouse)

Also query the **kusto-icm** MCP server (IcmDataWarehouse cluster) for structured data:

\`\`\`kusto
Incidents
| where IncidentId == ${icmId}
| project IncidentId, Title, Severity, Status, OwningTeamName, CreateDate, MitigateDate, ResolveDate, Summary, Keywords, RoutingId
\`\`\`

Pull the discussion timeline:
\`\`\`kusto
Incidents
| where IncidentId == ${icmId}
| join kind=inner IncidentDiscussions on IncidentId
| project CreateDate, Author, Text
| order by CreateDate asc
\`\`\`

Merge dashboard data and Kusto data — prefer dashboard values where both exist.

If the ICM ID is UNKNOWN or both sources return no results, derive context from:
"${incident}"

Write your findings to **${sharedWs}/step1_incident.json**:
\`\`\`json
{
  "icmId": "",
  "title": "",
  "severity": "",
  "status": "",
  "owningTeam": "",
  "owningService": "",
  "owner": "",
  "tags": [],
  "duration": "",
  "createDate": "",
  "mitigateDate": "",
  "resolveDate": "",
  "summary": "",
  "aiSummary": "",
  "authoredSummary": "",
  "keywords": [],
  "discussions": [{"time":"","author":"","text":""}],
  "troubleshootingNotes": "",
  "dashboardUrl": "https://portal.microsofticm.com/imp/v5/incidents/details/${icmId}/summary",
  "forest": "",
  "region": ""
}
\`\`\`

Output a concise summary of findings including the incident title, status, severity, owning team, and key points from the AI summary.`,
      },

      // ── STEP 2 ── Extract Identifiers (deps: step 1) ──────────────────────
      {
        id:           'step_2_extract_ids',
        title:        'Extract Resource Identifiers',
        description:  'Extract TenantId, BotId, ThreadId, CorrelationId, UserId from incident',
        complexity:   'low',
        dependencies: ['step_1_load_incident'],
        parallel_ok:  false,
        allowAllTools: true,
        status:  'pending',
        output:  '',
        workspace: sharedWs,
        claude_prompt: `You are a DRI engineer on Microsoft Teams / Copilot Studio on-call.

${ctx}

## Step 2: Extract Resource Identifiers

Read **${sharedWs}/step1_incident.json** for context.

From the incident title, summary, keywords, or any HAR/SAZ file path mentioned by the user, extract:

- **TenantId** — GUID \`xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx\`
- **UserId / OID** — GUID
- **MeetingId / ThreadId / ConversationId**
- **BotId / AppId**
- **ResourceId** or **ServiceTreeId**
- **CorrelationId / RequestId**
- **ClientSessionId**
- **Timestamp** — exact datetime of first failure

If the user provided a HAR or SAZ file path, use the Read tool to open it and search for:
- \`x-ms-correlation-id\` headers
- \`x-ms-client-request-id\` headers
- URL path segments containing GUIDs
- \`tenantId=\`, \`userId=\`, \`oid=\` query params

Note the **Forest** and **Region** from step1 — needed for later Kusto queries.

Incident: "${incident}"

Write findings to **${sharedWs}/step2_identifiers.json**:
\`\`\`json
{
  "tenantId": "",
  "userId": "",
  "threadId": "",
  "botId": "",
  "correlationId": "",
  "clientSessionId": "",
  "resourceId": "",
  "timestamp": "",
  "forest": "",
  "region": ""
}
\`\`\`

Output a structured list of all identifiers found, and clearly note any that are missing.`,
      },

      // ── STEPS 3+4+5 ── Kusto + ACL + Hot Shard (deps: step 2, sequential) ─
      {
        id:           'step_3_kusto_acl_shard',
        title:        'Kusto · Tenant Telemetry + ACL + Hot Shard',
        description:  'Query tenant anomalies, ACL permissions, and shard health (Steps 3–5)',
        complexity:   'high',
        dependencies: ['step_2_extract_ids'],
        parallel_ok:  false,
        allowAllTools: true,
        status:  'pending',
        output:  '',
        workspace: sharedWs,
        claude_prompt: `You are a DRI engineer on Microsoft Teams / Copilot Studio on-call.

${ctx}

Read **${sharedWs}/step2_identifiers.json** for TenantId, Forest, Region, and Timestamp.

---
## Step 3: Kusto Investigation — Tenant & Forest Info

Use the **kusto-spoons** MCP server (o365monitoring cluster):

\`\`\`kusto
TeamsAnalyticsEvents
| where TenantId == "<TenantId>"
| where Timestamp between (datetime(<incident_time> - 1h) .. datetime(<incident_time> + 1h))
| summarize EventCount=count(), ErrorCount=countif(IsError==true) by bin(Timestamp, 5m), EventType
| order by Timestamp asc
\`\`\`

\`\`\`kusto
ForestInfo
| where TenantId == "<TenantId>"
| project TenantId, Forest, DataResidencyRegion, IsGovernment, CloudType
\`\`\`

---
## Step 4: Check ACL Permissions

Use the **kusto-icm** MCP server (IcmDataWarehouse):
\`\`\`kusto
Incidents
| where Keywords contains "<TenantId>"
| where Title contains "ACL" or Title contains "permission" or Title contains "access"
| where CreateDate > ago(7d)
| project IncidentId, Title, Status, CreateDate, Severity
\`\`\`

Also use the **ado-ecs** MCP server to check recent ECS config changes or experiments targeting the affected tenant or forest.

---
## Step 5: Hot Shard Analysis

Use the **kusto-spoons** MCP server (o365monitoring):
\`\`\`kusto
ShardHealthMetrics
| where Timestamp > ago(2h)
| where TenantId == "<TenantId>" or ShardId in (
    TenantShardMapping
    | where TenantId == "<TenantId>"
    | project ShardId
)
| summarize AvgLatencyMs=avg(LatencyMs), P99LatencyMs=percentile(LatencyMs, 99), RequestCount=count()
    by bin(Timestamp, 5m), ShardId
| where AvgLatencyMs > 500 or P99LatencyMs > 2000
| order by Timestamp desc
\`\`\`

\`\`\`kusto
ShardRebalanceEvents
| where Timestamp > ago(24h)
| where TenantId == "<TenantId>"
| project Timestamp, EventType, SourceShard, DestinationShard, Status
\`\`\`

Write findings to **${sharedWs}/step3_kusto.json**:
\`\`\`json
{
  "tenantAnomalies": [],
  "forestInfo": {},
  "aclIssues": [],
  "ecsChanges": [],
  "shardHealth": { "isHotShard": false, "avgLatencyMs": 0, "p99LatencyMs": 0 },
  "shardRebalancing": []
}
\`\`\`

Output a clear summary of tenant anomalies, ACL issues, and shard health.`,
      },

      // ── STEPS 6+7 ── ADO Work Items + TSGs (deps: step 2, parallel) ────────
      {
        id:           'step_4_ado_tsgs',
        title:        'ADO Work Items & Troubleshooting Guides',
        description:  'Search O365Exchange ADO for related bugs, hotfixes, and TSGs (Steps 6–7)',
        complexity:   'medium',
        dependencies: ['step_2_extract_ids'],
        parallel_ok:  true,
        allowAllTools: true,
        status:  'pending',
        output:  '',
        workspace: sharedWs,
        claude_prompt: `You are a DRI engineer on Microsoft Teams / Copilot Studio on-call.

${ctx}

Read **${sharedWs}/step2_identifiers.json** for TenantId, CorrelationId, and error context.

---
## Step 6: Check for Related ADO Work Items

Use the **ado** MCP server to search **O365Exchange** for:
- Active bugs containing the TenantId or ICM ID **${icmId}**
- Recent deployments to the affected service/forest around the incident time
- Any known hotfixes currently in flight

---
## Step 7: Check for Related TSG or .md files

Use the **ado** MCP server to search **O365Exchange** repositories for:
- Troubleshooting guides (TSGs) or README files related to the incident error pattern
- Any existing workarounds documented for similar issues
- Suggested action items from previous hotfixes

Write findings to **${sharedWs}/step4_ado.json**:
\`\`\`json
{
  "bugs": [{"id":"","title":"","status":"","url":"","description":""}],
  "deployments": [{"date":"","service":"","version":"","forest":""}],
  "hotfixes": [{"id":"","title":"","url":"","status":""}],
  "tsgs": [{"title":"","url":"","summary":"","actionItems":[]}]
}
\`\`\`

Output a clear summary of related work items and applicable TSGs.`,
      },

      // ── STEP 8 ── Geneva Log URLs (deps: step 2, parallel) ─────────────────
      {
        id:           'step_5_geneva_logs',
        title:        'Generate Geneva Log URLs',
        description:  'Build Geneva log deep-links for LogMessage, IncomingRequest, OutgoingRequest (Step 8)',
        complexity:   'low',
        dependencies: ['step_2_extract_ids'],
        parallel_ok:  true,
        allowAllTools: true,
        status:  'pending',
        output:  '',
        workspace: sharedWs,
        claude_prompt: `You are a DRI engineer on Microsoft Teams / Copilot Studio on-call.

${ctx}

Read **${sharedWs}/step2_identifiers.json** for ThreadId and Timestamp.

---
## Step 8: Generate Geneva Log URLs

Generate Geneva log deep-links for **LogMessage**, **IncomingRequest**, and **OutgoingRequest** using the exact SMBA/APX template below.

Replace:
- \`{TIMESTAMP}\` → the incident timestamp in ISO 8601 format (e.g. \`2024-03-15T14:30:00\`)
- \`{THREAD_ID}\` → the extracted ThreadId / ConversationId

**URL template:**
\`\`\`
https://portal.microsoftgeneva.com/logs/dgrep?be=DGrep&time={TIMESTAMP}&offset=~5&offsetUnit=Minutes&UTC=true&ep=Diagnostics%20PROD&ns=SkypeSMB&en=IncomingRequest,LogMessage,OutgoingRequest&conditions=[["AnyField","contains","{THREAD_ID}"]]&clientQuery=orderby%20PreciseTimeStamp%20asc%0Awhere%20it.any("LogErrorResponseBodyHandler")%20or%20Level%20%3D%3D%202&chartEditorVisible=true&chartType=line&chartLayers=[["New%20Layer",""]]
\`\`\`

Generate three separate specialised URLs (one per log type — adjust the \`en=\` parameter):
1. **LogMessage** — \`en=LogMessage\`
2. **IncomingRequest** — \`en=IncomingRequest\`
3. **OutgoingRequest** — \`en=OutgoingRequest\`

Write findings to **${sharedWs}/step5_geneva.json**:
\`\`\`json
{
  "timestamp": "",
  "threadId": "",
  "logMessageUrl": "",
  "incomingRequestUrl": "",
  "outgoingRequestUrl": "",
  "combinedUrl": ""
}
\`\`\`

Output all four URLs clearly so the on-call engineer can click them immediately.`,
      },

      // ── STEPS 9+10 ── Related ICMs + Docs (deps: step 2, parallel) ─────────
      {
        id:           'step_6_related_icms_docs',
        title:        'Related ICMs & Microsoft Documentation',
        description:  'Find related incidents, mitigation plans, and learn.microsoft.com resources (Steps 9–10)',
        complexity:   'medium',
        dependencies: ['step_2_extract_ids'],
        parallel_ok:  true,
        allowAllTools: true,
        status:  'pending',
        output:  '',
        workspace: sharedWs,
        claude_prompt: `You are a DRI engineer on Microsoft Teams / Copilot Studio on-call.

${ctx}

Read **${sharedWs}/step2_identifiers.json** for TenantId and error context.

---
## Step 9: Check for Related ICMs and Hot Fixes

Query the **kusto-icm** MCP server (IcmDataWarehouse):
\`\`\`kusto
Incidents
| where Keywords contains "<TenantId>" or Title contains "<ErrorPattern>"
| where CreateDate > ago(30d)
| project IncidentId, Title, Severity, Status, OwningTeamName, CreateDate, MitigateDate, ResolveDate, Summary
| order by CreateDate desc
| take 10
\`\`\`

Also search for related hotfix PRs and mitigation plans via the **ado** MCP server.

---
## Step 10: Check for Related Microsoft Public Documentation

Use the **ado** MCP server to search the wiki for documentation matching the error patterns, error codes, or service names found in the incident.
Search for terms like the error pattern, service name, or known TSG titles.

Write findings to **${sharedWs}/step6_related.json**:
\`\`\`json
{
  "relatedIcms": [
    {"id":"","title":"","severity":"","status":"","createDate":"","mitigateDate":"","resolution":""}
  ],
  "hotfixes": [{"id":"","title":"","url":"","status":""}],
  "docs": [
    {"title":"","url":"","description":""}
  ]
}
\`\`\`

Output a clear summary of related incidents and documentation links.`,
      },

      // ── STEP 11 ── Generate HTML Report (deps: steps 3–6) ──────────────────
      {
        id:           'step_7_generate_report',
        title:        'Generate Rich HTML Investigation Report',
        description:  'Compile all findings into a self-contained HTML report and open in browser (Step 11)',
        complexity:   'high',
        dependencies: ['step_3_kusto_acl_shard', 'step_4_ado_tsgs', 'step_5_geneva_logs', 'step_6_related_icms_docs'],
        parallel_ok:  false,
        allowAllTools: true,
        status:  'pending',
        output:  '',
        workspace: sharedWs,
        claude_prompt: `You are a DRI engineer completing a full incident investigation for Microsoft Teams / Copilot Studio.

${ctx}

---
## Step 11: Generate Rich HTML Report

Read ALL findings files from **${sharedWs}/**:
- \`step1_incident.json\`   — ICM details & discussion timeline
- \`step2_identifiers.json\` — TenantId, BotId, ThreadId, CorrelationId, etc.
- \`step3_kusto.json\`       — Tenant telemetry, ACL issues, shard health
- \`step4_ado.json\`         — ADO bugs, deployments, hotfixes, TSGs
- \`step5_geneva.json\`      — Geneva log URLs
- \`step6_related.json\`     — Related ICMs and documentation

Synthesise all findings and write a **complete self-contained HTML file** to:
\`/tmp/icm_${icmId}_report.html\`

The HTML MUST use the exact template below. Replace **every** \`{{PLACEHOLDER}}\` with actual data.
Use \`N/A\` or \`Unknown\` for missing fields — never leave a placeholder unfilled.

\`\`\`html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ICM {{ICM_ID}} — DRI Report</title>
<style>
:root {
  --sev1:#c50f1f;--sev1-bg:#fde7e9;--sev2:#da3b01;--sev2-bg:#fde7e9;
  --sev3:#c19c00;--sev3-bg:#fff8e1;--sev4:#0078d4;--sev4-bg:#e8f4fd;
  --active:#c50f1f;--active-bg:#fde7e9;--mitigated:#da3b01;--mitigated-bg:#fff4e5;
  --resolved:#107c10;--resolved-bg:#e6f4ea;--ms-blue:#0078d4;--ms-dark:#1b1b1b;
  --surface:#fff;--surface2:#f8f9fa;--surface3:#f0f2f5;--border:#e1e4e8;
  --text-primary:#1b1b1b;--text-secondary:#616161;--text-muted:#8a8a8a;
  --success:#107c10;--warning:#c19c00;--danger:#c50f1f;
  --radius:8px;--radius-lg:12px;
  --shadow:0 2px 8px rgba(0,0,0,.08);--shadow-lg:0 4px 20px rgba(0,0,0,.12);
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,system-ui,sans-serif;background:var(--surface3);color:var(--text-primary);line-height:1.5;font-size:14px}
.header{background:var(--ms-dark);color:#fff;padding:0 32px;height:56px;display:flex;align-items:center;gap:16px;position:sticky;top:0;z-index:100;box-shadow:0 1px 4px rgba(0,0,0,.3)}
.header-logo{display:flex;align-items:center;gap:10px;font-size:15px;font-weight:600}
.hd{width:1px;height:24px;background:rgba(255,255,255,.2)}
.header-title{font-size:14px;color:rgba(255,255,255,.7)}
.header-icm{font-size:14px;font-weight:700;color:#60cdff}
.hs{flex:1}
.header-meta{font-size:12px;color:rgba(255,255,255,.5)}
.print-btn{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:#fff;padding:5px 14px;border-radius:4px;font-size:12px;cursor:pointer}
.print-btn:hover{background:rgba(255,255,255,.18)}
.page{max-width:1280px;margin:0 auto;padding:24px 24px 48px}
.hero{background:var(--surface);border-radius:var(--radius-lg);box-shadow:var(--shadow);padding:24px 28px;margin-bottom:20px;border-top:4px solid var(--sev-color,var(--ms-blue))}
.hero-badges{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:12px}
.hero-title{font-size:20px;font-weight:700;line-height:1.3}
.hero-meta-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:16px;margin-top:16px;padding-top:16px;border-top:1px solid var(--border)}
.meta-item label{display:block;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--text-muted);margin-bottom:3px}
.meta-item span{font-size:13px;font-weight:500}
.mono{font-family:'Cascadia Code',Consolas,monospace;font-size:12px}
.badge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:100px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px}
.badge-dot{width:6px;height:6px;border-radius:50%;background:currentColor}
.badge-sev1,.badge-sev2{color:var(--sev1);background:var(--sev1-bg)}
.badge-sev3{color:var(--sev3);background:var(--sev3-bg)}
.badge-sev4{color:var(--sev4);background:var(--sev4-bg)}
.badge-active{color:var(--active);background:var(--active-bg)}
.badge-mitigated{color:var(--mitigated);background:var(--mitigated-bg)}
.badge-resolved{color:var(--resolved);background:var(--resolved-bg)}
.badge-info{color:var(--ms-blue);background:var(--sev4-bg)}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
.badge-pulse{animation:pulse 1.8s infinite}
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}
@media(max-width:900px){.grid-2{grid-template-columns:1fr}}
.card{background:var(--surface);border-radius:var(--radius-lg);box-shadow:var(--shadow);overflow:hidden;margin-bottom:0}
.card-header{padding:14px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.card-header-icon{width:28px;height:28px;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:15px}
.card-header h2{font-size:13px;font-weight:700;flex:1}
.card-count{font-size:11px;background:var(--surface3);border-radius:100px;padding:2px 8px;color:var(--text-secondary);font-weight:600}
.card-body{padding:16px 20px}
.id-table{width:100%;border-collapse:collapse}
.id-table tr:not(:last-child) td{border-bottom:1px solid var(--border)}
.id-table td{padding:8px 0;vertical-align:middle}
.id-label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);white-space:nowrap;padding-right:16px;width:150px}
.id-value-wrap{display:flex;align-items:center;gap:8px}
.id-value{font-family:'Cascadia Code',Consolas,monospace;font-size:12px;word-break:break-all}
.copy-btn{flex-shrink:0;background:var(--surface3);border:1px solid var(--border);border-radius:4px;padding:2px 8px;font-size:11px;cursor:pointer;color:var(--text-secondary);transition:all .15s;white-space:nowrap}
.copy-btn:hover{background:var(--ms-blue);color:#fff;border-color:var(--ms-blue)}
.copy-btn.copied{background:var(--success);color:#fff;border-color:var(--success)}
.stack-trace{background:#1e1e1e;padding:16px 18px;overflow:auto;font-family:'Cascadia Code',Consolas,monospace;font-size:12px;line-height:1.7;max-height:420px}
.line-error{color:#f48771;display:block}.line-warn{color:#cca700;display:block}.line-info{color:#9cdcfe;display:block}
.line-at{color:#c586c0;display:block}.line-normal{color:#d4d4d4;display:block}
.line-highlight{color:#ffff00;background:rgba(255,255,0,.08);display:block;font-weight:700}
.timeline{position:relative;padding-left:28px}
.timeline::before{content:'';position:absolute;left:8px;top:4px;bottom:4px;width:2px;background:var(--border)}
.timeline-item{position:relative;margin-bottom:20px}
.timeline-item:last-child{margin-bottom:0}
.timeline-dot{position:absolute;left:-24px;top:4px;width:12px;height:12px;border-radius:50%;background:var(--ms-blue);border:2px solid var(--surface);box-shadow:0 0 0 2px var(--ms-blue)}
.timeline-dot.error{background:var(--danger);box-shadow:0 0 0 2px var(--danger)}
.timeline-dot.success{background:var(--success);box-shadow:0 0 0 2px var(--success)}
.timeline-dot.warn{background:var(--warning);box-shadow:0 0 0 2px var(--warning)}
.timeline-time{font-size:11px;color:var(--text-muted);font-family:monospace;margin-bottom:2px}
.timeline-author{font-size:11px;font-weight:700;color:var(--ms-blue)}
.timeline-content{font-size:13px;margin-top:3px}
.timeline-content code{background:var(--surface3);padding:1px 5px;border-radius:3px;font-size:11px;font-family:'Cascadia Code',Consolas,monospace}
.evidence-list{list-style:none}
.evidence-item{display:flex;align-items:flex-start;gap:12px;padding:10px 0;border-bottom:1px solid var(--border)}
.evidence-item:last-child{border-bottom:none}
.evidence-icon{width:20px;height:20px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;flex-shrink:0;margin-top:1px;font-weight:700}
.evidence-icon.pass{background:var(--resolved-bg);color:var(--success)}
.evidence-icon.fail{background:var(--active-bg);color:var(--danger)}
.evidence-icon.warn{background:var(--sev3-bg);color:var(--warning)}
.evidence-icon.unknown{background:var(--surface3);color:var(--text-muted)}
.evidence-label{font-size:13px;font-weight:600}
.evidence-detail{font-size:12px;color:var(--text-secondary);margin-top:2px}
.action-list{list-style:none}
.action-item{display:flex;align-items:flex-start;gap:12px;padding:12px 0;border-bottom:1px solid var(--border)}
.action-item:last-child{border-bottom:none}
.action-num{min-width:24px;height:24px;background:var(--ms-blue);color:#fff;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0;margin-top:1px}
.action-num.pri-critical{background:var(--danger)}
.action-num.pri-high{background:var(--mitigated)}
.action-num.pri-medium{background:var(--warning)}
.action-title{font-size:13px;font-weight:600}
.action-body{font-size:12px;color:var(--text-secondary);margin-top:3px}
.action-body code{background:var(--surface3);padding:1px 5px;border-radius:3px;font-family:'Cascadia Code',Consolas,monospace;font-size:11px}
.action-tag{display:inline-block;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;padding:1px 7px;border-radius:100px;margin-left:6px;vertical-align:middle}
.tag-customer{background:#e8f4fd;color:#0078d4}
.tag-oncall{background:#fde7e9;color:#c50f1f}
.tag-platform{background:#fff8e1;color:#c19c00}
.geneva-grid{display:flex;flex-direction:column;gap:10px}
.geneva-card{background:var(--surface3);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px}
.geneva-label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);margin-bottom:6px}
.geneva-timestamp{font-family:monospace;font-size:11px;background:var(--surface);border:1px solid var(--border);border-radius:4px;padding:2px 8px;color:var(--ms-blue);display:inline-block;margin-bottom:8px}
.geneva-links{display:flex;flex-wrap:wrap;gap:8px}
.geneva-link{display:inline-flex;align-items:center;gap:6px;padding:6px 14px;border-radius:4px;font-size:12px;font-weight:600;text-decoration:none;transition:all .15s;border:1px solid}
.geneva-link.log{background:#1e3a5f;color:#60cdff;border-color:#1e4a7a}
.geneva-link.log:hover{background:#1a3d6e}
.geneva-link.incoming{background:#1e3b2a;color:#6ccb7e;border-color:#1e5c30}
.geneva-link.incoming:hover{background:#1a4a28}
.geneva-link.outgoing{background:#3b2a1e;color:#e8a86e;border-color:#5c3a1e}
.geneva-link.outgoing:hover{background:#4a2a1a}
.icm-table{width:100%;border-collapse:collapse}
.icm-table th{background:var(--surface3);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);padding:8px 12px;text-align:left;border-bottom:2px solid var(--border)}
.icm-table td{padding:10px 12px;border-bottom:1px solid var(--border);font-size:13px}
.icm-table tr:hover td{background:var(--surface3)}
.icm-link{color:var(--ms-blue);text-decoration:none;font-weight:600}
.icm-link:hover{text-decoration:underline}
.rca-box{background:linear-gradient(135deg,#fff8e1 0%,#fde7e9 100%);border:1px solid #f0c060;border-left:4px solid var(--warning);border-radius:var(--radius);padding:16px 18px;margin-bottom:20px}
.rca-box h3{font-size:13px;font-weight:700;color:var(--warning);margin-bottom:8px}
.rca-box p{font-size:13px;line-height:1.6}
.doc-list{list-style:none}
.doc-item{padding:9px 0;border-bottom:1px solid var(--border);display:flex;align-items:flex-start;gap:10px}
.doc-item:last-child{border-bottom:none}
.doc-link{color:var(--ms-blue);font-size:13px;font-weight:500;text-decoration:none}
.doc-link:hover{text-decoration:underline}
.doc-desc{font-size:12px;color:var(--text-secondary);margin-top:2px}
.escalation-path{display:flex;align-items:center;flex-wrap:wrap;gap:4px;margin-top:8px}
.escalation-team{background:var(--surface3);border:1px solid var(--border);border-radius:var(--radius);padding:8px 14px;font-size:12px;font-weight:600;display:flex;flex-direction:column;gap:2px}
.escalation-team small{font-size:10px;color:var(--text-muted);font-weight:400}
.escalation-arrow{font-size:18px;color:var(--text-muted);padding:0 4px}
.escalation-notes{margin-top:14px;font-size:13px;color:var(--text-secondary)}
.mb-20{margin-bottom:20px}
.report-footer{margin-top:32px;padding-top:16px;border-top:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;font-size:11px;color:var(--text-muted)}
@media print{body{background:#fff}.header{background:#1b1b1b!important;print-color-adjust:exact}.card,.hero{box-shadow:none;border:1px solid var(--border)}.print-btn{display:none}}
</style>
</head>
<body>
<header class="header">
  <div class="header-logo">
    <svg width="22" height="22" viewBox="0 0 24 24"><rect x="2" y="2" width="9" height="9" fill="#f25022"/><rect x="13" y="2" width="9" height="9" fill="#7fba00"/><rect x="2" y="13" width="9" height="9" fill="#00a4ef"/><rect x="13" y="13" width="9" height="9" fill="#ffb900"/></svg>
    Teams DRI
  </div>
  <div class="hd"></div>
  <span class="header-title">Incident Report</span>
  <span class="header-icm">ICM {{ICM_ID}}</span>
  <div class="hs"></div>
  <span class="header-meta">{{REPORT_TIMESTAMP}}</span>
  <button class="print-btn" onclick="window.print()">⎙ Export PDF</button>
</header>
<div class="page">
  <div class="hero" style="--sev-color:{{SEV_COLOR}}">
    <div class="hero-badges">
      <span class="badge badge-sev{{SEVERITY}}"><span class="badge-dot"></span>SEV {{SEVERITY}}</span>
      <span class="badge badge-{{STATUS_CLASS}}"><span class="badge-dot {{STATUS_PULSE}}"></span>{{STATUS}}</span>
      {{EXTRA_TAGS}}
    </div>
    <div class="hero-title">{{INCIDENT_TITLE}}</div>
    <div class="hero-meta-grid">
      <div class="meta-item"><label>ICM ID</label><span class="mono">{{ICM_ID}}</span></div>
      <div class="meta-item"><label>Created</label><span>{{CREATE_DATE}}</span></div>
      <div class="meta-item"><label>Age</label><span>{{INCIDENT_AGE}}</span></div>
      <div class="meta-item"><label>Mitigated</label><span>{{MITIGATE_DATE}}</span></div>
      <div class="meta-item"><label>Resolved</label><span>{{RESOLVE_DATE}}</span></div>
      <div class="meta-item"><label>Owning Team</label><span>{{OWNING_TEAM}}</span></div>
      <div class="meta-item"><label>Forest / Region</label><span>{{FOREST}} / {{REGION}}</span></div>
      <div class="meta-item"><label>Last Updated</label><span>{{LAST_UPDATED}}</span></div>
    </div>
  </div>
  <div class="rca-box">
    <h3>⚡ Root Cause Hypothesis</h3>
    <p>{{ROOT_CAUSE_HYPOTHESIS}}</p>
  </div>
  <div class="grid-2 mb-20">
    <div class="card">
      <div class="card-header"><div class="card-header-icon" style="background:#e8f4fd">🔑</div><h2>Resource Identifiers</h2></div>
      <div class="card-body"><table class="id-table">{{IDENTIFIER_ROWS}}</table></div>
    </div>
    <div class="card">
      <div class="card-header"><div class="card-header-icon" style="background:#fde7e9">🔥</div><h2>Error Stack Trace</h2><span class="badge badge-active" style="font-size:10px">{{ERROR_CODE}}</span></div>
      <div class="card-body" style="padding:0"><div class="stack-trace">{{STACK_TRACE_HTML}}</div></div>
    </div>
  </div>
  <div class="grid-2 mb-20">
    <div class="card">
      <div class="card-header"><div class="card-header-icon" style="background:#e6f4ea">✅</div><h2>Evidence Checklist</h2></div>
      <div class="card-body"><ul class="evidence-list">{{EVIDENCE_ITEMS}}</ul></div>
    </div>
    <div class="card">
      <div class="card-header"><div class="card-header-icon" style="background:#fff8e1">⚡</div><h2>Recommended Next Steps</h2></div>
      <div class="card-body"><ol class="action-list">{{ACTION_ITEMS}}</ol></div>
    </div>
  </div>
  <div class="grid-2 mb-20">
    <div class="card">
      <div class="card-header"><div class="card-header-icon" style="background:#1e1e1e;color:#60cdff">📊</div><h2>Geneva Log Links</h2></div>
      <div class="card-body"><div class="geneva-grid">{{GENEVA_SECTIONS}}</div></div>
    </div>
    <div class="card">
      <div class="card-header"><div class="card-header-icon" style="background:#f0f2f5">💬</div><h2>Discussion Timeline</h2><span class="card-count">{{DISCUSSION_COUNT}} entries</span></div>
      <div class="card-body" style="max-height:380px;overflow-y:auto"><div class="timeline">{{TIMELINE_ITEMS}}</div></div>
    </div>
  </div>
  <div class="grid-2 mb-20">
    <div class="card">
      <div class="card-header"><div class="card-header-icon" style="background:#fde7e9">🔗</div><h2>Related ICMs</h2><span class="card-count">{{RELATED_ICM_COUNT}}</span></div>
      <div class="card-body" style="padding:0">
        <table class="icm-table">
          <thead><tr><th>ICM</th><th>Title</th><th>Sev</th><th>Status</th><th>Date</th></tr></thead>
          <tbody>{{RELATED_ICM_ROWS}}</tbody>
        </table>
      </div>
    </div>
    <div class="card">
      <div class="card-header"><div class="card-header-icon" style="background:#e8f4fd">📚</div><h2>Documentation &amp; TSG</h2></div>
      <div class="card-body"><ul class="doc-list">{{DOC_ITEMS}}</ul></div>
    </div>
  </div>
  <div class="card mb-20">
    <div class="card-header"><div class="card-header-icon" style="background:#fde7e9">🚨</div><h2>Escalation Path</h2></div>
    <div class="card-body">
      <div class="escalation-path">{{ESCALATION_STEPS}}</div>
      <div class="escalation-notes">{{ESCALATION_NOTES}}</div>
    </div>
  </div>
  <div class="report-footer">
    <span>ICM <strong>{{ICM_ID}}</strong> · Teams DRI Automated Report</span>
    <span>Generated by Claude DRI Agent · {{REPORT_TIMESTAMP}}</span>
  </div>
</div>
<script>
function copyId(btn,text){navigator.clipboard.writeText(text).then(()=>{btn.textContent='Copied!';btn.classList.add('copied');setTimeout(()=>{btn.textContent='Copy';btn.classList.remove('copied')},2000)});}
(function(){const m={'1':'#c50f1f','2':'#da3b01','3':'#c19c00','4':'#0078d4'};const s=document.querySelector('[class*="badge-sev"]');if(s){const n=s.className.match(/badge-sev(\\d)/)?.[1];if(n)document.querySelector('.hero').style.setProperty('--sev-color',m[n]);}})();
</script>
</body>
</html>
\`\`\`

After writing the HTML file, output a **markdown summary** for the DRI tab with:

## 🔍 Summary
[2-3 sentences about the incident and root cause]

## 📊 Key Identifiers
- **Tenant:** ...
- **Bot/App:** ...
- **Thread:** ...

## ⚡ Immediate Next Steps
1. [action — owner]
2. [action — owner]

## 📊 Geneva Logs
- [LogMessage URL]
- [IncomingRequest URL]
- [OutgoingRequest URL]

## 🔗 Report
Full HTML report written to: \`/tmp/icm_${icmId}_report.html\``,
      },
    ],
  };
}
