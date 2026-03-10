import { v4 as uuidv4 } from 'uuid';
import * as os from 'os';
import { Plan } from './models';

/** Pull a bare ICM number out of free text, e.g. "ICM 750444345" or "750444345" */
function extractIcmId(text: string): string | null {
  const m = text.match(/\b(\d{7,12})\b/);
  return m ? m[1] : null;
}

const DEFAULT_PORT = 3333;

export function createDriPlan(incidentDescription: string, port = DEFAULT_PORT): Plan {
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

## Step 1: Load Incident Details from IcmDataWarehouse (via local ADX proxy)

Fetch incident details by calling the local orchestrator ADX endpoint — **no browser, no Playwright needed**:

\`\`\`
GET http://localhost:${port}/api/adx/icm/${icmId}
\`\`\`

This executes the Kusto query:
\`\`\`kusto
IncidentDescriptions
| where IncidentId == ${icmId}
\`\`\`
against \`icmcluster.kusto.windows.net / IcmDataWarehouse\` using a bearer token automatically sourced from ambient-mcp.

Use the **Bash** tool to call the endpoint:
\`\`\`bash
curl -s http://localhost:${port}/api/adx/icm/${icmId}
\`\`\`

Parse the \`rows\` array from the JSON response — each row is a flat object with column names as keys.

Key columns to extract from \`IncidentDescriptions\`:
- \`IncidentId\`, \`Title\`, \`Severity\`, \`Status\`
- \`OwningTeamName\`, \`OwningContactAlias\`
- \`CreateDate\`, \`MitigateDate\`, \`ResolveDate\`
- \`Summary\`, \`AISummary\`, \`AuthoredSummary\`, \`Keywords\`
- \`Forest\`, \`DataResidencyRegion\`
- \`TroubleshootingNotes\`, \`RoutingId\`

If the endpoint returns an error (e.g. token expired), note the error and derive context from:
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

Output a concise summary: incident title, severity, status, owning team, and key points from the AI summary.`,
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

      // ── STEP 11 ── Compile Report JSON (deps: steps 3–6) ───────────────────
      {
        id:           'step_7_generate_report',
        title:        'Compile Investigation Report',
        description:  'Synthesise all findings into a structured report JSON',
        complexity:   'medium',
        dependencies: ['step_3_kusto_acl_shard', 'step_4_ado_tsgs', 'step_5_geneva_logs', 'step_6_related_icms_docs'],
        parallel_ok:  false,
        allowAllTools: true,
        status:  'pending',
        output:  '',
        workspace: sharedWs,
        claude_prompt: `You are a DRI engineer completing a full incident investigation for Microsoft Teams / Copilot Studio.

${ctx}

## Step 11: Compile Investigation Report

Read ALL findings files from **${sharedWs}/**:
- \`step1_incident.json\`   — ICM details & discussion timeline
- \`step2_identifiers.json\` — TenantId, BotId, ThreadId, CorrelationId, etc.
- \`step3_kusto.json\`       — Tenant telemetry, ACL issues, shard health
- \`step4_ado.json\`         — ADO bugs, deployments, hotfixes, TSGs
- \`step5_geneva.json\`      — Geneva log URLs
- \`step6_related.json\`     — Related ICMs and documentation

Synthesise all findings and write a compact JSON report to **${sharedWs}/report.json**:

\`\`\`json
{
  "icmId": "${icmId}",
  "title": "",
  "severity": 0,
  "status": "",
  "owningTeam": "",
  "createDate": "",
  "mitigateDate": "",
  "resolveDate": "",
  "forest": "",
  "region": "",
  "rootCauseHypothesis": "",
  "identifiers": { "tenantId":"", "userId":"", "botId":"", "threadId":"", "correlationId":"" },
  "errorSignal": "",
  "evidenceChecklist": [
    { "label": "", "status": "pass|fail|warn|unknown", "detail": "" }
  ],
  "nextSteps": [
    { "title": "", "detail": "", "priority": "critical|high|medium" }
  ],
  "genevaUrls": { "logMessage":"", "incomingRequest":"", "outgoingRequest":"" },
  "discussions": [{ "time":"", "author":"", "text":"" }],
  "relatedIcms": [{ "id":"", "title":"", "severity":0, "status":"", "createDate":"" }],
  "docs": [{ "title":"", "url":"", "description":"" }],
  "escalationPath": ["Team A → Team B → Team C"],
  "escalationNotes": "",
  "generatedAt": ""
}
\`\`\`

After writing the JSON, output a **markdown summary**:

## 🔍 Summary
[2–3 sentences about the incident and root cause hypothesis]

## 📊 Key Identifiers
- **Tenant:** ...
- **Correlation ID:** ...

## ⚡ Immediate Next Steps
1. [action — owner]
2. [action — owner]

## 📊 Geneva Logs
- LogMessage: [url]
- IncomingRequest: [url]

## 🔗 Report
Report JSON written to: \`${sharedWs}/report.json\`
View full HTML report at: http://localhost:${port}/api/dri/${icmId}/report`,
      },
    ],
  };
}
