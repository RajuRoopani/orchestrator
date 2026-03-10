---
description: "Investigate Teams ICM incidents with specialized troubleshooting workflows including extracting resource identifiers from HAR/SAZ files, finding the threadId and timestamp details, generating Geneva logs URLs for LogMessage,IncomingRequest,OutgoingRequest only, querying Kusto for tenant/forest information, extracting the errors in Geneva logs, detect the similar incidents also PRs/Bugs/TSG(Trouble shooting Guides) in ADO, read Microsoft public documentations for fixes, providing next steps for mitigation and resolution."
---

Investigate ICM incident **$ARGUMENTS** using the following structured workflow. Work through each step systematically and report findings as you go.

## Step 1: Load Incident Details from ICM

Query the IcmDataWarehouse Kusto cluster for the incident:

```kusto
// Use kusto-icm MCP server (IcmDataWarehouse)
Incidents
| where IncidentId == $ARGUMENTS
| project IncidentId, Title, Severity, Status, OwningTeamName, CreateDate, MitigateDate, ResolveDate, Summary, Keywords, RoutingId
```

Also pull the incident timeline:
```kusto
Incidents
| where IncidentId == $ARGUMENTS
| join kind=inner IncidentDiscussions on IncidentId
| project CreateDate, Author, Text
| order by CreateDate asc
```

## Step 2: Extract Resource Identifiers

From the incident title, summary, keywords, or any attached HAR/SAZ files provided by the user, extract:

- **TenantId** (GUID format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`)
- **UserId / OID** (GUID)
- **MeetingId / ThreadId**
- **ResourceId** or **ServiceTreeId**
- **CorrelationId** / **RequestId**
- **ClientSessionId**

If the user provides a HAR or SAZ file path, read it and search for:
- `x-ms-correlation-id` headers
- `x-ms-client-request-id` headers
- URL path segments containing GUIDs
- `tenantId=`, `userId=`, `oid=` query params


Note the **Forest** and **Region** — these will be referenced in later steps.

## Step 3: Kusto Investigation — Tenant & Forest Info

```kusto
// Use kusto-spoons MCP server (o365monitoring)
// Check for tenant anomalies around incident time
TeamsAnalyticsEvents
| where TenantId == "<TenantId>"
| where Timestamp between (datetime(<incident_time> - 1h) .. datetime(<incident_time> + 1h))
| summarize EventCount=count(), ErrorCount=countif(IsError==true) by bin(Timestamp, 5m), EventType
| order by Timestamp asc
```

```kusto
// Check forest/region routing
ForestInfo
| where TenantId == "<TenantId>"
| project TenantId, Forest, DataResidencyRegion, IsGovernment, CloudType
```

## Step 4: Check ACL Permissions

If the incident involves permission/access issues:

```kusto
// Use kusto-icm MCP server (IcmDataWarehouse)
// Look for related ACL incidents
Incidents
| where Keywords contains "<TenantId>"
| where Title contains "ACL" or Title contains "permission" or Title contains "access"
| where CreateDate > ago(7d)
| project IncidentId, Title, Status, CreateDate, Severity
```

Then check via ADO if there are any related ECS config changes:
- Use `ado-ecs` MCP server to check recent approvals related to the affected service
- Look for ECS experiments targeting the affected tenant or forest


## Step 5: Hot Shard Analysis

```kusto
// Use kusto-spoons MCP server (o365monitoring)
// Check if tenant is on a hot shard
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
```

Also check for shard rebalancing events:
```kusto
ShardRebalanceEvents
| where Timestamp > ago(24h)
| where TenantId == "<TenantId>"
| project Timestamp, EventType, SourceShard, DestinationShard, Status
```

## Step 6: Check for Related ADO Work Items

Use the `ado` MCP server to search O365Exchange for:
- Active bugs with the tenant ID or incident ID
- Recent deployments to the affected service/forest around the incident time
- Any known hotfixes in flight

## Step 7: Check for Related TroubleShootingGuide (TSG) or .md files

Use the `ado` MCP server to search O365Exchange for:
- Trouble shooting guides in repos or readme files related to incident which can help
- Recent deployments to the affected service/forest around the incident time
- Any known hotfixes and suggested action items

## Step 8:  Generate Geneva logs
- Geneva logs URLs for LogMessage,IncomingRequest,OutgoingRequest using below template with exact timestamp from the incident logs(specific to SMBA/APX service)

https://portal.microsoftgeneva.com/logs/dgrep?be=DGrep&time=`{TIMESTAMP}`&offset=~5&offsetUnit=Minutes&UTC=true&ep=Diagnostics%20PROD&ns=SkypeSMB&en=IncomingRequest,LogMessage,OutgoingRequest&conditions=[["AnyField","contains","{CONV_ID}"]]&clientQuery=orderby%20PreciseTimeStamp%20asc%0Awhere%20it.any("LogErrorResponseBodyHandler")%20or%20Level%20%3D%3D%202&chartEditorVisible=true&chartType=line&chartLayers=[["New%20Layer",""]]%20

## Step 9: Check for related ICMs and related hot fixes or mitigation plans
Query the IcmDataWarehouse Kusto cluster for the related incidents and their mitigation plans/fixes:

```kusto
// Use kusto-icm MCP server (IcmDataWarehouse)
Incidents
| where IncidentId == $ARGUMENTS
| project IncidentId, Title, Severity, Status, OwningTeamName, CreateDate, MitigateDate, ResolveDate, Summary, Keywords, RoutingId
```

## Step 10: Check for Related Microsoft public documentation resources

Use the `websearch` or `playwright` MCP server to search O365Exchange for:
- Learning resources or public documentation of related resources
- https://learn.microsoft.com


## Step 11: Generate Rich HTML Report

After completing all investigation steps above, generate a **complete self-contained HTML file** at `/tmp/icm_$ARGUMENTS_report.html` using the Write tool, then open it in the browser with `mcp__playwright__browser_navigate` to `file:///tmp/icm_$ARGUMENTS_report.html`.

Replace every `{{PLACEHOLDER}}` with actual findings. Use `N/A` or `Unknown` if data is unavailable — never leave a placeholder unfilled.

The HTML file to write:

```html
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

  <!-- HERO -->
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

  <!-- ROOT CAUSE -->
  <div class="rca-box">
    <h3>⚡ Root Cause Hypothesis</h3>
    <p>{{ROOT_CAUSE_HYPOTHESIS}}</p>
  </div>

  <!-- ROW 1: Identifiers + Stack Trace -->
  <div class="grid-2 mb-20">
    <div class="card">
      <div class="card-header">
        <div class="card-header-icon" style="background:#e8f4fd">🔑</div>
        <h2>Resource Identifiers</h2>
      </div>
      <div class="card-body">
        <table class="id-table">{{IDENTIFIER_ROWS}}</table>
      </div>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-header-icon" style="background:#fde7e9">🔥</div>
        <h2>Error Stack Trace</h2>
        <span class="badge badge-active" style="font-size:10px">{{ERROR_CODE}}</span>
      </div>
      <div class="card-body" style="padding:0">
        <div class="stack-trace">{{STACK_TRACE_HTML}}</div>
      </div>
    </div>
  </div>

  <!-- ROW 2: Evidence + Actions -->
  <div class="grid-2 mb-20">
    <div class="card">
      <div class="card-header">
        <div class="card-header-icon" style="background:#e6f4ea">✅</div>
        <h2>Evidence Checklist</h2>
      </div>
      <div class="card-body">
        <ul class="evidence-list">{{EVIDENCE_ITEMS}}</ul>
      </div>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-header-icon" style="background:#fff8e1">⚡</div>
        <h2>Recommended Next Steps</h2>
      </div>
      <div class="card-body">
        <ol class="action-list">{{ACTION_ITEMS}}</ol>
      </div>
    </div>
  </div>

  <!-- ROW 3: Geneva + Timeline -->
  <div class="grid-2 mb-20">
    <div class="card">
      <div class="card-header">
        <div class="card-header-icon" style="background:#1e1e1e;color:#60cdff">📊</div>
        <h2>Geneva Log Links</h2>
      </div>
      <div class="card-body">
        <div class="geneva-grid">{{GENEVA_SECTIONS}}</div>
      </div>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-header-icon" style="background:#f0f2f5">💬</div>
        <h2>Discussion Timeline</h2>
        <span class="card-count">{{DISCUSSION_COUNT}} entries</span>
      </div>
      <div class="card-body" style="max-height:380px;overflow-y:auto">
        <div class="timeline">{{TIMELINE_ITEMS}}</div>
      </div>
    </div>
  </div>

  <!-- ROW 4: Related ICMs + Docs -->
  <div class="grid-2 mb-20">
    <div class="card">
      <div class="card-header">
        <div class="card-header-icon" style="background:#fde7e9">🔗</div>
        <h2>Related ICMs</h2>
        <span class="card-count">{{RELATED_ICM_COUNT}}</span>
      </div>
      <div class="card-body" style="padding:0">
        <table class="icm-table">
          <thead><tr><th>ICM</th><th>Title</th><th>Sev</th><th>Status</th><th>Date</th></tr></thead>
          <tbody>{{RELATED_ICM_ROWS}}</tbody>
        </table>
      </div>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-header-icon" style="background:#e8f4fd">📚</div>
        <h2>Documentation &amp; TSG</h2>
      </div>
      <div class="card-body">
        <ul class="doc-list">{{DOC_ITEMS}}</ul>
      </div>
    </div>
  </div>

  <!-- ESCALATION -->
  <div class="card mb-20">
    <div class="card-header">
      <div class="card-header-icon" style="background:#fde7e9">🚨</div>
      <h2>Escalation Path</h2>
    </div>
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
function copyId(btn,text){
  navigator.clipboard.writeText(text).then(()=>{
    btn.textContent='Copied!';btn.classList.add('copied');
    setTimeout(()=>{btn.textContent='Copy';btn.classList.remove('copied')},2000);
  });
}
(function(){
  const m={'1':'#c50f1f','2':'#da3b01','3':'#c19c00','4':'#0078d4'};
  const s=document.querySelector('[class*="badge-sev"]');
  if(s){const n=s.className.match(/badge-sev(\d)/)?.[1];if(n)document.querySelector('.hero').style.setProperty('--sev-color',m[n]);}
})();
</script>
</body>
</html>
```

---

### Placeholder Reference

Fill every placeholder using the data gathered in Steps 1–10:

| Placeholder | Fill with |
|---|---|
| `{{ICM_ID}}` | The incident number (e.g. `750444345`) |
| `{{INCIDENT_TITLE}}` | Full incident title from Kusto/portal |
| `{{SEVERITY}}` | `1`, `2`, `3`, or `4` |
| `{{SEV_COLOR}}` | `#c50f1f` for Sev1/2, `#c19c00` for Sev3, `#0078d4` for Sev4 |
| `{{STATUS}}` | `ACTIVE`, `MITIGATED`, or `RESOLVED` |
| `{{STATUS_CLASS}}` | `active`, `mitigated`, or `resolved` |
| `{{STATUS_PULSE}}` | `badge-pulse` if ACTIVE, empty otherwise |
| `{{EXTRA_TAGS}}` | Each ICM tag as `<span class="badge badge-info">TAG</span>` |
| `{{CREATE_DATE}}` | Human-readable create date |
| `{{INCIDENT_AGE}}` | Days/hours since creation |
| `{{MITIGATE_DATE}}` | Mitigation date or `—` |
| `{{RESOLVE_DATE}}` | Resolution date or `—` |
| `{{OWNING_TEAM}}` | Current owning team name |
| `{{FOREST}}` / `{{REGION}}` | Forest and region from Kusto, or `Unknown` |
| `{{LAST_UPDATED}}` | Most recent discussion timestamp |
| `{{REPORT_TIMESTAMP}}` | Current datetime |
| `{{ROOT_CAUSE_HYPOTHESIS}}` | 2–4 sentence root cause summary |
| `{{ERROR_CODE}}` | Primary error code (e.g. `ConversationBlockedByUser`) |
| `{{IDENTIFIER_ROWS}}` | One `<tr>` per identifier: TenantId, BotId, EnvironmentId, ConversationId, CorrelationId, ThreadId, etc. Use copy button pattern: `<tr><td class="id-label">TENANT ID</td><td><div class="id-value-wrap"><span class="id-value">GUID</span><button class="copy-btn" onclick="copyId(this,'GUID')">Copy</button></div></td></tr>` |
| `{{STACK_TRACE_HTML}}` | Each log line as a `<span>` with class: `line-error` for errors, `line-at` for stack frames, `line-highlight` for the root-cause frame, `line-info` for info, `line-normal` for other |
| `{{EVIDENCE_ITEMS}}` | Each finding as `<li class="evidence-item"><div class="evidence-icon pass/fail/warn/unknown">✓/✗/!/？</div><div><div class="evidence-label">Label</div><div class="evidence-detail">Detail</div></div></li>` |
| `{{ACTION_ITEMS}}` | Each action as `<li class="action-item"><div class="action-num pri-critical/pri-high/pri-medium">N</div><div><div class="action-title">Title <span class="action-tag tag-oncall/tag-customer/tag-platform">TAG</span></div><div class="action-body">Detail with <code>snippets</code></div></div></li>` |
| `{{GENEVA_SECTIONS}}` | One `<div class="geneva-card">` per reporter with label, timestamp badge, and 3 links (log/incoming/outgoing) using the Geneva URL template from Step 8 |
| `{{DISCUSSION_COUNT}}` | Total number of discussion entries |
| `{{TIMELINE_ITEMS}}` | One `<div class="timeline-item">` per discussion entry. Use `timeline-dot error` for error reports, `timeline-dot success` for workarounds, `timeline-dot warn` for warnings, default for informational. Include `timeline-time`, `timeline-author`, `timeline-content` |
| `{{RELATED_ICM_COUNT}}` | Number of related ICMs found |
| `{{RELATED_ICM_ROWS}}` | One `<tr>` per related ICM with clickable ICM link, title, severity badge, status badge, date |
| `{{DOC_ITEMS}}` | Each doc as `<li class="doc-item"><div>📄</div><div><a class="doc-link" href="URL" target="_blank">Title</a><div class="doc-desc">Description</div></div></li>` |
| `{{ESCALATION_STEPS}}` | Chain of teams: `<div class="escalation-team">TEAM<small>Role</small></div><div class="escalation-arrow">→</div>...` |
| `{{ESCALATION_NOTES}}` | 1–2 sentences on when/how to escalate |

