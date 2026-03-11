/* ─── Resilient JSON parser ──────────────────────────────────────────────────── */
async function safeJson(response) {
  const text = await response.text();
  if (!text.trim()) throw new Error(`Server returned empty response (HTTP ${response.status}) — is the server running?`);
  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`Server returned non-JSON (HTTP ${response.status}): ${text.slice(0, 120)}`);
  }
}

/* ─── State ─────────────────────────────────────────────────────────────────── */
let ws = null;
// Map of planId → plan object
const plans = new Map();
// Map of planId → { startTime, timerHandle }
const executionTimers = new Map();
let streamingBubble = null;
let streamingText = '';
let planningTimer = null;
let planningStartTime = null;

/* ─── WebSocket ─────────────────────────────────────────────────────────────── */
function connect() {
  ws = new WebSocket(`ws://${location.host}`);
  ws.onopen = () => console.log('[WS] Connected');
  ws.onclose = () => { setConnectionStatus(false); setTimeout(connect, 2000); };
  ws.onerror = (e) => console.error('[WS] Error', e);
  ws.onmessage = (e) => {
    try { handleMessage(JSON.parse(e.data)); }
    catch (err) { console.error('[WS] Parse error', err); }
  };
}

function handleMessage(msg) {
  switch (msg.type) {
    case 'connected':       setConnectionStatus(true); break;
    case 'plan_token':      appendPlanToken(msg.token); break;
    case 'plan_ready':
      if (msg.plan.type === 'dri') onDriPlanReady(msg.plan);
      else finalizePlanStream(msg.plan);
      break;
    case 'task_started':    onTaskStarted(msg.planId, msg.taskId); break;
    case 'task_output':     appendTaskOutput(msg.planId, msg.taskId, msg.chunk, msg.eventType); break;
    case 'task_completed':  onTaskCompleted(msg.planId, msg.taskId, msg.success); break;
    case 'execution_done':  onExecutionDone(msg.stats); break;
    case 'dri_summary':       onDriSummary(msg.planId, msg.content, msg.icmId); break;
    case 'activity_summary':       renderActivitySummary(msg.summary, true); setActivityBtnReady(); break;
    case 'activity_error':         showToast('❌ Activity generation failed: ' + msg.message, 'error'); setActivityBtnReady(); break;
    case 'ambient_token_refreshed': showToast('⚡ ICM token refreshed from ambient-mcp', 'info'); break;
    case 'error':
      showToast('⚠ ' + msg.message, 'error');
      appendChatMessage('assistant', `⚠ Error: ${msg.message}`);
      setUIBusy(false);
      break;
    default: console.log('[WS] Unknown:', msg.type);
  }
}

/* ─── Connection Status ─────────────────────────────────────────────────────── */
function setConnectionStatus(connected) {
  document.querySelector('.status-dot').className = `status-dot ${connected ? 'connected' : 'disconnected'}`;
  document.querySelector('.status-text').textContent = connected ? 'Connected' : 'Disconnected';
}

/* ─── Chat ──────────────────────────────────────────────────────────────────── */
function sendMessage() {
  const input = document.getElementById('chatInput');
  const text = input.value.trim();
  if (!text) return;
  input.value = '';

  // Route /DRI commands to the DRI investigation flow
  if (text.toLowerCase().startsWith('/dri')) {
    const incident = text.slice(4).trim();
    sendDriCommand(incident || text);
    return;
  }

  appendChatMessage('user', text);
  setUIBusy(true);

  streamingText = '';
  planningStartTime = Date.now();
  streamingBubble = appendChatMessage('assistant', '🧠 Thinking… (0s)', true);
  planningTimer = setInterval(() => {
    if (streamingBubble) {
      const s = Math.round((Date.now() - planningStartTime) / 1000);
      streamingBubble.textContent = `🧠 Generating plan… (${s}s)`;
    }
  }, 1000);

  fetch('/api/plan', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ message: text }),
  })
    .then((r) => safeJson(r))
    .then((d) => { if (!d.success) throw new Error(d.error || 'Unknown error'); })
    .catch((err) => {
      clearPlanningTimer();
      showToast('Failed: ' + err.message, 'error');
      if (streamingBubble) { streamingBubble.classList.remove('streaming'); streamingBubble.textContent = '❌ ' + err.message; streamingBubble = null; }
      setUIBusy(false);
    });
}

function appendChatMessage(role, text, streaming = false) {
  const container = document.getElementById('chatMessages');
  const wrapper = document.createElement('div');
  wrapper.className = `message ${role}`;
  const bubble = document.createElement('div');
  bubble.className = `message-bubble${streaming ? ' streaming' : ''}`;
  bubble.textContent = text;
  wrapper.appendChild(bubble);
  container.appendChild(wrapper);
  container.scrollTop = container.scrollHeight;
  if (!streaming) saveChatHistory();
  return bubble;
}

/* ─── Chat History (server-backed, localStorage fallback) ────────────────────── */
const CHAT_STORAGE_KEY = 'orchestrator_chat_history';
let _chatSaveTimer = null;

function collectChatMessages() {
  return [...document.querySelectorAll('#chatMessages .message')].map((el) => ({
    role: el.classList.contains('user') ? 'user' : 'assistant',
    text: el.querySelector('.message-bubble')?.textContent ?? '',
  }));
}

function saveChatHistory() {
  const messages = collectChatMessages();
  // localStorage fallback (instant)
  try { localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify(messages.slice(-200))); } catch {}
  // Debounce server save — wait 800ms after last message
  clearTimeout(_chatSaveTimer);
  _chatSaveTimer = setTimeout(() => {
    fetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages: messages.slice(-200) }),
    }).catch(() => {});
  }, 800);
}

function renderChatMessages(messages) {
  if (!messages.length) return;
  const container = document.getElementById('chatMessages');
  container.innerHTML = '';
  for (const { role, text } of messages) {
    const wrapper = document.createElement('div');
    wrapper.className = `message ${role}`;
    const bubble = document.createElement('div');
    bubble.className = 'message-bubble';
    bubble.textContent = text;
    wrapper.appendChild(bubble);
    container.appendChild(wrapper);
  }
  container.scrollTop = container.scrollHeight;
}

function restoreChatHistory() {
  // Try server first, fall back to localStorage
  fetch('/api/chat')
    .then((r) => safeJson(r))
    .then((messages) => {
      if (Array.isArray(messages) && messages.length) {
        renderChatMessages(messages);
      } else {
        // Fallback to localStorage
        const saved = JSON.parse(localStorage.getItem(CHAT_STORAGE_KEY) ?? '[]');
        renderChatMessages(saved);
      }
    })
    .catch(() => {
      try {
        const saved = JSON.parse(localStorage.getItem(CHAT_STORAGE_KEY) ?? '[]');
        renderChatMessages(saved);
      } catch {}
    });
}

function appendPlanToken(token) {
  streamingText += token;
  clearPlanningTimer();
  if (streamingBubble) {
    const s = Math.round((Date.now() - (planningStartTime || Date.now())) / 1000);
    streamingBubble.textContent = `📋 Processing… (${s}s, ${streamingText.length} chars)`;
  }
}

function clearPlanningTimer() {
  if (planningTimer) { clearInterval(planningTimer); planningTimer = null; }
}

function finalizePlanStream(plan) {
  clearPlanningTimer();
  plans.set(plan.id, plan);

  if (streamingBubble) {
    streamingBubble.classList.remove('streaming');
    streamingBubble.textContent = `✅ Plan ready: "${plan.title}" (${plan.tasks.length} tasks)`;
    streamingBubble = null;
  }

  renderPlanPanel(plan);
  switchTab('plan');
  setUIBusy(false);
  showToast('📋 Plan ready — approve to execute', 'info');
}

/* ─── Plan Tab ──────────────────────────────────────────────────────────────── */
function renderPlanPanel(plan) {
  document.getElementById('planEmpty').style.display = 'none';

  // Create or replace plan card
  let card = document.getElementById(`plan-card-${plan.id}`);
  if (!card) {
    card = document.createElement('div');
    card.id = `plan-card-${plan.id}`;
    card.className = 'plan-container';
    document.getElementById('planContainer').appendChild(card);
  }
  document.getElementById('planContainer').style.display = 'flex';

  card.innerHTML = '';

  // Header
  const header = document.createElement('div');
  header.className = 'plan-header';
  header.innerHTML = `
    <div class="plan-info">
      <h2>${escHtml(plan.title)}</h2>
      <p class="plan-description">${escHtml(plan.description)}</p>
      <div class="plan-meta">
        <span class="meta-badge"><span class="meta-icon">⏱</span>${escHtml(plan.estimated_duration)}</span>
        <span class="meta-badge"><span class="meta-icon">📌</span>${plan.tasks.length} tasks</span>
      </div>
    </div>
    <button class="execute-btn" id="exec-btn-${plan.id}" onclick="executePlan('${plan.id}')">✅ Approve &amp; Execute</button>
  `;
  card.appendChild(header);

  // Task tree
  const tree = document.createElement('div');
  tree.className = 'task-tree';
  for (const task of plan.tasks) tree.appendChild(renderTaskTreeItem(plan.id, task));
  card.appendChild(tree);
}

function renderTaskTreeItem(planId, task) {
  const item = document.createElement('div');
  item.className = 'task-item';
  item.id = `tree-${planId}-${task.id}`;

  let badges = `<span class="complexity-badge ${task.complexity}">${task.complexity}</span>`;
  if (task.parallel_ok) badges += `<span class="parallel-badge">⚡ parallel</span>`;

  item.innerHTML = `
    <div class="task-item-header">
      <span class="task-status-icon" id="tree-icon-${planId}-${task.id}">${getStatusIcon('pending')}</span>
      <span class="task-title">${escHtml(task.title)}</span>
      ${badges}
    </div>
    ${task.description ? `<div class="task-item-desc">${escHtml(task.description)}</div>` : ''}
    ${task.dependencies?.length ? `<div class="task-deps">Depends on: ${task.dependencies.map(d => `<span>${d}</span>`).join(', ')}</div>` : ''}
  `;
  return item;
}

function getStatusIcon(status) {
  return { pending: '⏳', running: '▶️', completed: '✅', failed: '❌' }[status] || '⏳';
}

function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

/* ─── Execute ───────────────────────────────────────────────────────────────── */
function executePlan(planId) {
  const plan = plans.get(planId);
  if (!plan) return;

  document.getElementById(`exec-btn-${planId}`).disabled = true;

  renderDashboardPanel(plan);
  switchTab('dashboard');

  const startTime = Date.now();
  const timerHandle = setInterval(() => {
    const el = document.getElementById(`stat-time-${planId}`);
    if (el) el.textContent = `${Math.round((Date.now() - startTime) / 1000)}s`;
  }, 1000);
  executionTimers.set(planId, { startTime, timerHandle });

  fetch('/api/execute', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ planId }),
  })
    .then((r) => safeJson(r))
    .then((d) => {
      if (!d.success) throw new Error(d.error);
      appendChatMessage('assistant', `⚡ Executing: "${plan.title}"…`);
    })
    .catch((err) => {
      stopPlanTimer(planId);
      showToast('Execution failed: ' + err.message, 'error');
    });
}

/* ─── Dashboard Tab ─────────────────────────────────────────────────────────── */
function renderDashboardPanel(plan) {
  document.getElementById('dashboardEmpty').style.display = 'none';

  let panel = document.getElementById(`dash-${plan.id}`);
  if (!panel) {
    panel = document.createElement('div');
    panel.id = `dash-${plan.id}`;
    panel.className = 'dashboard-panel';
    document.getElementById('dashboardContainer').appendChild(panel);
  }
  document.getElementById('dashboardContainer').style.display = 'flex';

  panel.innerHTML = `
    <div class="dash-plan-title">
      <span>${escHtml(plan.title)}</span>
      <button class="cancel-plan-btn" onclick="cancelPlan('${plan.id}')">✕ Cancel</button>
    </div>
    <div class="stats-bar">
      <div class="stat-item"><span class="stat-value" id="stat-total-${plan.id}">${plan.tasks.length}</span><span class="stat-label">Total</span></div>
      <div class="stat-item"><span class="stat-value success" id="stat-done-${plan.id}">0</span><span class="stat-label">Done</span></div>
      <div class="stat-item"><span class="stat-value running" id="stat-run-${plan.id}">0</span><span class="stat-label">Running</span></div>
      <div class="stat-item"><span class="stat-value error" id="stat-fail-${plan.id}">0</span><span class="stat-label">Failed</span></div>
      <div class="stat-item"><span class="stat-value" id="stat-time-${plan.id}">0s</span><span class="stat-label">Elapsed</span></div>
    </div>
    <div class="task-cards" id="cards-${plan.id}"></div>
  `;

  const cardsEl = document.getElementById(`cards-${plan.id}`);
  for (const task of plan.tasks) cardsEl.appendChild(renderTaskCard(plan.id, task));
}

function renderTaskCard(planId, task) {
  const card = document.createElement('div');
  card.className = 'task-card status-pending';
  card.id = `card-${planId}-${task.id}`;

  card.innerHTML = `
    <div class="task-card-header" onclick="toggleOutput('${planId}','${task.id}')">
      <span class="task-card-icon" id="icon-${planId}-${task.id}">⏳</span>
      <span class="task-card-title">${escHtml(task.title)}</span>
      <span class="task-card-time" id="time-${planId}-${task.id}"></span>
      <button class="expand-btn" id="expand-${planId}-${task.id}">▼</button>
    </div>
    <div class="task-output" id="out-${planId}-${task.id}">
      <div class="terminal" id="term-${planId}-${task.id}">(waiting for output…)</div>
    </div>
  `;
  return card;
}

function toggleOutput(planId, taskId) {
  const area = document.getElementById(`out-${planId}-${taskId}`);
  const btn  = document.getElementById(`expand-${planId}-${taskId}`);
  if (!area) return;
  area.classList.toggle('visible');
  btn.classList.toggle('expanded');
  if (area.classList.contains('visible')) {
    const t = document.getElementById(`term-${planId}-${taskId}`);
    if (t) t.scrollTop = t.scrollHeight;
  }
}

/* ─── Task Event Handlers ───────────────────────────────────────────────────── */
function onTaskStarted(planId, taskId) {
  setCardStatus(planId, taskId, 'running');
  updatePlanStats(planId);
  showToast(`▶ ${taskId} started`, 'info');
}

function onTaskCompleted(planId, taskId, success) {
  const status = success ? 'completed' : 'failed';
  setCardStatus(planId, taskId, status);

  const timer = executionTimers.get(planId);
  const timeEl = document.getElementById(`time-${planId}-${taskId}`);
  if (timeEl && timer) timeEl.textContent = `${((Date.now() - timer.startTime) / 1000).toFixed(1)}s`;

  if (!success) {
    const area = document.getElementById(`out-${planId}-${taskId}`);
    if (area && !area.classList.contains('visible')) toggleOutput(planId, taskId);
  }

  updatePlanStats(planId);
  showToast(success ? `✅ ${taskId} done` : `❌ ${taskId} failed`, success ? 'success' : 'error');
}

function setCardStatus(planId, taskId, status) {
  const card = document.getElementById(`card-${planId}-${taskId}`);
  if (card) card.className = `task-card status-${status}`;

  const icon = document.getElementById(`icon-${planId}-${taskId}`);
  if (icon) {
    if (status === 'running') icon.innerHTML = '<div class="running-ring"></div>';
    else icon.textContent = getStatusIcon(status);
  }

  const treeIcon = document.getElementById(`tree-icon-${planId}-${taskId}`);
  if (treeIcon) treeIcon.textContent = getStatusIcon(status);
}

function appendTaskOutput(planId, taskId, chunk, eventType) {
  const term = document.getElementById(`term-${planId}-${taskId}`);
  if (!term) return;
  if (term.textContent === '(waiting for output…)') term.textContent = '';

  const span = document.createElement('span');
  if (eventType === 'tool_use' || chunk.startsWith('[Tool:')) span.className = 'tool-call';
  else if (eventType === 'tool_result' || chunk.startsWith('[Result]')) span.className = 'tool-result';
  else if (eventType === 'error') span.className = 'err-text';
  span.textContent = chunk;
  term.appendChild(span);

  const area = document.getElementById(`out-${planId}-${taskId}`);
  if (area?.classList.contains('visible')) term.scrollTop = term.scrollHeight;
}

function updatePlanStats(planId) {
  // Collect cards from dashboard panel or DRI panel (same ID pattern for cards)
  let cards = document.querySelectorAll(`#cards-${planId} .task-card`);
  if (cards.length === 0) cards = document.querySelectorAll(`#dri-cards-${planId} .task-card`);

  let done = 0, running = 0, failed = 0;
  for (const c of cards) {
    if (c.classList.contains('status-completed')) done++;
    else if (c.classList.contains('status-running')) running++;
    else if (c.classList.contains('status-failed')) failed++;
  }
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
  // Dashboard stats
  set(`stat-done-${planId}`, done);
  set(`stat-run-${planId}`, running);
  set(`stat-fail-${planId}`, failed);
  // DRI stats
  set(`dri-done-${planId}`, done);
  set(`dri-run-${planId}`, running);
  set(`dri-fail-${planId}`, failed);
}

function onExecutionDone(stats) {
  stopPlanTimer(stats.planId);
  const elapsed = (stats.durationMs / 1000).toFixed(1);
  // Update timer in both dashboard and DRI panel
  const timerEl = document.getElementById(`stat-time-${stats.planId}`)
    || document.getElementById(`dri-time-${stats.planId}`);
  if (timerEl) timerEl.textContent = `${elapsed}s`;

  const plan = plans.get(stats.planId);
  if (plan?.type === 'dri') return; // DRI summary handled by onDriSummary

  const title = plan ? plan.title : stats.planId;
  const msg = `🏁 "${title}": ${stats.completedTasks}/${stats.totalTasks} tasks in ${elapsed}s`;
  appendChatMessage('assistant', msg);
  showToast(msg, stats.failedTasks > 0 ? 'error' : 'success');
}

function stopPlanTimer(planId) {
  const t = executionTimers.get(planId);
  if (t) { clearInterval(t.timerHandle); executionTimers.delete(planId); }
}

function cancelPlan(planId) {
  fetch(`/api/plan/${planId}/cancel`, { method: 'POST' });
  stopPlanTimer(planId);
  showToast('Plan cancelled', 'info');
}

/* ─── Tabs ──────────────────────────────────────────────────────────────────── */
function switchTab(name) {
  document.querySelectorAll('.tab').forEach((t) => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach((c) => c.classList.remove('active'));
  document.querySelector(`[data-tab="${name}"]`).classList.add('active');
  document.getElementById(`tab-${name}`).classList.add('active');
  if (name === 'history') loadHistoryTab();
  if (name === 'icm') switchToIcmTab();
}

/* ─── UI Helpers ─────────────────────────────────────────────────────────────── */
function setUIBusy(busy) {
  document.getElementById('sendBtn').disabled = busy;
  document.getElementById('chatInput').disabled = busy;
}

function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  document.getElementById('toastContainer').appendChild(toast);
  setTimeout(() => toast.remove(), 3200);
}

/* ─── DRI Investigation ─────────────────────────────────────────────────────── */
function sendDriCommand(incident) {
  appendChatMessage('user', `/DRI ${incident}`);
  setUIBusy(true);

  streamingText = '';
  planningStartTime = Date.now();
  streamingBubble = appendChatMessage('assistant', '🚨 Preparing DRI investigation…', true);
  planningTimer = setInterval(() => {
    if (streamingBubble) {
      const s = Math.round((Date.now() - planningStartTime) / 1000);
      streamingBubble.textContent = `🚨 Setting up investigation steps… (${s}s)`;
    }
  }, 1000);

  fetch('/api/dri', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ message: incident }),
  })
    .then((r) => safeJson(r))
    .then((d) => { if (!d.success) throw new Error(d.error || 'Unknown error'); })
    .catch((err) => {
      clearPlanningTimer();
      showToast('DRI failed: ' + err.message, 'error');
      if (streamingBubble) { streamingBubble.classList.remove('streaming'); streamingBubble.textContent = '❌ ' + err.message; streamingBubble = null; }
      setUIBusy(false);
    });
}

function onDriPlanReady(plan) {
  clearPlanningTimer();
  plans.set(plan.id, plan);

  if (streamingBubble) {
    streamingBubble.classList.remove('streaming');
    streamingBubble.textContent = `🚨 DRI investigation running — ${plan.tasks.length} steps`;
    streamingBubble = null;
  }

  renderDriPanel(plan);
  switchTab('dri');
  setUIBusy(false);
  showToast('🚨 DRI investigation started', 'info');

  // Start elapsed timer
  const startTime = Date.now();
  const timerHandle = setInterval(() => {
    const el = document.getElementById(`dri-time-${plan.id}`);
    if (el) el.textContent = `${Math.round((Date.now() - startTime) / 1000)}s`;
  }, 1000);
  executionTimers.set(plan.id, { startTime, timerHandle });
}

function renderDriPanel(plan) {
  document.getElementById('driEmpty').style.display = 'none';

  let panel = document.getElementById(`dri-panel-${plan.id}`);
  if (!panel) {
    panel = document.createElement('div');
    panel.id = `dri-panel-${plan.id}`;
    panel.className = 'dri-panel';
    document.getElementById('driContainer').appendChild(panel);
  }
  document.getElementById('driContainer').style.display = 'flex';

  panel.innerHTML = `
    <div class="dri-panel-header">
      <div class="dri-badge-row">
        <span class="dri-badge">🚨 DRI</span>
        <span class="dri-panel-title">${escHtml(plan.title)}</span>
      </div>
      <div class="dri-panel-desc">${escHtml(plan.description)}</div>
    </div>
    <div class="stats-bar dri-stats-bar">
      <div class="stat-item">
        <span class="stat-value" id="dri-total-${plan.id}">${plan.tasks.length}</span>
        <span class="stat-label">Steps</span>
      </div>
      <div class="stat-item">
        <span class="stat-value success" id="dri-done-${plan.id}">0</span>
        <span class="stat-label">Done</span>
      </div>
      <div class="stat-item">
        <span class="stat-value running" id="dri-run-${plan.id}">0</span>
        <span class="stat-label">Running</span>
      </div>
      <div class="stat-item">
        <span class="stat-value error" id="dri-fail-${plan.id}">0</span>
        <span class="stat-label">Failed</span>
      </div>
      <div class="stat-item">
        <span class="stat-value" id="dri-time-${plan.id}">0s</span>
        <span class="stat-label">Elapsed</span>
      </div>
    </div>
    <div class="task-cards" id="dri-cards-${plan.id}"></div>
    <div class="dri-summary-box" id="dri-summary-${plan.id}" style="display:none"></div>
  `;

  const cardsEl = document.getElementById(`dri-cards-${plan.id}`);
  for (const task of plan.tasks) cardsEl.appendChild(renderDriStepCard(plan.id, task));
}

function renderDriStepCard(planId, task) {
  const card = document.createElement('div');
  // Reuse the same card-{planId}-{taskId} ID pattern so existing task event handlers work
  card.id = `card-${planId}-${task.id}`;
  card.className = 'task-card dri-step-card status-pending';

  card.innerHTML = `
    <div class="task-card-header" onclick="toggleOutput('${planId}','${task.id}')">
      <span class="task-card-icon" id="icon-${planId}-${task.id}">⏳</span>
      <div class="dri-step-info">
        <span class="task-card-title">${escHtml(task.title)}</span>
        <span class="dri-step-desc">${escHtml(task.description)}</span>
      </div>
      <span class="task-card-time" id="time-${planId}-${task.id}"></span>
      <button class="expand-btn" id="expand-${planId}-${task.id}">▼</button>
    </div>
    <div class="task-output" id="out-${planId}-${task.id}">
      <div class="terminal" id="term-${planId}-${task.id}">(waiting for output…)</div>
    </div>
  `;
  return card;
}

function onDriSummary(planId, content, icmId) {
  stopPlanTimer(planId);

  const summaryEl = document.getElementById(`dri-summary-${planId}`);
  if (!summaryEl) return;

  summaryEl.style.display = 'block';

  if (icmId && icmId !== 'UNKNOWN') {
    summaryEl.innerHTML = `
      <div class="dri-summary-header">
        <span class="dri-summary-icon">📋</span>
        <span>Investigation Report — ICM ${escHtml(icmId)}</span>
        <a class="dri-report-link" href="/api/dri/${encodeURIComponent(icmId)}/report" target="_blank">Open in new tab ↗</a>
      </div>
      <iframe
        class="dri-report-iframe"
        src="/api/dri/${encodeURIComponent(icmId)}/report"
        title="ICM ${escHtml(icmId)} Investigation Report"
      ></iframe>
    `;
  } else {
    summaryEl.innerHTML = `
      <div class="dri-summary-header">
        <span class="dri-summary-icon">📋</span>
        <span>Investigation Report</span>
      </div>
      <div class="dri-summary-content">${markdownToHtml(content)}</div>
    `;
  }

  appendChatMessage('assistant', '🚨 DRI investigation complete. Full report available in the DRI tab.');
  showToast('🚨 DRI investigation complete', 'success');

  setTimeout(() => summaryEl.scrollIntoView({ behavior: 'smooth', block: 'start' }), 200);
}

function markdownToHtml(text) {
  return text
    // Section headers
    .replace(/^## (.+)$/gm, '<h3 class="dri-h3">$1</h3>')
    .replace(/^### (.+)$/gm, '<h4 class="dri-h4">$1</h4>')
    // Bold
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    // Inline code
    .replace(/`([^`\n]+)`/g, '<code class="dri-code">$1</code>')
    // Bullet list items
    .replace(/^[-*] (.+)$/gm, '<li>$1</li>')
    // Numbered list items
    .replace(/^\d+\. (.+)$/gm, '<li>$1</li>')
    // Wrap consecutive li items in ul
    .replace(/(<li>[\s\S]*?<\/li>)(\s*<li>[\s\S]*?<\/li>)*/g, (m) => `<ul>${m}</ul>`)
    // Blockquotes
    .replace(/^> (.+)$/gm, '<blockquote class="dri-quote">$1</blockquote>')
    // Double newlines → paragraph break
    .replace(/\n\n/g, '<br><br>')
    // Remaining newlines
    .replace(/\n/g, '<br>');
}

/* ─── ICM Dashboard Tab ─────────────────────────────────────────────────────── */
let icmAllRows = [];

const ICM_TOKEN_KEY = 'orchestrator_icm_token';
const ICM_TEAM_KEY  = 'orchestrator_icm_team';
const ICM_TEAMS_KEY = 'orchestrator_icm_teams';

const DEFAULT_TEAMS = [
  { id: 85183, label: 'Team 85183' },
  { id: 53893, label: 'Team 53893' },
];

function getStoredTeams() {
  try { return JSON.parse(localStorage.getItem(ICM_TEAMS_KEY) ?? 'null') || DEFAULT_TEAMS; }
  catch { return DEFAULT_TEAMS; }
}

function saveStoredTeams(teams) {
  localStorage.setItem(ICM_TEAMS_KEY, JSON.stringify(teams));
}

function rebuildTeamDropdown() {
  const sel = document.getElementById('icmTeamIdFilter');
  if (!sel) return;
  const current = localStorage.getItem(ICM_TEAM_KEY) ?? '';
  const teams = getStoredTeams();
  sel.innerHTML = '<option value="">All teams</option>' +
    teams.map(t => `<option value="${t.id}" ${String(t.id) === current ? 'selected' : ''}>${escHtml(t.label)} (${t.id})</option>`).join('');
}

function promptAddTeam() {
  const idStr = window.prompt('Enter Team ID (number):');
  if (!idStr?.trim()) return;
  const id = parseInt(idStr.trim(), 10);
  if (isNaN(id)) return showToast('Invalid team ID', 'error');
  const label = window.prompt('Team name / label:', `Team ${id}`) ?? `Team ${id}`;
  const teams = getStoredTeams();
  if (!teams.find(t => t.id === id)) {
    teams.push({ id, label: label.trim() || `Team ${id}` });
    saveStoredTeams(teams);
  }
  localStorage.setItem(ICM_TEAM_KEY, String(id));
  rebuildTeamDropdown();
  loadIcms(true);
}

function switchToIcmTab() {
  rebuildTeamDropdown();
  const saved = localStorage.getItem(ICM_TOKEN_KEY);
  if (!saved) {
    showIcmTokenPanel();
  } else if (!icmAllRows.length) {
    loadIcms(false);
  }
}

function toggleIcmTokenPanel() {
  const banner = document.getElementById('icmTokenBanner');
  if (banner.style.display === 'none') showIcmTokenPanel();
  else hideIcmTokenPanel();
}

function showIcmTokenPanel() {
  document.getElementById('icmTokenBanner').style.display = 'block';
  const saved = localStorage.getItem(ICM_TOKEN_KEY);
  document.getElementById('icmTokenInput').value = saved ? '(token saved — paste new one to replace)' : '';
  setTimeout(() => {
    const input = document.getElementById('icmTokenInput');
    input.focus();
    if (input.value.startsWith('(token')) input.select();
  }, 50);
}

function hideIcmTokenPanel() {
  document.getElementById('icmTokenBanner').style.display = 'none';
  document.getElementById('icmTokenInput').value = '';
}

// Keep old name as alias for any existing callers
function showIcmTokenBanner() { showIcmTokenPanel(); }

function clearIcmToken() {
  localStorage.removeItem(ICM_TOKEN_KEY);
  localStorage.removeItem(ICM_TEAM_KEY);
  icmAllRows = [];
  fetch('/api/icm/token', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ token: '' }) }).catch(() => {});
  // Reset table area
  document.getElementById('icmFilters').style.display = 'none';
  document.getElementById('icmRefreshBtn').style.display = 'none';
  document.getElementById('icmTable').style.display = 'none';
  document.getElementById('icmCount').textContent = '';
  document.getElementById('icmFetchedAt').textContent = '';
  document.getElementById('icmEmpty').style.display = 'block';
  document.getElementById('icmEmpty').textContent = 'Click ⚙ Token to set your ICM authorization token, then Refresh to load incidents.';
  document.getElementById('icmTokenInput').value = '';
  showToast('Token cleared', 'info');
}

function saveIcmToken() {
  const raw = document.getElementById('icmTokenInput').value.trim();
  if (!raw || raw.startsWith('(token')) return showToast('Paste a new Authorization token first', 'error');
  const token = raw.replace(/^Bearer\s+/i, '');
  const teamIdVal = document.getElementById('icmTeamIdFilter')?.value?.trim();
  localStorage.setItem(ICM_TOKEN_KEY, token);
  if (teamIdVal) localStorage.setItem(ICM_TEAM_KEY, teamIdVal);
  else localStorage.removeItem(ICM_TEAM_KEY);
  fetch('/api/icm/token', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ token, teamId: teamIdVal ? Number(teamIdVal) : null }) })
    .then((r) => safeJson(r))
    .then((d) => {
      hideIcmTokenPanel();
      const label = [d.alias, d.teamId ? `team ${d.teamId}` : ''].filter(Boolean).join(' · ');
      if (label) {
        document.getElementById('icmAliasLabel').textContent = label;
        document.getElementById('icmAliasLabel').style.display = 'inline';
      }
      showToast(`✓ Token saved${d.alias ? ' for ' + d.alias : ''}`, 'success');
      document.getElementById('icmFilters').style.display = 'flex';
      document.getElementById('icmRefreshBtn').style.display = 'inline-flex';
      loadIcms(true);
    })
    .catch((err) => showToast('Failed to save token: ' + err.message, 'error'));
}

function loadIcms(forceRefresh = false) {
  // Send team ID to server before fetching so it scopes the query
  const teamIdVal = document.getElementById('icmTeamIdFilter')?.value?.trim();
  const teamId = teamIdVal ? Number(teamIdVal) : null;
  const token = localStorage.getItem(ICM_TOKEN_KEY);
  if (token) {
    fetch('/api/icm/token', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token, teamId }),
    }).catch(() => {});
  }

  document.getElementById('icmEmpty').style.display = 'none';
  document.getElementById('icmTable').style.display = 'none';
  document.getElementById('icmLoading').style.display = 'flex';
  document.getElementById('icmRefreshBtn').disabled = true;
  document.getElementById('icmCount').textContent = '';

  const doFetch = () => fetch('/api/icm/active')
    .then((r) => safeJson(r))
    .then((d) => {
      document.getElementById('icmLoading').style.display = 'none';
      document.getElementById('icmRefreshBtn').disabled = false;
      if (!d.success) {
        if (d.error === 'NO_TOKEN' || d.error === 'TOKEN_EXPIRED') {
          showIcmTokenPanel();
          document.getElementById('icmEmpty').style.display = 'block';
          document.getElementById('icmEmpty').textContent = d.error === 'TOKEN_EXPIRED'
            ? '⚠ Token expired. Paste a new one using ⚙ Token above.'
            : 'Click ⚙ Token to set your ICM authorization token.';
          return;
        }
        throw new Error(d.error || 'Unknown error');
      }
      // Success — show filters + refresh button
      document.getElementById('icmFilters').style.display = 'flex';
      document.getElementById('icmRefreshBtn').style.display = 'inline-flex';
      icmAllRows = d.data ?? [];
      if (d.fetchedAt) {
        const t = new Date(d.fetchedAt);
        document.getElementById('icmFetchedAt').textContent = `Fetched ${t.toLocaleTimeString()}`;
      }
      applyIcmFilters();
    })
    .catch((err) => {
      document.getElementById('icmLoading').style.display = 'none';
      document.getElementById('icmRefreshBtn').disabled = false;
      document.getElementById('icmEmpty').style.display = 'block';
      document.getElementById('icmEmpty').textContent = '⚠ ' + err.message;
      showToast('ICM fetch failed: ' + err.message, 'error');
    });

  if (forceRefresh) {
    fetch('/api/icm/refresh', { method: 'POST' }).then(doFetch).catch(doFetch);
  } else {
    doFetch();
  }
}

function refreshIcms() { loadIcms(true); }

function onTeamIdFilterChange() {
  const val = document.getElementById('icmTeamIdFilter').value;
  if (val) localStorage.setItem(ICM_TEAM_KEY, val);
  else localStorage.removeItem(ICM_TEAM_KEY);
  loadIcms(true);
}

function applyIcmFilters() {
  const titleQ  = document.getElementById('icmFilterTitle').value.toLowerCase();
  const sevQ    = document.getElementById('icmFilterSev').value;
  const statusQ = document.getElementById('icmFilterStatus').value;

  const filtered = icmAllRows.filter((r) => {
    if (titleQ  && !r.Title?.toLowerCase().includes(titleQ))           return false;
    if (sevQ    && String(r.Severity) !== sevQ)                        return false;
    if (statusQ && r.Status !== statusQ)                               return false;
    return true;
  });

  renderIcmTable(filtered);
}

function renderIcmTable(rows) {
  const tbody = document.getElementById('icmTableBody');
  const table = document.getElementById('icmTable');
  const empty = document.getElementById('icmEmpty');
  const statsBar = document.getElementById('icmStatsBar');

  document.getElementById('icmCount').textContent = `${rows.length} incident${rows.length !== 1 ? 's' : ''}`;

  // Update stats bar from full dataset (not filtered)
  const allRows = icmAllRows.length ? icmAllRows : rows;
  const sev1 = allRows.filter(r => r.Severity === 1).length;
  const sev2 = allRows.filter(r => r.Severity === 2).length;
  const sev3 = allRows.filter(r => r.Severity === 3).length;
  const active = allRows.filter(r => r.Status === 'Active').length;
  if (statsBar) {
    document.getElementById('icmStatSev1').textContent = sev1;
    document.getElementById('icmStatSev2').textContent = sev2;
    document.getElementById('icmStatSev3').textContent = sev3;
    document.getElementById('icmStatActive').textContent = active;
    statsBar.style.display = allRows.length ? 'flex' : 'none';
  }

  if (!rows.length) {
    table.style.display = 'none';
    empty.style.display = 'block';
    empty.textContent = icmAllRows.length ? 'No incidents match your filters.' : 'No active ICMs found.';
    return;
  }

  empty.style.display = 'none';
  table.style.display = 'table';
  tbody.innerHTML = '';

  for (const r of rows) {
    const tr = document.createElement('tr');
    const created = r.CreateDate ? new Date(r.CreateDate).toLocaleDateString() : '—';
    const sevClass = `sev-badge sev-${r.Severity}`;
    const statusClass = r.Status === 'Active' ? 'icm-status-active' : r.Status === 'Mitigated' ? 'icm-status-mitigated' : 'icm-status-other';
    const flags = [
      r.IsOutage            ? '<span class="icm-flag outage">Outage</span>' : '',
      r.IsCustomerImpacting ? '<span class="icm-flag cri">CRI</span>'    : '',
    ].filter(Boolean).join('');
    tr.innerHTML = `
      <td><span class="${sevClass}">${r.Severity}</span></td>
      <td><a class="icm-id-link" href="https://portal.microsofticm.com/imp/v3/incidents/detail/${r.IncidentId}/home" target="_blank">${r.IncidentId}</a></td>
      <td class="icm-title-cell" title="${escHtml(r.Title ?? '')}">${escHtml(r.Title ?? '—')}</td>
      <td><span class="${statusClass}">${escHtml(r.Status ?? '—')}</span></td>
      <td class="icm-team-cell">${escHtml(r.OwningTeamName ?? '—')}</td>
      <td class="icm-contact-cell">${escHtml(r.ContactAlias ?? '—')}</td>
      <td>${created}</td>
      <td>${flags || '—'}</td>
      <td>
        <button class="icm-dri-btn" onclick="startDriFromIcm(${r.IncidentId})">🚨 Investigate</button>
      </td>
    `;
    tbody.appendChild(tr);
  }
}

function startDriFromIcm(icmId) {
  document.getElementById('chatInput').value = `/DRI ${icmId}`;
  switchTab('dri');
  sendMessage();
}

/* ─── History Tab ───────────────────────────────────────────────────────────── */
function loadHistoryTab() {
  const list = document.getElementById('historyList');
  list.innerHTML = '<div class="history-empty">Loading…</div>';
  fetch('/api/history')
    .then((r) => safeJson(r))
    .then((items) => renderHistoryList(items))
    .catch(() => { list.innerHTML = '<div class="history-empty">Failed to load history.</div>'; });
}

function renderHistoryList(items) {
  const list = document.getElementById('historyList');
  if (!items.length) {
    list.innerHTML = '<div class="history-empty">No DRI investigations yet.</div>';
    return;
  }
  list.innerHTML = '';
  for (const item of items) {
    const el = document.createElement('div');
    el.className = 'history-item';
    const date = new Date(item.startedAt);
    const dateStr = date.toLocaleDateString() + ' ' + date.toLocaleTimeString();
    const dur = item.durationMs ? `${Math.round(item.durationMs / 1000)}s` : '—';
    const hasReport = item.icmId && item.icmId !== 'UNKNOWN';
    el.innerHTML = `
      <div class="history-item-header">
        <span class="history-icm-badge">ICM ${escHtml(item.icmId ?? '?')}</span>
        <span class="history-item-title">${escHtml(item.title ?? 'DRI Investigation')}</span>
        <span class="history-item-time">${dateStr}</span>
      </div>
      <div class="history-item-meta">
        <span class="history-meta-pill">${item.completedTasks ?? 0}/${item.totalTasks ?? 0} steps</span>
        ${item.failedTasks ? `<span class="history-meta-pill error">${item.failedTasks} failed</span>` : ''}
        <span class="history-meta-pill">${dur}</span>
        ${hasReport ? `<a class="history-meta-pill link" href="/api/dri/${encodeURIComponent(item.icmId)}/report" target="_blank">View Report ↗</a>` : ''}
        <button class="history-rerun-btn" onclick="rerunDri('${escHtml(item.icmId ?? '')}')">↩ Re-run</button>
        <button class="history-delete-btn" onclick="deleteHistory('${escHtml(item.planId)}', this)">✕</button>
      </div>
    `;
    list.appendChild(el);
  }
}

function rerunDri(icmId) {
  if (!icmId || icmId === 'UNKNOWN') return;
  document.getElementById('chatInput').value = `/DRI ${icmId}`;
  switchTab('dri');
  sendMessage();
}

function deleteHistory(planId, btn) {
  fetch(`/api/history/${planId}`, { method: 'DELETE' })
    .then(() => btn.closest('.history-item').remove())
    .catch(() => showToast('Failed to delete', 'error'));
}

/* ─── Activity Summary ───────────────────────────────────────────────────────── */
function renderActivitySummary(summary, isNew = false) {
  const feed = document.getElementById('activityFeed');
  if (!feed || !summary) return;

  document.getElementById('activityEmpty').style.display = 'none';

  const card = document.createElement('div');
  card.className = 'activity-card' + (isNew ? ' activity-card--new' : '');

  const genTime = new Date(summary.generatedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  const sectionsHtml = (summary.sections || []).map((s) => {
    if (!s.items?.length) return '';
    const items = s.items.map((it) => `<li>${escHtml(it)}</li>`).join('');
    return `<div class="act-section">
      <div class="act-section-title">${s.icon} ${escHtml(s.title)}</div>
      <ul class="act-items">${items}</ul>
    </div>`;
  }).join('');

  const insightsHtml = (summary.insights || []).map((i) =>
    `<span class="act-insight">💡 ${escHtml(i)}</span>`
  ).join('');

  const browsersHtml = (summary.browsers || []).slice(0, 4).map((b) =>
    `<a class="act-browser-link" href="${escHtml(b.url)}" target="_blank" title="${escHtml(b.url)}">
      <span class="act-browser-time">${escHtml(b.time)}</span>
      <span class="act-browser-title">${escHtml(b.title)}</span>
    </a>`
  ).join('');

  card.innerHTML = `
    <div class="act-header">
      <div class="act-headline">${escHtml(summary.headline)}</div>
      <div class="act-meta">
        <span class="act-period">⏱ ${escHtml(summary.periodLabel)}</span>
        <span class="act-generated">Generated ${genTime}</span>
        ${isNew ? '<span class="act-badge-new">NEW</span>' : ''}
      </div>
    </div>
    <div class="act-body">
      <div class="act-sections">${sectionsHtml}</div>
      ${browsersHtml ? `<div class="act-browsers"><div class="act-section-title">🌐 Recent Browser Activity</div>${browsersHtml}</div>` : ''}
    </div>
    ${insightsHtml ? `<div class="act-insights">${insightsHtml}</div>` : ''}
  `;

  // Prepend newest at top
  feed.insertBefore(card, feed.firstChild);

  // Keep max 5 cards in DOM
  while (feed.children.length > 6) feed.removeChild(feed.lastChild);

  if (isNew) {
    showToast('📊 Activity summary updated', 'info');
    // Switch to dashboard tab if not already there
  }
}

function setActivityBtnReady() {
  const btn = document.getElementById('activityGenBtn');
  if (!btn) return;
  btn.disabled = false;
  btn.textContent = '↻ Generate Now';
}

function triggerActivityGenerate() {
  const btn = document.getElementById('activityGenBtn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Generating…'; }
  fetch('/api/activity-summary/generate', { method: 'POST' })
    .then((r) => safeJson(r))
    .then(() => showToast('Generating activity summary…', 'info'))
    .catch((err) => {
      showToast('❌ Request failed: ' + err.message, 'error');
      setActivityBtnReady();
    });
}

function loadActivitySummary() {
  fetch('/api/activity-summary')
    .then((r) => safeJson(r))
    .then(({ latest, history }) => {
      const items = latest ? [latest, ...(history || []).filter((h) => h.generatedAt !== latest.generatedAt)] : (history || []);
      items.slice(0, 5).reverse().forEach((s) => renderActivitySummary(s, false));
    })
    .catch(() => {});
}

/* ─── Keyboard Shortcut ─────────────────────────────────────────────────────── */
document.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && (e.metaKey || e.ctrlKey) && document.activeElement === document.getElementById('chatInput')) {
    sendMessage();
  }
});

restoreChatHistory();
loadActivitySummary();

// Always restore ICM token to server on page load (server loses it on restart)
(function restoreIcmTokenOnLoad() {
  rebuildTeamDropdown(); // populate dropdown with saved teams + selection

  function applyTokenToUI(label, isAmbient) {
    const aliasEl = document.getElementById('icmAliasLabel');
    if (label) {
      aliasEl.innerHTML = isAmbient
        ? `${label} <span style="background:rgba(34,197,94,0.15);color:#4ade80;border:1px solid rgba(34,197,94,0.3);border-radius:4px;padding:1px 5px;font-size:0.75em;margin-left:4px;">⚡ auto</span>`
        : label;
      aliasEl.style.display = 'inline';
    }
    document.getElementById('icmFilters').style.display = 'flex';
    document.getElementById('icmRefreshBtn').style.display = 'inline-flex';
  }

  // Try ambient-mcp token first
  fetch('/api/ambient/icm-token')
    .then(r => safeJson(r))
    .then(d => {
      if (d.found && d.token) {
        // Store in localStorage so it persists
        localStorage.setItem(ICM_TOKEN_KEY, d.token);
        const teamId = localStorage.getItem(ICM_TEAM_KEY);
        const label = [d.alias || 'ambient token', teamId ? `team ${teamId}` : ''].filter(Boolean).join(' · ');
        applyTokenToUI(label, true);
        showToast('⚡ ICM token auto-loaded from ambient-mcp', 'success');
        // Re-apply team filter if saved
        if (teamId) {
          fetch('/api/icm/token', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ token: d.token, teamId: Number(teamId) }),
          }).catch(() => {});
        }
      } else {
        // Fall back to localStorage token
        const token = localStorage.getItem(ICM_TOKEN_KEY);
        const teamId = localStorage.getItem(ICM_TEAM_KEY);
        if (!token) return;
        fetch('/api/icm/token', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ token, teamId: teamId ? Number(teamId) : null }),
        }).then(r => safeJson(r)).then(data => {
          const label = [data.alias, data.teamId ? `team ${data.teamId}` : ''].filter(Boolean).join(' · ');
          applyTokenToUI(label, false);
        }).catch(() => {});
      }
    })
    .catch(() => {
      // ambient endpoint failed — fall back to localStorage
      const token = localStorage.getItem(ICM_TOKEN_KEY);
      const teamId = localStorage.getItem(ICM_TEAM_KEY);
      if (!token) return;
      fetch('/api/icm/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ token, teamId: teamId ? Number(teamId) : null }),
      }).then(r => safeJson(r)).then(data => {
        const label = [data.alias, data.teamId ? `team ${data.teamId}` : ''].filter(Boolean).join(' · ');
        applyTokenToUI(label, false);
      }).catch(() => {});
    });
})();

connect();
