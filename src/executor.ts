import { spawn, ChildProcess } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { Plan, Task, ExecutionStats } from './models';
import { WSManager } from './ws-manager';

export class Executor {
  private wsManager: WSManager;
  // Key: "<planId>/<taskId>" to support parallel plans
  private activeProcesses: Map<string, ChildProcess> = new Map();

  constructor(wsManager: WSManager) {
    this.wsManager = wsManager;
  }

  async execute(plan: Plan): Promise<ExecutionStats> {
    const startTime = Date.now();
    console.log(`[Executor] Starting plan: ${plan.id} (${plan.tasks.length} tasks)`);

    for (const task of plan.tasks) {
      task.status = 'pending';
      task.output = '';
    }

    for (const task of plan.tasks) {
      fs.mkdirSync(task.workspace, { recursive: true });
    }

    await this.executeGraph(plan);

    const stats: ExecutionStats = {
      totalTasks: plan.tasks.length,
      completedTasks: plan.tasks.filter((t) => t.status === 'completed').length,
      failedTasks: plan.tasks.filter((t) => t.status === 'failed').length,
      durationMs: Date.now() - startTime,
      planId: plan.id,
    };

    this.wsManager.broadcast({ type: 'execution_done', stats });
    console.log(`[Executor] Plan ${plan.id} done:`, stats);

    // For DRI plans, broadcast the summary from the final reporting task
    if (plan.type === 'dri') {
      const reportTask =
        plan.tasks.find((t) => t.id === 'step_7_generate_report') ??
        plan.tasks.find((t) => t.id === 'generate_summary') ??
        plan.tasks[plan.tasks.length - 1];
      if (reportTask?.output) {
        this.wsManager.broadcast({
          type: 'dri_summary',
          planId: plan.id,
          icmId: plan.icmId,
          content: reportTask.output,
        });
      }
    }

    return stats;
  }

  // ─── Proper parallel DAG scheduler ──────────────────────────────────────────
  // Uses Promise.race to pick up completions instantly and immediately
  // launches any newly unblocked tasks — no batch waiting.
  private async executeGraph(plan: Plan): Promise<void> {
    const { tasks } = plan;
    const completed = new Set<string>();
    const failed = new Set<string>();
    // Maps taskId → promise that resolves to { id, ok }
    const running = new Map<string, Promise<{ id: string; ok: boolean }>>();

    const launch = (task: Task) => {
      const p = this.runTask(plan.id, task).then(() => ({
        id: task.id,
        ok: task.status === 'completed',
      }));
      running.set(task.id, p);
      console.log(`[Executor][${plan.id}] Launched: ${task.id} (parallel_ok=${task.parallel_ok})`);
    };

    const markFailedDeps = () => {
      for (const task of tasks) {
        if (task.status !== 'pending' || running.has(task.id)) continue;
        if (task.dependencies.some((d) => failed.has(d))) {
          task.status = 'failed';
          task.output = 'Skipped: a dependency task failed';
          failed.add(task.id);
          this.wsManager.broadcast({
            type: 'task_completed',
            planId: plan.id,
            taskId: task.id,
            success: false,
          });
        }
      }
    };

    const launchReady = () => {
      for (const task of tasks) {
        if (task.status !== 'pending' || running.has(task.id)) continue;
        // All dependencies must be done
        if (!task.dependencies.every((d) => completed.has(d))) continue;
        // Non-parallel tasks wait until nothing else is running
        if (!task.parallel_ok && running.size > 0) continue;
        launch(task);
      }
    };

    // Kick off the initial wave
    markFailedDeps();
    launchReady();

    // Event loop: each iteration picks up exactly one completion and re-evaluates
    while (running.size > 0) {
      const { id, ok } = await Promise.race(running.values());
      running.delete(id);
      if (ok) completed.add(id);
      else failed.add(id);
      markFailedDeps();
      launchReady();
    }
  }

  private async runTask(planId: string, task: Task): Promise<void> {
    task.status = 'running';
    task.startedAt = Date.now();

    this.wsManager.broadcast({ type: 'task_started', planId, taskId: task.id });
    console.log(`[Executor][${planId}] Running: ${task.id} - ${task.title}`);

    const TASK_TIMEOUT_MS = 10 * 60 * 1000; // 10 minutes per task

    return new Promise((resolve) => {
      const claudeArgs = [
        '-p',
        task.claude_prompt,
        '--output-format', 'stream-json',
        '--verbose',
        '--dangerously-skip-permissions',
      ];

      // DRI tasks need unrestricted tool access (MCP servers: kusto-icm, kusto-spoons, ado, playwright)
      if (task.allowAllTools) {
        const mcpConfigPath = path.join(__dirname, '..', '.mcp.json');
        if (fs.existsSync(mcpConfigPath)) {
          claudeArgs.push('--mcp-config', mcpConfigPath);
        }
      }

      const env = { ...process.env };
      delete env['CLAUDECODE'];
      delete env['CLAUDE_CODE_ENTRYPOINT'];
      delete env['ANTHROPIC_API_KEY'];

      const proc = spawn('claude', claudeArgs, {
        cwd: task.workspace,
        env,
        stdio: ['ignore', 'pipe', 'pipe'],
      });

      const procKey = `${planId}/${task.id}`;
      this.activeProcesses.set(procKey, proc);

      const timeoutHandle = setTimeout(() => {
        console.warn(`[Executor][${planId}] Task ${task.id} timed out after ${TASK_TIMEOUT_MS / 1000}s — killing`);
        proc.kill('SIGTERM');
        task.output += '\n[TIMEOUT] Task exceeded time limit and was killed.';
        this.wsManager.broadcast({
          type: 'task_output',
          planId,
          taskId: task.id,
          chunk: '\n[TIMEOUT] Task exceeded time limit and was killed.\n',
          eventType: 'error',
        });
      }, TASK_TIMEOUT_MS);

      let buffer = '';

      const handleLine = (line: string): void => {
        if (!line.trim()) return;
        try {
          const event = JSON.parse(line) as StreamEvent;
          const chunk = this.extractChunk(event);
          if (chunk) {
            task.output += chunk;
            this.wsManager.broadcast({
              type: 'task_output',
              planId,
              taskId: task.id,
              chunk,
              eventType: event.type,
            });
          }
        } catch {
          if (line.trim()) {
            task.output += line + '\n';
            this.wsManager.broadcast({
              type: 'task_output',
              planId,
              taskId: task.id,
              chunk: line + '\n',
              eventType: 'raw',
            });
          }
        }
      };

      proc.stdout.on('data', (data: Buffer) => {
        buffer += data.toString();
        const lines = buffer.split('\n');
        buffer = lines.pop() ?? '';
        for (const line of lines) handleLine(line);
      });

      proc.stderr.on('data', (data: Buffer) => {
        const text = data.toString();
        console.error(`[Executor][${planId}/${task.id}] stderr:`, text);
        task.output += `[stderr]: ${text}`;
        this.wsManager.broadcast({
          type: 'task_output',
          planId,
          taskId: task.id,
          chunk: `[stderr]: ${text}`,
          eventType: 'error',
        });
      });

      proc.on('close', (code) => {
        clearTimeout(timeoutHandle);
        if (buffer.trim()) handleLine(buffer);
        this.activeProcesses.delete(procKey);
        task.completedAt = Date.now();
        task.status = code === 0 ? 'completed' : 'failed';
        console.log(`[Executor][${planId}] Task ${task.id} exited ${code}`);
        this.wsManager.broadcast({
          type: 'task_completed',
          planId,
          taskId: task.id,
          success: code === 0,
        });
        resolve();
      });

      proc.on('error', (err) => {
        console.error(`[Executor] Spawn failed for ${planId}/${task.id}:`, err);
        task.output += `Failed to spawn claude: ${err.message}`;
        task.status = 'failed';
        task.completedAt = Date.now();
        this.activeProcesses.delete(procKey);
        this.wsManager.broadcast({
          type: 'task_completed',
          planId,
          taskId: task.id,
          success: false,
        });
        resolve();
      });
    });
  }

  private extractChunk(event: StreamEvent): string {
    switch (event.type) {
      case 'assistant': {
        const content = event.message?.content;
        if (!Array.isArray(content)) return '';
        return content
          .filter((c: ContentBlock) => c.type === 'text')
          .map((c: ContentBlock) => c.text ?? '')
          .join('');
      }
      case 'tool_use':
        return `\n[Tool: ${event.tool_name ?? ''}] ${JSON.stringify(event.tool_input ?? {})}\n`;
      case 'tool_result':
        return `[Result]: ${String(event.content ?? '').slice(0, 500)}\n`;
      case 'result':
        return event.result ? `\n--- Result ---\n${event.result}\n` : '';
      default:
        return '';
    }
  }

  cancelTask(planId: string, taskId: string): boolean {
    const key = `${planId}/${taskId}`;
    const proc = this.activeProcesses.get(key);
    if (proc) {
      proc.kill('SIGTERM');
      this.activeProcesses.delete(key);
      console.log(`[Executor] Cancelled ${key}`);
      return true;
    }
    return false;
  }

  cancelPlan(planId: string): void {
    for (const [key, proc] of this.activeProcesses) {
      if (key.startsWith(`${planId}/`)) {
        proc.kill('SIGTERM');
        this.activeProcesses.delete(key);
      }
    }
  }
}

interface ContentBlock { type: string; text?: string; }
interface StreamEvent {
  type: string;
  message?: { content: ContentBlock[] };
  tool_name?: string;
  tool_input?: unknown;
  content?: unknown;
  result?: string;
}
