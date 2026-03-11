import { spawn } from 'child_process';
import { v4 as uuidv4 } from 'uuid';
import { Plan, PlanJsonSchema } from './models';
import { WSManager } from './ws-manager';

const PLAN_PROMPT = `You are a software engineering orchestrator. Break the given task into 2-6 concrete subtasks.

Respond with ONLY valid JSON, no markdown, no explanation:
{
  "title": "short plan title",
  "description": "one sentence summary",
  "estimated_duration": "e.g. 10 minutes",
  "tasks": [
    {
      "id": "task_1",
      "title": "short title",
      "description": "what this task does",
      "claude_prompt": "Complete self-contained prompt for another Claude instance. Include all context needed. End with: Work in the provided workspace directory.",
      "complexity": "low",
      "dependencies": [],
      "parallel_ok": false
    }
  ]
}

Rules:
- complexity: "low" (<5min), "medium" (5-15min), "high" (>15min)
- dependencies: array of task IDs that must finish first ([] for none)
- parallel_ok: true if safe to run concurrently with siblings
- claude_prompt must be fully self-contained with all context

Available local proxy endpoints (use Bash/curl, no auth needed):
- Microsoft Graph: GET http://localhost:3333/api/graph?path=/me  (any Graph v1.0 path)
  e.g. /me, /me/memberOf, /me/manager, /users/{id}, /groups
- Kusto/ADX (IcmDataWarehouse): GET http://localhost:3333/api/adx/icm/{icmId}
- ICM active incidents: GET http://localhost:3333/api/icm/active`;

const TIMEOUT_MS = 120_000; // 2 minutes

export class Orchestrator {
  private wsManager: WSManager;

  constructor(wsManager: WSManager) {
    this.wsManager = wsManager;
  }

  async generatePlan(userMessage: string): Promise<Plan> {
    console.log('[Orchestrator] Generating plan for:', userMessage);

    const prompt = `${PLAN_PROMPT}\n\nTask:\n${userMessage}`;
    const fullText = await this.runClaude(prompt);

    const planData = this.parsePlanJson(fullText);

    const workspaceBase = `${process.env.HOME}/Work6/experiments/orchestrator-workspaces`;
    const sessionId = `session_${Date.now()}`;

    const plan: Plan = {
      id: uuidv4(),
      title: planData.title,
      description: planData.description,
      estimated_duration: planData.estimated_duration,
      createdAt: Date.now(),
      tasks: planData.tasks.map((t: RawTask) => ({
        ...t,
        status: 'pending' as const,
        output: '',
        workspace: `${workspaceBase}/${sessionId}/task_${t.id}`,
      })),
    };

    console.log(`[Orchestrator] Plan generated: ${plan.tasks.length} tasks`);
    return plan;
  }

  private runClaude(prompt: string): Promise<string> {
    return new Promise((resolve, reject) => {
      const env = { ...process.env };
      delete env['CLAUDECODE'];
      delete env['CLAUDE_CODE_ENTRYPOINT'];
      delete env['ANTHROPIC_API_KEY'];

      const proc = spawn('claude', [
        '-p', prompt,
        '--output-format', 'stream-json',
        '--verbose',
        '--model', 'sonnet',
      ], { env, stdio: ['ignore', 'pipe', 'pipe'] });

      let buffer = '';
      let fullText = '';
      let resolved = false;

      // Hard timeout
      const timer = setTimeout(() => {
        if (!resolved) {
          proc.kill();
          reject(new Error('Plan generation timed out after 2 minutes'));
        }
      }, TIMEOUT_MS);

      proc.stdout.on('data', (data: Buffer) => {
        buffer += data.toString();
        const lines = buffer.split('\n');
        buffer = lines.pop() ?? '';

        for (const line of lines) {
          if (!line.trim()) continue;
          try {
            const event = JSON.parse(line) as StreamEvent;

            // Only capture text from assistant events (not result — it duplicates)
            if (event.type === 'assistant') {
              const content = event.message?.content;
              if (Array.isArray(content)) {
                const text = content
                  .filter((c) => c.type === 'text')
                  .map((c) => c.text ?? '')
                  .join('');
                if (text) {
                  fullText += text;
                  this.wsManager.broadcast({ type: 'plan_token', token: text });
                }
              }
            }
          } catch {
            // ignore non-JSON lines
          }
        }
      });

      proc.stderr.on('data', (data: Buffer) => {
        const msg = data.toString().trim();
        if (msg) console.error('[Orchestrator] stderr:', msg);
      });

      proc.on('close', (code) => {
        clearTimeout(timer);
        resolved = true;

        // Flush remaining buffer
        if (buffer.trim()) {
          try {
            const event = JSON.parse(buffer) as StreamEvent;
            if (event.type === 'assistant') {
              const content = event.message?.content;
              if (Array.isArray(content)) {
                const text = content.filter((c) => c.type === 'text').map((c) => c.text ?? '').join('');
                if (text) fullText += text;
              }
            }
          } catch { /* ignore */ }
        }

        if (code === 0 && fullText.trim()) {
          resolve(fullText);
        } else {
          reject(new Error(`claude exited with code ${code}. Output: "${fullText.slice(0, 300)}"`));
        }
      });

      proc.on('error', (err) => {
        clearTimeout(timer);
        reject(new Error(`Failed to spawn claude: ${err.message}`));
      });
    });
  }

  private parsePlanJson(text: string): RawPlan {
    const cleaned = text
      .replace(/^```(?:json)?\s*/m, '')
      .replace(/\s*```\s*$/m, '')
      .trim();

    try {
      const parsed = JSON.parse(cleaned) as RawPlan;
      this.validatePlan(parsed);
      return parsed;
    } catch {
      const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const parsed = JSON.parse(jsonMatch[0]) as RawPlan;
        this.validatePlan(parsed);
        return parsed;
      }
      throw new Error(`Failed to parse plan JSON. Got: ${cleaned.slice(0, 300)}`);
    }
  }

  private validatePlan(plan: RawPlan): void {
    if (!plan.title || !Array.isArray(plan.tasks) || plan.tasks.length === 0) {
      throw new Error('Invalid plan: missing title or tasks');
    }
    for (const task of plan.tasks) {
      if (!task.id || !task.title || !task.claude_prompt) {
        throw new Error(`Invalid task: ${JSON.stringify(task)}`);
      }
    }
  }
}

interface ContentBlock { type: string; text?: string; }
interface StreamEvent {
  type: string;
  message?: { content: ContentBlock[] };
  result?: string;
}
interface RawTask {
  id: string;
  title: string;
  description: string;
  claude_prompt: string;
  complexity: 'low' | 'medium' | 'high';
  dependencies: string[];
  parallel_ok: boolean;
}
interface RawPlan {
  title: string;
  description: string;
  estimated_duration: string;
  tasks: RawTask[];
}

export { PlanJsonSchema };
