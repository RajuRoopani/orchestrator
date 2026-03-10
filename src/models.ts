export type TaskStatus = 'pending' | 'running' | 'completed' | 'failed';
export type TaskComplexity = 'low' | 'medium' | 'high';

export interface Task {
  id: string;
  title: string;
  description: string;
  claude_prompt: string;
  complexity: TaskComplexity;
  dependencies: string[];
  parallel_ok: boolean;
  status: TaskStatus;
  output: string;
  workspace: string;
  startedAt?: number;
  completedAt?: number;
  /** When true, no --allowedTools restriction is passed — all MCP tools are available */
  allowAllTools?: boolean;
}

export interface Plan {
  id: string;
  title: string;
  description: string;
  estimated_duration: string;
  tasks: Task[];
  createdAt: number;
  type?: 'standard' | 'dri';
  icmId?: string;
}

export interface ExecutionStats {
  totalTasks: number;
  completedTasks: number;
  failedTasks: number;
  durationMs: number;
  planId: string;
}

export type WSMessage =
  | { type: 'plan_token'; token: string }
  | { type: 'plan_ready'; plan: Plan }
  | { type: 'task_started'; planId: string; taskId: string }
  | { type: 'task_output'; planId: string; taskId: string; chunk: string; eventType: string }
  | { type: 'task_completed'; planId: string; taskId: string; success: boolean }
  | { type: 'execution_done'; stats: ExecutionStats }
  | { type: 'dri_summary'; planId: string; content: string; icmId?: string }
  | { type: 'activity_summary'; summary: unknown }
  | { type: 'error'; message: string }
  | { type: 'connected' };

export const PlanJsonSchema = {
  type: 'object',
  properties: {
    title: { type: 'string' },
    description: { type: 'string' },
    estimated_duration: { type: 'string' },
    tasks: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          id: { type: 'string' },
          title: { type: 'string' },
          description: { type: 'string' },
          claude_prompt: { type: 'string' },
          complexity: { type: 'string', enum: ['low', 'medium', 'high'] },
          dependencies: { type: 'array', items: { type: 'string' } },
          parallel_ok: { type: 'boolean' },
        },
        required: ['id', 'title', 'description', 'claude_prompt', 'complexity', 'dependencies', 'parallel_ok'],
        additionalProperties: false,
      },
    },
  },
  required: ['title', 'description', 'estimated_duration', 'tasks'],
  additionalProperties: false,
} as const;
