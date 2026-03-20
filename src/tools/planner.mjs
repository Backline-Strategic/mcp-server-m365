/**
 * Microsoft Planner operations — read-only.
 * Write operations require Tasks.ReadWrite.All + admin consent.
 * All functions accept graphClient: { fetch, write, userId }
 */

import { graphGet } from '../graph.mjs';
import { withErrorHandler, makeError } from '../lib/errors.mjs';
import { compact } from '../lib/response.mjs';

function formatPlannerTask(t) {
  return compact({
    id: t.id,
    title: t.title,
    planId: t.planId,
    bucketId: t.bucketId,
    percentComplete: t.percentComplete,
    dueDateTime: t.dueDateTime,
    priority: t.priority,
    assignedTo: t.assignments
      ? Object.keys(t.assignments).map(userId => ({ userId }))
      : undefined,
  });
}

/**
 * List incomplete Planner tasks assigned to the user.
 */
export async function listTasks(graphClient, { planId, bucketId, limit = 50 } = {}) {
  return withErrorHandler('planner_list_tasks', async () => {
    const params = {
      $select: 'id,title,planId,bucketId,percentComplete,dueDateTime,priority,assignments',
      $top: String(limit),
    };

    let tasks;
    if (planId && bucketId) {
      tasks = await graphGet(graphClient.fetch, `/planner/buckets/${bucketId}/tasks`, params);
    } else if (planId) {
      tasks = await graphGet(graphClient.fetch, `/planner/plans/${planId}/tasks`, params);
    } else {
      tasks = await graphGet(
        graphClient.fetch,
        `/users/${encodeURIComponent(graphClient.userId)}/planner/tasks`,
        params
      );
    }

    const incomplete = (tasks ?? []).filter(t => t.percentComplete < 100);
    return { success: true, count: incomplete.length, tasks: incomplete.map(formatPlannerTask) };
  });
}

/**
 * Get a single Planner task with details (description, checklist).
 */
export async function getTask(graphClient, { taskId }) {
  return withErrorHandler('planner_get_task', async () => {
    if (!taskId) return makeError('taskId is required');

    const [task, details] = await Promise.all([
      graphClient.fetch(`/planner/tasks/${taskId}`),
      graphClient.fetch(`/planner/tasks/${taskId}/details`).catch(() => null),
    ]);

    return {
      success: true,
      task: compact({
        ...formatPlannerTask(task),
        description: details?.description,
        checklist: details?.checklist
          ? Object.values(details.checklist).map(item => ({
              title: item.title,
              isChecked: item.isChecked,
            }))
          : undefined,
        references: details?.references
          ? Object.entries(details.references).map(([url, ref]) => ({
              url,
              alias: ref.alias,
            }))
          : undefined,
      }),
    };
  });
}

/**
 * List plans the user has access to (via group membership).
 */
export async function listPlans(graphClient, { limit = 20 } = {}) {
  return withErrorHandler('planner_list_plans', async () => {
    // Get user's groups, then their plans
    const groups = await graphGet(
      graphClient.fetch,
      `/users/${encodeURIComponent(graphClient.userId)}/memberOf`,
      { $select: 'id,displayName', $top: String(limit) }
    );

    const plans = [];
    for (const group of (groups ?? []).slice(0, 10)) {
      try {
        const groupPlans = await graphGet(
          graphClient.fetch,
          `/groups/${group.id}/planner/plans`,
          { $select: 'id,title,owner' }
        );
        for (const p of (groupPlans ?? [])) {
          plans.push({ id: p.id, title: p.title, groupId: group.id, groupName: group.displayName });
        }
      } catch { /* group may not have Planner */ }
    }

    return { success: true, count: plans.length, plans };
  });
}

/**
 * List buckets in a plan.
 */
export async function listBuckets(graphClient, { planId }) {
  return withErrorHandler('planner_list_buckets', async () => {
    if (!planId) return makeError('planId is required');
    const buckets = await graphGet(
      graphClient.fetch,
      `/planner/plans/${planId}/buckets`,
      { $select: 'id,name,orderHint' }
    );
    return {
      success: true,
      planId,
      count: (buckets ?? []).length,
      buckets: (buckets ?? []).map(b => ({ id: b.id, name: b.name })),
    };
  });
}

// ── MCP tool definitions ──────────────────────────────────────────────────────

export const PLANNER_TOOL_DEFS = [
  {
    name: 'planner_list_tasks',
    description: 'List incomplete Planner tasks assigned to the user. Optionally filter by plan or bucket.',
    inputSchema: {
      type: 'object',
      properties: {
        planId: { type: 'string', description: 'Filter to tasks in this plan' },
        bucketId: { type: 'string', description: 'Filter to tasks in this bucket' },
        limit: { type: 'number', description: 'Max tasks (default 50)' },
      },
    },
  },
  {
    name: 'planner_get_task',
    description: 'Get a Planner task with full details: description, checklist, references.',
    inputSchema: {
      type: 'object',
      properties: {
        taskId: { type: 'string' },
      },
      required: ['taskId'],
    },
  },
  {
    name: 'planner_list_plans',
    description: 'List Planner plans the user has access to via group membership.',
    inputSchema: {
      type: 'object',
      properties: {
        limit: { type: 'number', description: 'Max groups to check (default 20)' },
      },
    },
  },
  {
    name: 'planner_list_buckets',
    description: 'List buckets in a Planner plan.',
    inputSchema: {
      type: 'object',
      properties: {
        planId: { type: 'string' },
      },
      required: ['planId'],
    },
  },
];
