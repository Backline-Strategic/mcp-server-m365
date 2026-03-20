/**
 * Microsoft To-Do operations.
 * All functions accept a graphClient: { fetch, write, userId } from getGraphClient().
 */

import { graphGet } from '../graph.mjs';
import { withErrorHandler, makeError } from '../lib/errors.mjs';
import { compact, toUtcIso } from '../lib/response.mjs';

// ── Helpers ───────────────────────────────────────────────────────────────────

async function getAllLists(graphClient) {
  const lists = await graphGet(
    graphClient.fetch,
    `/users/${encodeURIComponent(graphClient.userId)}/todo/lists`,
    { $top: '50' }
  );
  return lists ?? [];
}

async function getListId(graphClient, listName) {
  const lists = await getAllLists(graphClient);
  const lower = listName?.toLowerCase();
  const match = lower
    ? lists.find(l => l.displayName?.toLowerCase() === lower)
    : lists.find(l => l.wellknownListName === 'defaultList');

  if (!match) {
    const available = lists.map(l => l.displayName).join(', ');
    throw new Error(`List "${listName}" not found. Available: ${available}`);
  }
  return match.id;
}

async function getInboxListId(graphClient) {
  const lists = await getAllLists(graphClient);
  const inbox = lists.find(l => l.displayName?.toLowerCase() === 'inbox')
    ?? lists.find(l => l.wellknownListName === 'defaultList');
  if (!inbox) throw new Error('No Inbox or default To-Do list found');
  return inbox.id;
}

async function findTaskByTitle(graphClient, title, listName) {
  const lists = await getAllLists(graphClient);
  const searchLists = listName
    ? lists.filter(l => l.displayName?.toLowerCase() === listName.toLowerCase())
    : lists;

  const lowerTitle = title.toLowerCase();

  for (const list of searchLists) {
    const tasks = await graphGet(
      graphClient.fetch,
      `/users/${encodeURIComponent(graphClient.userId)}/todo/lists/${list.id}/tasks`,
      { $top: '100' }
    );
    const match = (tasks ?? []).find(t => t.title?.toLowerCase().includes(lowerTitle));
    if (match) return { listId: list.id, task: match };
  }

  throw new Error(`No task matching "${title}"${listName ? ` in list "${listName}"` : ''}`);
}

function formatTask(task, listDisplayName) {
  return compact({
    id: task.id,
    title: task.title,
    status: task.status,
    importance: task.importance,
    dueDate: task.dueDateTime?.dateTime
      ? toUtcIso(task.dueDateTime.dateTime).slice(0, 10)
      : undefined,
    list: listDisplayName,
  });
}

function buildDueDatePayload(dueDate) {
  if (dueDate === null) return { dueDateTime: null };
  if (dueDate) {
    return {
      dueDateTime: {
        dateTime: new Date(dueDate).toISOString().slice(0, 10) + 'T00:00:00.000Z',
        timeZone: 'UTC',
      },
    };
  }
  return {};
}

// ── Exported tool functions ───────────────────────────────────────────────────

/**
 * List incomplete tasks from To-Do Inbox or a specific list.
 */
export async function listTasks(graphClient, {
  listName,
  status,
  limit = 50,
} = {}) {
  return withErrorHandler('todo_list_tasks', async () => {
    const lists = await getAllLists(graphClient);
    const targetLists = listName
      ? lists.filter(l => l.displayName?.toLowerCase() === listName.toLowerCase())
      : lists.filter(l =>
          l.wellknownListName === 'defaultList' ||
          l.wellknownListName === 'flaggedEmails' ||
          l.displayName?.toLowerCase() === 'inbox'
        );

    const allTasks = [];
    for (const list of targetLists) {
      const tasks = await graphGet(
        graphClient.fetch,
        `/users/${encodeURIComponent(graphClient.userId)}/todo/lists/${list.id}/tasks`,
        { $top: String(limit) }
      );
      for (const t of (tasks ?? [])) {
        const matchesStatus = !status || t.status === status;
        const isNotDone = !status && t.status !== 'completed';
        if (matchesStatus || isNotDone) {
          allTasks.push(formatTask(t, list.displayName));
        }
      }
    }

    return { success: true, count: allTasks.length, tasks: allTasks };
  });
}

/**
 * Get a single task by ID or title search.
 */
export async function getTask(graphClient, { taskId, title, listName }) {
  return withErrorHandler('todo_get_task', async () => {
    if (!taskId && !title) return makeError('Provide taskId or title');

    let listId, task;
    if (taskId) {
      const lists = await getAllLists(graphClient);
      for (const list of lists) {
        try {
          const t = await graphClient.fetch(
            `/users/${encodeURIComponent(graphClient.userId)}/todo/lists/${list.id}/tasks/${taskId}`
          );
          if (t?.id) { listId = list.id; task = t; break; }
        } catch { /* not in this list */ }
      }
      if (!task) return makeError(`Task ID "${taskId}" not found in any list`);
    } else {
      const found = await findTaskByTitle(graphClient, title, listName);
      listId = found.listId;
      task = found.task;
    }

    const lists = await getAllLists(graphClient);
    const list = lists.find(l => l.id === listId);
    return { success: true, task: formatTask(task, list?.displayName) };
  });
}

/**
 * Create one task (title) or multiple tasks (titles).
 */
export async function createTask(graphClient, {
  title,
  titles,
  dueDate,
  importance = 'normal',
  listName,
} = {}) {
  return withErrorHandler('todo_create_task', async () => {
    if (!title && (!titles || titles.length === 0)) {
      return makeError('Provide title (string) or titles (array)');
    }

    const listId = listName
      ? await getListId(graphClient, listName)
      : await getInboxListId(graphClient);

    const taskList = titles ?? [title];
    const results = [];

    for (const t of taskList) {
      const payload = {
        title: t,
        importance,
        status: 'notStarted',
        ...buildDueDatePayload(dueDate),
      };
      const created = await graphClient.write(
        `/users/${encodeURIComponent(graphClient.userId)}/todo/lists/${listId}/tasks`,
        payload
      );
      results.push({ title: created.title, id: created.id, status: 'created' });
    }

    return {
      success: true,
      created: results.length,
      tasks: results,
    };
  });
}

/**
 * Update a task: status, title, due date, importance.
 */
export async function updateTask(graphClient, {
  taskId,
  title,
  listName,
  status,
  newTitle,
  dueDate,
  importance,
} = {}) {
  return withErrorHandler('todo_update_task', async () => {
    if (!taskId && !title) return makeError('Provide taskId or title to identify the task');

    let listId, resolvedTaskId;

    if (taskId) {
      const lists = await getAllLists(graphClient);
      for (const list of lists) {
        try {
          const t = await graphClient.fetch(
            `/users/${encodeURIComponent(graphClient.userId)}/todo/lists/${list.id}/tasks/${taskId}`
          );
          if (t?.id) { listId = list.id; resolvedTaskId = t.id; break; }
        } catch { /* not in this list */ }
      }
      if (!listId) return makeError(`Task ID "${taskId}" not found in any list`);
    } else {
      const found = await findTaskByTitle(graphClient, title, listName);
      listId = found.listId;
      resolvedTaskId = found.task.id;
    }

    const patch = {
      ...(status ? { status } : {}),
      ...(newTitle ? { title: newTitle } : {}),
      ...(importance ? { importance } : {}),
      ...buildDueDatePayload(dueDate),
    };

    if (Object.keys(patch).length === 0) {
      return makeError('No fields to update — provide status, newTitle, dueDate, or importance');
    }

    const result = await graphClient.write(
      `/users/${encodeURIComponent(graphClient.userId)}/todo/lists/${listId}/tasks/${resolvedTaskId}`,
      patch,
      'PATCH'
    );

    return {
      success: true,
      task: formatTask(result ?? { id: resolvedTaskId, ...patch }),
    };
  });
}

/**
 * Delete a task by ID or title search.
 */
export async function deleteTask(graphClient, { taskId, title, listName } = {}) {
  return withErrorHandler('todo_delete_task', async () => {
    if (!taskId && !title) return makeError('Provide taskId or title');

    let listId, resolvedTaskId, resolvedTitle;

    if (taskId) {
      const lists = await getAllLists(graphClient);
      for (const list of lists) {
        try {
          const t = await graphClient.fetch(
            `/users/${encodeURIComponent(graphClient.userId)}/todo/lists/${list.id}/tasks/${taskId}`
          );
          if (t?.id) { listId = list.id; resolvedTaskId = t.id; resolvedTitle = t.title; break; }
        } catch { /* not in this list */ }
      }
      if (!listId) return makeError(`Task ID "${taskId}" not found`);
    } else {
      const found = await findTaskByTitle(graphClient, title, listName);
      listId = found.listId;
      resolvedTaskId = found.task.id;
      resolvedTitle = found.task.title;
    }

    await graphClient.write(
      `/users/${encodeURIComponent(graphClient.userId)}/todo/lists/${listId}/tasks/${resolvedTaskId}`,
      null,
      'DELETE'
    );

    return { success: true, deleted: { id: resolvedTaskId, title: resolvedTitle } };
  });
}

/**
 * List all To-Do lists for the user.
 */
export async function listLists(graphClient) {
  return withErrorHandler('todo_list_lists', async () => {
    const lists = await getAllLists(graphClient);
    return {
      success: true,
      count: lists.length,
      lists: lists.map(l => ({
        id: l.id,
        name: l.displayName,
        isDefault: l.wellknownListName === 'defaultList',
        isFlagged: l.wellknownListName === 'flaggedEmails',
      })),
    };
  });
}

// ── MCP tool definitions ──────────────────────────────────────────────────────

export const TODO_TOOL_DEFS = [
  {
    name: 'todo_list_tasks',
    description: 'List incomplete tasks from To-Do Inbox or a specific list.',
    inputSchema: {
      type: 'object',
      properties: {
        listName: { type: 'string', description: 'List name to query (default: Inbox + default list)' },
        status: { type: 'string', enum: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'], description: 'Filter by status (default: incomplete only)' },
        limit: { type: 'number', description: 'Max tasks per list (default 50)' },
      },
    },
  },
  {
    name: 'todo_get_task',
    description: 'Get a single To-Do task by ID or title search.',
    inputSchema: {
      type: 'object',
      properties: {
        taskId: { type: 'string', description: 'Task ID (faster)' },
        title: { type: 'string', description: 'Title to search for (case-insensitive partial match)' },
        listName: { type: 'string', description: 'Restrict search to this list' },
      },
    },
  },
  {
    name: 'todo_create_task',
    description: 'Create one task (title) or multiple tasks (titles) in To-Do Inbox.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Single task title' },
        titles: { type: 'array', items: { type: 'string' }, description: 'Multiple task titles (batch create)' },
        dueDate: { type: 'string', description: 'Due date YYYY-MM-DD' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Default: normal' },
        listName: { type: 'string', description: 'Target list (default: Inbox)' },
      },
    },
  },
  {
    name: 'todo_update_task',
    description: 'Update a To-Do task: mark complete, rename, change due date, set importance.',
    inputSchema: {
      type: 'object',
      properties: {
        taskId: { type: 'string', description: 'Task ID (faster)' },
        title: { type: 'string', description: 'Title to search for' },
        listName: { type: 'string', description: 'Restrict search to this list' },
        status: { type: 'string', enum: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'] },
        newTitle: { type: 'string', description: 'Rename the task' },
        dueDate: { type: 'string', description: 'New due date YYYY-MM-DD. Pass null to clear.' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'] },
      },
    },
  },
  {
    name: 'todo_delete_task',
    description: 'Delete a To-Do task by ID or title search.',
    inputSchema: {
      type: 'object',
      properties: {
        taskId: { type: 'string', description: 'Task ID (faster)' },
        title: { type: 'string', description: 'Title to search for' },
        listName: { type: 'string', description: 'Restrict search to this list' },
      },
    },
  },
  {
    name: 'todo_list_lists',
    description: 'List all To-Do lists for the user.',
    inputSchema: { type: 'object', properties: {} },
  },
];
