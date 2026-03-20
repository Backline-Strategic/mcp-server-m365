#!/usr/bin/env node

/**
 * mcp-server-m365 — Standalone MCP server for Microsoft 365 via Graph API
 *
 * Supports: To-Do, Calendar, Planner (read), Email, Teams Chats (read), Files (search)
 *
 * Auth: m365-accounts.json config file, OR env vars M365_TENANT_ID/CLIENT_ID/CLIENT_SECRET/USER_ID
 *
 * Register in .mcp.json:
 *   "mcp-server-m365": {
 *     "command": "node",
 *     "args": ["/path/to/mcp-server-m365/src/index.mjs"],
 *     "env": { "M365_ACCOUNTS_FILE": "/path/to/m365-accounts.json" }
 *   }
 */

import { createInterface } from 'node:readline';
import { getGraphClient } from './graph.mjs';
import { TODO_TOOL_DEFS, listTasks, getTask, createTask, updateTask, deleteTask, listLists } from './tools/todo.mjs';
import { CALENDAR_TOOL_DEFS, listEvents, getEvent, searchEvents, createEvent, updateEvent, deleteEvent } from './tools/calendar.mjs';
import { PLANNER_TOOL_DEFS, listTasks as plannerListTasks, getTask as plannerGetTask, listPlans, listBuckets } from './tools/planner.mjs';
import { EMAIL_TOOL_DEFS, searchEmail, getEmail, listUnread, createDraft, updateEmail } from './tools/email.mjs';
import { CHAT_TOOL_DEFS, listChats, getChatMessages } from './tools/chat.mjs';
import { FILES_TOOL_DEFS, searchFiles } from './tools/files.mjs';
import { SNAPSHOT_TOOL_DEFS, snapshot } from './tools/snapshot.mjs';

// ── Tool registry ─────────────────────────────────────────────────────────────

// Each tool def gets a _handler that receives (args, accountName) and returns a result object.
// The account parameter is resolved from args or defaults to 'default'.

function withAccount(fn) {
  return async (args) => {
    const { account = 'default', ...rest } = args;
    const client = await getGraphClient(account);
    return fn(client, rest);
  };
}

const TOOL_HANDLERS = {
  // To-Do
  todo_list_tasks:   withAccount(listTasks),
  todo_get_task:     withAccount(getTask),
  todo_create_task:  withAccount(createTask),
  todo_update_task:  withAccount(updateTask),
  todo_delete_task:  withAccount(deleteTask),
  todo_list_lists:   withAccount(listLists),

  // Calendar
  calendar_list_events:   withAccount(listEvents),
  calendar_get_event:     withAccount(getEvent),
  calendar_search_events: withAccount(searchEvents),
  calendar_create_event:  withAccount(createEvent),
  calendar_update_event:  withAccount(updateEvent),
  calendar_delete_event:  withAccount(deleteEvent),

  // Planner
  planner_list_tasks:    withAccount(plannerListTasks),
  planner_get_task:      withAccount(plannerGetTask),
  planner_list_plans:    withAccount(listPlans),
  planner_list_buckets:  withAccount(listBuckets),

  // Email
  email_search:       withAccount(searchEmail),
  email_get:          withAccount(getEmail),
  email_list_unread:  withAccount(listUnread),
  email_draft_create: withAccount(createDraft),
  email_update:       withAccount(updateEmail),

  // Chats
  chat_list:          withAccount(listChats),
  chat_get_messages:  withAccount(getChatMessages),

  // Files
  files_search: withAccount(searchFiles),

  // Snapshot
  m365_snapshot: withAccount(snapshot),
};

// Add 'account' param to all tool definitions
const accountSchema = {
  account: {
    type: 'string',
    description: 'Account name from m365-accounts.json (omit to use "default")',
  },
};

function addAccountParam(toolDef) {
  return {
    ...toolDef,
    inputSchema: {
      ...toolDef.inputSchema,
      properties: {
        account: accountSchema.account,
        ...(toolDef.inputSchema.properties ?? {}),
      },
    },
  };
}

const ALL_TOOL_DEFS = [
  ...TODO_TOOL_DEFS,
  ...CALENDAR_TOOL_DEFS,
  ...PLANNER_TOOL_DEFS,
  ...EMAIL_TOOL_DEFS,
  ...CHAT_TOOL_DEFS,
  ...FILES_TOOL_DEFS,
  ...SNAPSHOT_TOOL_DEFS,
].map(addAccountParam);

// ── MCP protocol ──────────────────────────────────────────────────────────────

function send(id, result) {
  process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id, result }) + '\n');
}

function sendError(id, code, message) {
  process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id, error: { code, message } }) + '\n');
}

const rl = createInterface({ input: process.stdin, terminal: false });

rl.on('line', async (line) => {
  const raw = line.trim();
  if (!raw) return;

  let msg;
  try {
    msg = JSON.parse(raw);
  } catch {
    return;
  }

  const { id, method, params } = msg;

  try {
    if (method === 'initialize') {
      send(id, {
        protocolVersion: '2024-11-05',
        capabilities: { tools: {} },
        serverInfo: { name: 'mcp-server-m365', version: '0.1.0' },
      });
      return;
    }

    if (method?.startsWith('notifications/')) return;

    if (method === 'tools/list') {
      send(id, { tools: ALL_TOOL_DEFS });
      return;
    }

    if (method === 'tools/call') {
      const { name, arguments: args = {} } = params;
      const handler = TOOL_HANDLERS[name];

      if (!handler) {
        sendError(id, -32601, `Unknown tool: ${name}`);
        return;
      }

      const result = await handler(args);
      const text = JSON.stringify(result);
      const isError = result.success === false;
      send(id, { content: [{ type: 'text', text }], isError });
      return;
    }

    sendError(id, -32601, `Method not found: ${method}`);

  } catch (err) {
    process.stderr.write(`[mcp-server-m365] error: ${err.message}\n`);
    sendError(id, -32603, `Internal error: ${err.message}`);
  }
});

process.on('SIGTERM', () => process.exit(0));
process.on('SIGINT', () => process.exit(0));

process.stderr.write('[mcp-server-m365] ready\n');
