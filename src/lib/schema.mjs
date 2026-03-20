/**
 * Reusable JSON Schema fragments for MCP tool definitions.
 */

export const accountParam = {
  account: {
    type: 'string',
    description: 'Account name from m365-accounts.json (omit to use "default")',
  },
};

export const paginationParams = {
  limit: { type: 'number', description: 'Max items to return (default varies per tool)' },
};

export const dateRangeParams = {
  startDate: { type: 'string', description: 'Start date YYYY-MM-DD' },
  endDate: { type: 'string', description: 'End date YYYY-MM-DD' },
};

export const timeZoneParam = {
  timeZone: { type: 'string', description: 'IANA timezone (e.g. America/Edmonton). Defaults to UTC.' },
};

export const taskIdOrTitle = {
  taskId: { type: 'string', description: 'Task ID (if known — faster than title search)' },
  title: { type: 'string', description: 'Task title — case-insensitive partial match' },
  listName: { type: 'string', description: 'Restrict search to this To-Do list name' },
};

export const taskFields = {
  status: {
    type: 'string',
    enum: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'],
    description: 'Task completion status',
  },
  importance: {
    type: 'string',
    enum: ['low', 'normal', 'high'],
    description: 'Task priority',
  },
  dueDate: {
    type: 'string',
    description: 'Due date ISO 8601 (YYYY-MM-DD). Pass null to clear.',
  },
};

export const calendarEventFields = {
  subject: { type: 'string', description: 'Event title' },
  start: { type: 'string', description: 'Start datetime ISO 8601 (e.g. 2026-03-20T10:00:00)' },
  end: { type: 'string', description: 'End datetime ISO 8601' },
  location: { type: 'string', description: 'Location display name' },
  body: { type: 'string', description: 'Event body / notes (plain text)' },
  isAllDay: { type: 'boolean', description: 'All-day event' },
  isOnlineMeeting: { type: 'boolean', description: 'Create Teams meeting link' },
  attendees: {
    type: 'array',
    items: {
      type: 'object',
      properties: {
        email: { type: 'string' },
        name: { type: 'string' },
        type: { type: 'string', enum: ['required', 'optional'], description: 'Default: required' },
      },
      required: ['email'],
    },
    description: 'Attendees array',
  },
};
