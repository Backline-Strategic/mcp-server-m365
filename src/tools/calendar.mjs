/**
 * Microsoft Calendar operations.
 * All functions accept graphClient: { fetch, write, userId }
 */

import { graphGet } from '../graph.mjs';
import { withErrorHandler, makeError } from '../lib/errors.mjs';
import { compact, toUtcIso } from '../lib/response.mjs';

const DEFAULT_TZ = 'America/Edmonton';

// ── Helpers ───────────────────────────────────────────────────────────────────

function formatEvent(e, tz) {
  return compact({
    id: e.id,
    subject: e.subject,
    start: e.start?.dateTime,
    end: e.end?.dateTime,
    timeZone: e.start?.timeZone ?? tz,
    isAllDay: e.isAllDay ?? false,
    location: e.location?.displayName,
    organizer: e.organizer?.emailAddress?.name ?? e.organizer?.emailAddress?.address,
    attendees: e.attendees?.length
      ? e.attendees.map(a => compact({
          name: a.emailAddress?.name,
          email: a.emailAddress?.address,
          response: a.status?.response,
        }))
      : undefined,
    isOnlineMeeting: e.isOnlineMeeting,
    onlineMeetingUrl: e.onlineMeetingUrl,
    bodyPreview: e.bodyPreview?.slice(0, 150),
    calendarName: e._calendarName,
  });
}

function buildEventBody(data, isCreate = false) {
  const tz = data.timeZone ?? DEFAULT_TZ;
  const body = {};

  if (isCreate || data.subject) body.subject = data.subject;
  if (data.start) body.start = { dateTime: data.start, timeZone: tz };
  if (data.end) body.end = { dateTime: data.end, timeZone: tz };
  if (data.location) body.location = { displayName: data.location };
  if (data.body) body.body = { contentType: 'text', content: data.body };
  if (data.isAllDay != null) body.isAllDay = data.isAllDay;
  if (data.isOnlineMeeting) body.isOnlineMeeting = true;
  if (data.attendees?.length) {
    body.attendees = data.attendees.map(a => ({
      emailAddress: { address: a.email, name: a.name ?? a.email },
      type: a.type ?? 'required',
    }));
  }

  return body;
}

// ── Exported tool functions ───────────────────────────────────────────────────

/**
 * List events in a date range across all calendars (deduped).
 */
export async function listEvents(graphClient, {
  startDate,
  endDate,
  timeZone = DEFAULT_TZ,
  limit = 50,
} = {}) {
  return withErrorHandler('calendar_list_events', async () => {
    const now = new Date();
    const start = startDate ?? now.toISOString().slice(0, 10);
    const end = endDate ?? start;
    const startDateTime = `${start}T07:00:00`;
    const endDateTime = `${end}T22:00:00`;
    const preferHeader = { Prefer: `outlook.timezone="${timeZone}"` };

    const userPath = `/users/${encodeURIComponent(graphClient.userId)}`;

    // List all calendars
    const calendars = await graphGet(graphClient.fetch, `${userPath}/calendars`, {
      $select: 'id,name',
      $top: '25',
    });
    const calList = Array.isArray(calendars) ? calendars : [{ id: null, name: 'Calendar' }];

    // Query each calendar in parallel
    const allEvents = [];
    await Promise.all(calList.map(async (cal) => {
      try {
        const calPath = cal.id
          ? `${userPath}/calendars/${cal.id}/calendarView`
          : `${userPath}/calendarView`;
        const events = await graphGet(graphClient.fetch, calPath, {
          startDateTime,
          endDateTime,
          $select: 'id,subject,start,end,location,organizer,isAllDay,bodyPreview,attendees,isOnlineMeeting,onlineMeetingUrl,seriesMasterId',
          $orderby: 'start/dateTime asc',
          $top: String(limit),
        }, { headers: preferHeader });

        for (const e of (events ?? [])) {
          allEvents.push({ ...e, _calendarName: cal.name ?? 'Calendar' });
        }
      } catch { /* skip inaccessible calendars */ }
    }));

    // Deduplicate by event ID
    const seen = new Set();
    const deduped = [];
    for (const e of allEvents) {
      if (!seen.has(e.id)) { seen.add(e.id); deduped.push(e); }
    }
    deduped.sort((a, b) => (a.start?.dateTime ?? '').localeCompare(b.start?.dateTime ?? ''));

    return {
      success: true,
      range: { start, end, timeZone },
      count: deduped.length,
      events: deduped.map(e => formatEvent(e, timeZone)),
    };
  });
}

/**
 * Get a single event by ID with full details.
 */
export async function getEvent(graphClient, { eventId }) {
  return withErrorHandler('calendar_get_event', async () => {
    if (!eventId) return makeError('eventId is required');
    const e = await graphClient.fetch(
      `/users/${encodeURIComponent(graphClient.userId)}/events/${eventId}`
    );
    return { success: true, event: formatEvent(e) };
  });
}

/**
 * Search events by keyword in subject/body.
 */
export async function searchEvents(graphClient, { query, startDate, endDate, limit = 25 }) {
  return withErrorHandler('calendar_search_events', async () => {
    if (!query) return makeError('query is required');

    const userPath = `/users/${encodeURIComponent(graphClient.userId)}`;
    const params = {
      $search: `"${query.replace(/"/g, '')}"`,
      $select: 'id,subject,start,end,location,organizer,isOnlineMeeting',
      $top: String(limit),
    };
    if (startDate) params.$filter = `start/dateTime ge '${startDate}T00:00:00'`;

    const events = await graphGet(graphClient.fetch, `${userPath}/events`, params);

    return {
      success: true,
      query,
      count: (events ?? []).length,
      events: (events ?? []).map(e => formatEvent(e)),
    };
  });
}

/**
 * Create a calendar event.
 */
export async function createEvent(graphClient, data) {
  return withErrorHandler('calendar_create_event', async () => {
    if (!data.subject) return makeError('subject is required');
    if (!data.start) return makeError('start is required');
    if (!data.end) return makeError('end is required');

    const body = buildEventBody(data, true);
    const event = await graphClient.write(
      `/users/${encodeURIComponent(graphClient.userId)}/events`,
      body
    );

    return {
      success: true,
      event: compact({
        id: event.id,
        subject: event.subject,
        start: event.start?.dateTime,
        end: event.end?.dateTime,
        webLink: event.webLink,
        onlineMeetingUrl: event.onlineMeeting?.joinUrl,
      }),
    };
  });
}

/**
 * Update an existing calendar event.
 */
export async function updateEvent(graphClient, data) {
  return withErrorHandler('calendar_update_event', async () => {
    if (!data.eventId) return makeError('eventId is required');

    const body = buildEventBody(data, false);
    if (Object.keys(body).length === 0) {
      return makeError('No fields to update — provide subject, start, end, location, body, or attendees');
    }

    const event = await graphClient.write(
      `/users/${encodeURIComponent(graphClient.userId)}/events/${data.eventId}`,
      body,
      'PATCH'
    );

    return {
      success: true,
      event: event
        ? compact({ id: event.id, subject: event.subject, start: event.start?.dateTime, end: event.end?.dateTime })
        : { id: data.eventId, updated: true },
    };
  });
}

/**
 * Delete a calendar event.
 */
export async function deleteEvent(graphClient, { eventId }) {
  return withErrorHandler('calendar_delete_event', async () => {
    if (!eventId) return makeError('eventId is required');
    await graphClient.write(
      `/users/${encodeURIComponent(graphClient.userId)}/events/${eventId}`,
      null,
      'DELETE'
    );
    return { success: true, deleted: { eventId } };
  });
}

// ── MCP tool definitions ──────────────────────────────────────────────────────

export const CALENDAR_TOOL_DEFS = [
  {
    name: 'calendar_list_events',
    description: 'List calendar events in a date range across all calendars.',
    inputSchema: {
      type: 'object',
      properties: {
        startDate: { type: 'string', description: 'Start date YYYY-MM-DD (default: today)' },
        endDate: { type: 'string', description: 'End date YYYY-MM-DD (default: same as startDate)' },
        timeZone: { type: 'string', description: 'IANA timezone (default: America/Edmonton)' },
        limit: { type: 'number', description: 'Max events per calendar (default 50)' },
      },
    },
  },
  {
    name: 'calendar_get_event',
    description: 'Get a single calendar event with full details by event ID.',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'Event ID (from calendar_list_events)' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_search_events',
    description: 'Search calendar events by keyword in subject or body.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search keyword or phrase' },
        startDate: { type: 'string', description: 'Restrict to events on/after this date YYYY-MM-DD' },
        limit: { type: 'number', description: 'Max results (default 25)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'calendar_create_event',
    description: 'Create a calendar event. Default timezone: America/Edmonton.',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string' },
        start: { type: 'string', description: 'ISO 8601 e.g. 2026-03-20T10:00:00' },
        end: { type: 'string', description: 'ISO 8601 e.g. 2026-03-20T11:00:00' },
        timeZone: { type: 'string', description: 'IANA timezone (default: America/Edmonton)' },
        location: { type: 'string' },
        body: { type: 'string', description: 'Plain text notes' },
        isAllDay: { type: 'boolean' },
        isOnlineMeeting: { type: 'boolean', description: 'Create Teams meeting link' },
        attendees: {
          type: 'array',
          items: { type: 'object', properties: { email: { type: 'string' }, name: { type: 'string' }, type: { type: 'string' } }, required: ['email'] },
        },
      },
      required: ['subject', 'start', 'end'],
    },
  },
  {
    name: 'calendar_update_event',
    description: 'Update an existing calendar event by event ID.',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string' },
        subject: { type: 'string' },
        start: { type: 'string' },
        end: { type: 'string' },
        timeZone: { type: 'string' },
        location: { type: 'string' },
        body: { type: 'string' },
        attendees: { type: 'array', items: { type: 'object', properties: { email: { type: 'string' }, name: { type: 'string' } }, required: ['email'] } },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_delete_event',
    description: 'Delete a calendar event by event ID.',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string' },
      },
      required: ['eventId'],
    },
  },
];
