/**
 * Composite M365 snapshot tool.
 *
 * When used standalone (MCP tool): returns compact multi-service context with `include` filter.
 * When used by BKG heartbeat: returns { source, timestamp, data, errors } format
 *   compatible with context-builder.mjs — calendar/todo/planner/email/chats shape preserved.
 *
 * BKG heartbeat compatibility notes:
 *   - context-builder reads: snap.data.calendar[].{title, start, end, isAllDay}
 *   - context-builder reads: snap.data.todo[].{title, dueDateTime, importance, status}
 *   - context-builder reads: snap.data.planner[].{title, dueDateTime, percentComplete}
 *   - context-builder reads: snap.data.email[].{subject, from, receivedDateTime, isRead, flagged, preview}
 *   - DO NOT change these field names
 */

import { graphGet } from '../graph.mjs';
import { compact, toUtcIso } from '../lib/response.mjs';

const SERVICES = ['calendar', 'todo', 'planner', 'email', 'chats', 'files'];

/**
 * Full M365 context snapshot.
 *
 * @param {object} graphClient - { fetch, write, userId }
 * @param {object} options
 * @param {string[]} [options.include] - Services to include (default: all)
 * @param {number} [options.calendarHours] - Hours ahead for calendar (default: end of day)
 * @param {boolean} [options.bkgFormat] - If true, return BKG heartbeat-compatible format
 * @param {string} [options.source] - Source key for BKG format (e.g. 'backline-m365')
 */
export async function snapshot(graphClient, {
  include = SERVICES,
  calendarHours,
  bkgFormat = false,
  source = 'default',
} = {}) {
  const data = {};
  const errors = [];
  const toInclude = new Set(include);

  // ── Calendar ──────────────────────────────────────────────────────────────
  if (toInclude.has('calendar')) {
    try {
      const now = new Date();
      const endOfDay = new Date(now);
      if (calendarHours) {
        endOfDay.setTime(now.getTime() + calendarHours * 60 * 60 * 1000);
      } else {
        endOfDay.setHours(23, 59, 59, 999);
      }

      const events = await graphGet(
        graphClient.fetch,
        `/users/${encodeURIComponent(graphClient.userId)}/calendarView`,
        {
          startDateTime: now.toISOString(),
          endDateTime: endOfDay.toISOString(),
          $select: 'subject,start,end,location,organizer,isAllDay',
          $orderby: 'start/dateTime asc',
          $top: '20',
        }
      );

      const toUtc = dt => dt ? (dt.endsWith('Z') ? dt : dt + 'Z') : null;
      data.calendar = (events ?? []).map(e => ({
        title: e.subject,
        start: toUtc(e.start?.dateTime),
        end: toUtc(e.end?.dateTime),
        isAllDay: e.isAllDay,
        location: e.location?.displayName ?? null,
        organizer: e.organizer?.emailAddress?.name ?? null,
      }));
    } catch (err) {
      errors.push({ source: 'calendar', message: err.message });
      data.calendar = [];
    }
  }

  // ── Todo ─────────────────────────────────────────────────────────────────
  if (toInclude.has('todo')) {
    try {
      const lists = await graphGet(
        graphClient.fetch,
        `/users/${encodeURIComponent(graphClient.userId)}/todo/lists`,
        { $top: '20' }
      );

      const targetLists = (lists ?? []).filter(l =>
        l.wellknownListName === 'defaultList' ||
        l.wellknownListName === 'flaggedEmails' ||
        l.displayName?.toLowerCase() === 'inbox'
      );

      const allTasks = [];
      for (const list of targetLists) {
        const tasks = await graphGet(
          graphClient.fetch,
          `/users/${encodeURIComponent(graphClient.userId)}/todo/lists/${list.id}/tasks`,
          { $top: '100' }
        );
        for (const t of (tasks ?? [])) {
          if (t.status !== 'completed') {
            allTasks.push({
              title: t.title,
              list: list.displayName,
              dueDateTime: t.dueDateTime?.dateTime ?? null,
              importance: t.importance,
              status: t.status,
            });
          }
        }
      }
      data.todo = allTasks;
    } catch (err) {
      errors.push({ source: 'todo', message: err.message });
      data.todo = [];
    }
  }

  // ── Planner ───────────────────────────────────────────────────────────────
  if (toInclude.has('planner')) {
    try {
      const tasks = await graphGet(
        graphClient.fetch,
        `/users/${encodeURIComponent(graphClient.userId)}/planner/tasks`,
        {
          $select: 'title,dueDateTime,percentComplete,bucketId,planId',
          $top: '100',
        }
      );
      data.planner = (tasks ?? [])
        .filter(t => t.percentComplete < 100)
        .map(t => ({
          title: t.title,
          dueDateTime: t.dueDateTime ?? null,
          percentComplete: t.percentComplete,
        }));
    } catch (err) {
      errors.push({ source: 'planner', message: err.message });
      data.planner = [];
    }
  }

  // ── Email ─────────────────────────────────────────────────────────────────
  if (toInclude.has('email')) {
    try {
      const SELECT = 'subject,from,receivedDateTime,isRead,flag,bodyPreview';
      const [unread, flagged] = await Promise.all([
        graphGet(graphClient.fetch, `/users/${encodeURIComponent(graphClient.userId)}/messages`, {
          $filter: 'isRead eq false',
          $select: SELECT,
          $top: '15',
        }),
        graphGet(graphClient.fetch, `/users/${encodeURIComponent(graphClient.userId)}/messages`, {
          $filter: "flag/flagStatus eq 'flagged'",
          $select: SELECT,
          $top: '10',
        }),
      ]);
      const seen = new Set();
      const all = [...(unread ?? []), ...(flagged ?? [])].filter(m => {
        if (seen.has(m.id)) return false;
        seen.add(m.id);
        return true;
      });
      data.email = all.map(m => ({
        subject: m.subject,
        from: m.from?.emailAddress?.name ?? m.from?.emailAddress?.address,
        receivedDateTime: m.receivedDateTime,
        isRead: m.isRead,
        flagged: m.flag?.flagStatus === 'flagged',
        preview: (m.bodyPreview ?? '').slice(0, 150),
      }));
    } catch (err) {
      errors.push({ source: 'email', message: err.message });
      data.email = [];
    }
  }

  // ── Chats ─────────────────────────────────────────────────────────────────
  if (toInclude.has('chats')) {
    try {
      const chats = await graphGet(
        graphClient.fetch,
        `/users/${encodeURIComponent(graphClient.userId)}/chats`,
        {
          $expand: 'lastMessagePreview',
          $select: 'id,topic,chatType,lastMessagePreview',
          $top: '20',
        }
      );

      const since = new Date(Date.now() - 24 * 60 * 60 * 1000);
      data.chats = (chats ?? [])
        .filter(c => {
          const lastAt = c.lastMessagePreview?.createdDateTime;
          return lastAt && new Date(lastAt) >= since;
        })
        .map(c => ({
          topic: c.topic ?? (c.chatType === 'oneOnOne' ? 'Direct Message' : 'Group Chat'),
          type: c.chatType,
          lastMessageAt: c.lastMessagePreview?.createdDateTime ?? null,
          lastMessageFrom: c.lastMessagePreview?.from?.user?.displayName ?? null,
          preview: (c.lastMessagePreview?.body?.content ?? '').replace(/<[^>]+>/g, '').slice(0, 150),
          isRead: c.lastMessagePreview?.isRead ?? true,
        }))
        .slice(0, 10);
    } catch (err) {
      errors.push({ source: 'chats', message: err.message });
      data.chats = [];
    }
  }

  // ── Files (recent — requires drive.recent or search) ──────────────────────
  if (toInclude.has('files')) {
    // drive/recent not supported for app-only auth — skip silently
    data.recentFiles = [];
  }

  // ── Return format ─────────────────────────────────────────────────────────

  if (bkgFormat) {
    // BKG heartbeat-compatible format: { source, timestamp, data, errors }
    return {
      source,
      timestamp: new Date().toISOString(),
      data,
      errors,
    };
  }

  // MCP tool format: compact success response
  return {
    success: true,
    timestamp: new Date().toISOString(),
    included: include,
    errors: errors.length ? errors : undefined,
    ...data,
  };
}

// ── MCP tool definitions ──────────────────────────────────────────────────────

export const SNAPSHOT_TOOL_DEFS = [
  {
    name: 'm365_snapshot',
    description: 'Full M365 context dump: calendar (today), unread email, todo tasks, planner tasks, recent chats. Use for session boot or when you need a complete picture. For targeted lookups, use the individual tools instead.',
    inputSchema: {
      type: 'object',
      properties: {
        include: {
          type: 'array',
          items: { type: 'string', enum: ['calendar', 'todo', 'planner', 'email', 'chats'] },
          description: 'Services to include (default: all). Omit services you don\'t need to reduce token cost.',
        },
        calendarHours: { type: 'number', description: 'Hours ahead for calendar window (default: end of day)' },
      },
    },
  },
];
