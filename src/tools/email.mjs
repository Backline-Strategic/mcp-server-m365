/**
 * Email operations via Microsoft Graph.
 * Note: email_send and email_delete are intentionally excluded.
 * All functions accept graphClient: { fetch, write, userId }
 */

import { graphGet } from '../graph.mjs';
import { withErrorHandler, makeError } from '../lib/errors.mjs';
import { compact, preview } from '../lib/response.mjs';

const EMAIL_SELECT = 'id,subject,from,receivedDateTime,isRead,flag,bodyPreview';

function formatMessage(m, includeBody = false) {
  return compact({
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.name
      ? `${m.from.emailAddress.name} <${m.from.emailAddress.address}>`
      : m.from?.emailAddress?.address,
    received: m.receivedDateTime?.slice(0, 16),
    isRead: m.isRead,
    flagged: m.flag?.flagStatus === 'flagged' ? true : undefined,
    preview: preview(m.bodyPreview),
    body: includeBody ? m.body?.content : undefined,
    attachments: m.hasAttachments ? true : undefined,
  });
}

/**
 * Search emails via KQL. Supports: "word", "from:addr", "subject:text"
 */
export async function searchEmail(graphClient, { query, limit = 10, folder } = {}) {
  return withErrorHandler('email_search', async () => {
    if (!query) return makeError('query is required');

    // Wrap plain terms in quotes; pass KQL operators raw
    const kql = /^(from:|subject:|hasattachment:|received:|to:)/i.test(query)
      ? query
      : `"${query.replace(/"/g, '')}"`;

    const params = {
      $search: kql,
      $select: EMAIL_SELECT,
      $top: String(Math.min(Math.max(1, limit), 20)),
    };

    const basePath = `/users/${encodeURIComponent(graphClient.userId)}`;
    const path = folder ? `${basePath}/mailFolders/${folder}/messages` : `${basePath}/messages`;

    const messages = await graphGet(graphClient.fetch, path, params);
    return {
      success: true,
      query,
      count: (messages ?? []).length,
      messages: (messages ?? []).map(m => formatMessage(m)),
    };
  });
}

/**
 * Get a single email with full body.
 */
export async function getEmail(graphClient, { messageId, includeBody = true }) {
  return withErrorHandler('email_get', async () => {
    if (!messageId) return makeError('messageId is required');

    const selectFields = includeBody
      ? `${EMAIL_SELECT},body,hasAttachments,toRecipients,ccRecipients`
      : EMAIL_SELECT;

    const m = await graphClient.fetch(
      `/users/${encodeURIComponent(graphClient.userId)}/messages/${messageId}`,
      { $select: selectFields }
    );

    return {
      success: true,
      message: compact({
        ...formatMessage(m, includeBody),
        to: m.toRecipients?.map(r => r.emailAddress?.address),
        cc: m.ccRecipients?.length ? m.ccRecipients.map(r => r.emailAddress?.address) : undefined,
      }),
    };
  });
}

/**
 * List unread and flagged emails (compact format).
 */
export async function listUnread(graphClient, { limit = 15, folder } = {}) {
  return withErrorHandler('email_list_unread', async () => {
    const basePath = `/users/${encodeURIComponent(graphClient.userId)}`;
    const msgPath = folder ? `${basePath}/mailFolders/${folder}/messages` : `${basePath}/messages`;

    const [unread, flagged] = await Promise.all([
      graphGet(graphClient.fetch, msgPath, {
        $filter: 'isRead eq false',
        $select: EMAIL_SELECT,
        $top: String(limit),
        $orderby: 'receivedDateTime desc',
      }),
      graphGet(graphClient.fetch, msgPath, {
        $filter: "flag/flagStatus eq 'flagged'",
        $select: EMAIL_SELECT,
        $top: '10',
        $orderby: 'receivedDateTime desc',
      }),
    ]);

    // Merge and deduplicate
    const seen = new Set();
    const all = [...(unread ?? []), ...(flagged ?? [])].filter(m => {
      if (seen.has(m.id)) return false;
      seen.add(m.id);
      return true;
    });

    return {
      success: true,
      count: all.length,
      messages: all.map(m => formatMessage(m)),
    };
  });
}

/**
 * Create a draft email (does not send).
 */
export async function createDraft(graphClient, {
  subject,
  toRecipients,
  body,
  ccRecipients,
  importance = 'normal',
} = {}) {
  return withErrorHandler('email_draft_create', async () => {
    if (!subject) return makeError('subject is required');
    if (!toRecipients?.length) return makeError('toRecipients is required (array of email strings)');

    const message = {
      subject,
      importance,
      toRecipients: toRecipients.map(addr => ({
        emailAddress: typeof addr === 'string' ? { address: addr } : addr,
      })),
    };

    if (body) message.body = { contentType: 'text', content: body };
    if (ccRecipients?.length) {
      message.ccRecipients = ccRecipients.map(addr => ({
        emailAddress: typeof addr === 'string' ? { address: addr } : addr,
      }));
    }

    const draft = await graphClient.write(
      `/users/${encodeURIComponent(graphClient.userId)}/messages`,
      message
    );

    return {
      success: true,
      draft: compact({
        id: draft.id,
        subject: draft.subject,
        to: toRecipients,
        isDraft: draft.isDraft,
        webLink: draft.webLink,
      }),
    };
  });
}

/**
 * Update email: flag/unflag, mark read/unread, move to folder.
 */
export async function updateEmail(graphClient, {
  messageId,
  isRead,
  flag,
  destinationFolder,
} = {}) {
  return withErrorHandler('email_update', async () => {
    if (!messageId) return makeError('messageId is required');

    const basePath = `/users/${encodeURIComponent(graphClient.userId)}`;

    // Move to folder first if requested
    if (destinationFolder) {
      await graphClient.write(
        `${basePath}/messages/${messageId}/move`,
        { destinationId: destinationFolder }
      );
    }

    const patch = {};
    if (isRead != null) patch.isRead = isRead;
    if (flag != null) patch.flag = { flagStatus: flag ? 'flagged' : 'notFlagged' };

    if (Object.keys(patch).length > 0) {
      await graphClient.write(`${basePath}/messages/${messageId}`, patch, 'PATCH');
    }

    return {
      success: true,
      updated: compact({ messageId, isRead, flagged: flag, movedTo: destinationFolder }),
    };
  });
}

// ── MCP tool definitions ──────────────────────────────────────────────────────

export const EMAIL_TOOL_DEFS = [
  {
    name: 'email_search',
    description: 'Search emails via KQL. Supports: "word", "from:addr@domain.com", "subject:invoice", "hasattachment:true". Returns compact results ordered newest first.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'KQL expression or plain search term' },
        limit: { type: 'number', description: 'Max results (default 10, max 20)' },
        folder: { type: 'string', description: 'Folder ID or well-known name (inbox, sentitems, archive)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'email_get',
    description: 'Get a single email with full body by message ID.',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string' },
        includeBody: { type: 'boolean', description: 'Include full body (default true)' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'email_list_unread',
    description: 'List unread and flagged emails (compact format, no body).',
    inputSchema: {
      type: 'object',
      properties: {
        limit: { type: 'number', description: 'Max unread messages (default 15)' },
        folder: { type: 'string', description: 'Folder ID or well-known name' },
      },
    },
  },
  {
    name: 'email_draft_create',
    description: 'Create a draft email. Does NOT send — drafts are saved for review before sending.',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string' },
        toRecipients: { type: 'array', items: { type: 'string' }, description: 'Array of email addresses' },
        body: { type: 'string', description: 'Plain text body' },
        ccRecipients: { type: 'array', items: { type: 'string' } },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Default: normal' },
      },
      required: ['subject', 'toRecipients'],
    },
  },
  {
    name: 'email_update',
    description: 'Flag/unflag, mark read/unread, or move an email to another folder.',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string' },
        isRead: { type: 'boolean' },
        flag: { type: 'boolean', description: 'true to flag, false to unflag' },
        destinationFolder: { type: 'string', description: 'Folder ID or well-known name (inbox, archive, deleteditems)' },
      },
      required: ['messageId'],
    },
  },
];
