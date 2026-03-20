/**
 * Teams chat operations — read-only.
 * All functions accept graphClient: { fetch, write, userId }
 */

import { graphGet } from '../graph.mjs';
import { withErrorHandler, makeError } from '../lib/errors.mjs';
import { compact, preview } from '../lib/response.mjs';

/**
 * List recent chats with last message preview.
 */
export async function listChats(graphClient, { limit = 20, since } = {}) {
  return withErrorHandler('chat_list', async () => {
    const chats = await graphGet(
      graphClient.fetch,
      `/users/${encodeURIComponent(graphClient.userId)}/chats`,
      {
        $expand: 'lastMessagePreview',
        $select: 'id,topic,chatType,lastMessagePreview',
        $top: String(limit),
      }
    );

    const cutoff = since ? new Date(since) : new Date(Date.now() - 24 * 60 * 60 * 1000);

    const filtered = (chats ?? [])
      .filter(c => {
        const lastAt = c.lastMessagePreview?.createdDateTime;
        return lastAt && new Date(lastAt) >= cutoff;
      })
      .map(c => compact({
        id: c.id,
        topic: c.topic ?? (c.chatType === 'oneOnOne' ? 'Direct Message' : 'Group Chat'),
        type: c.chatType,
        lastMessageAt: c.lastMessagePreview?.createdDateTime,
        lastMessageFrom: c.lastMessagePreview?.from?.user?.displayName,
        preview: preview(c.lastMessagePreview?.body?.content?.replace(/<[^>]+>/g, '')),
        isRead: c.lastMessagePreview?.isRead ?? true,
      }));

    return { success: true, count: filtered.length, chats: filtered };
  });
}

/**
 * Get messages from a specific chat.
 */
export async function getChatMessages(graphClient, { chatId, limit = 20, since } = {}) {
  return withErrorHandler('chat_get_messages', async () => {
    if (!chatId) return makeError('chatId is required');

    const params = {
      $top: String(limit),
      $orderby: 'createdDateTime desc',
    };
    if (since) params.$filter = `createdDateTime ge ${since}`;

    const messages = await graphGet(
      graphClient.fetch,
      `/users/${encodeURIComponent(graphClient.userId)}/chats/${chatId}/messages`,
      params
    );

    return {
      success: true,
      chatId,
      count: (messages ?? []).length,
      messages: (messages ?? []).reverse().map(m => compact({
        id: m.id,
        from: m.from?.user?.displayName ?? m.from?.application?.displayName,
        createdAt: m.createdDateTime,
        body: preview(m.body?.content?.replace(/<[^>]+>/g, ''), 300),
        type: m.messageType,
      })),
    };
  });
}

// ── MCP tool definitions ──────────────────────────────────────────────────────

export const CHAT_TOOL_DEFS = [
  {
    name: 'chat_list',
    description: 'List recent Teams chats with last message preview. Defaults to last 24 hours.',
    inputSchema: {
      type: 'object',
      properties: {
        limit: { type: 'number', description: 'Max chats (default 20)' },
        since: { type: 'string', description: 'ISO datetime — only chats with activity after this' },
      },
    },
  },
  {
    name: 'chat_get_messages',
    description: 'Get messages from a specific Teams chat.',
    inputSchema: {
      type: 'object',
      properties: {
        chatId: { type: 'string' },
        limit: { type: 'number', description: 'Max messages (default 20)' },
        since: { type: 'string', description: 'ISO datetime — only messages after this' },
      },
      required: ['chatId'],
    },
  },
];
