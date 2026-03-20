/**
 * OneDrive files — search only.
 * Note: drive/recent requires delegated auth and does not work with app-only.
 * drive/root/search works with app-only auth.
 * All functions accept graphClient: { fetch, write, userId }
 */

import { graphGet } from '../graph.mjs';
import { withErrorHandler, makeError } from '../lib/errors.mjs';
import { compact } from '../lib/response.mjs';

/**
 * Search OneDrive files by name or content keyword.
 */
export async function searchFiles(graphClient, { query, limit = 25 } = {}) {
  return withErrorHandler('files_search', async () => {
    if (!query) return makeError('query is required');

    const items = await graphGet(
      graphClient.fetch,
      `/users/${encodeURIComponent(graphClient.userId)}/drive/root/search(q='${encodeURIComponent(query)}')`,
      {
        $select: 'id,name,lastModifiedDateTime,webUrl,size,file,folder',
        $top: String(Math.min(limit, 50)),
      }
    );

    const files = (items ?? [])
      .filter(i => i.file) // files only, exclude folders
      .map(f => compact({
        id: f.id,
        name: f.name,
        modified: f.lastModifiedDateTime?.slice(0, 10),
        size: f.size,
        url: f.webUrl,
      }));

    return { success: true, query, count: files.length, files };
  });
}

// ── MCP tool definitions ──────────────────────────────────────────────────────

export const FILES_TOOL_DEFS = [
  {
    name: 'files_search',
    description: 'Search OneDrive files by name or content keyword.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search term — matches filename and content' },
        limit: { type: 'number', description: 'Max results (default 25, max 50)' },
      },
      required: ['query'],
    },
  },
];
