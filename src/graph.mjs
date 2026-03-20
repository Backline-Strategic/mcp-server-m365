/**
 * Microsoft Graph API client.
 * Handles MSAL token acquisition, authenticated fetch/write, and auto-pagination.
 *
 * Tool functions receive a graphClient object: { fetch, write, userId }
 * This is the only interface they need — auth source is irrelevant to them.
 */

import { getCredentials } from './auth.mjs';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

async function acquireToken({ tenantId, clientId, clientSecret }) {
  const { ConfidentialClientApplication } = await import('@azure/msal-node');
  const app = new ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
  });
  const result = await app.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });
  if (!result?.accessToken) throw new Error('MSAL returned no access token');
  return result.accessToken;
}

/**
 * Build an authenticated Graph client for a named account (loaded from config/env).
 * @param {string} accountName - Account name from m365-accounts.json
 */
export async function getGraphClient(accountName = 'default') {
  const creds = await getCredentials(accountName);
  return buildClient(creds);
}

/**
 * Build an authenticated Graph client from a credentials object directly.
 * Used when the caller manages auth (e.g. BKG uses macOS Keychain).
 * @param {{tenantId: string, clientId: string, clientSecret: string, userId: string}} creds
 */
export async function getGraphClientFromCreds(creds) {
  return buildClient(creds);
}

async function buildClient({ tenantId, clientId, clientSecret, userId }) {
  const token = await acquireToken({ tenantId, clientId, clientSecret });

  async function graphFetch(path, params = {}, optionsOrRawText = false) {
    const { default: fetch } = await import('node-fetch');

    let rawText = false;
    let extraHeaders = {};
    if (typeof optionsOrRawText === 'boolean') {
      rawText = optionsOrRawText;
    } else if (optionsOrRawText && typeof optionsOrRawText === 'object') {
      rawText = optionsOrRawText.rawText ?? false;
      extraHeaders = optionsOrRawText.headers ?? {};
    }

    const qs = new URLSearchParams();
    for (const [k, v] of Object.entries(params)) qs.set(k, v);
    const url = `${GRAPH_BASE}${path}${qs.toString() ? '?' + qs.toString() : ''}`;

    const resp = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...extraHeaders,
      },
    });

    if (!resp.ok) {
      const body = await resp.text();
      throw new Error(`Graph ${resp.status} ${resp.statusText}: ${body.slice(0, 200)}`);
    }
    return rawText ? resp.text() : resp.json();
  }

  graphFetch._headers = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };

  async function graphWrite(path, body, method = 'POST') {
    const { default: fetch } = await import('node-fetch');
    const url = `${GRAPH_BASE}${path}`;
    const resp = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: body != null ? JSON.stringify(body) : undefined,
    });
    if (!resp.ok) {
      const errBody = await resp.text();
      throw new Error(`Graph ${method} ${resp.status} ${resp.statusText}: ${errBody.slice(0, 300)}`);
    }
    if (resp.status === 204) return null;
    return resp.json();
  }

  return { fetch: graphFetch, write: graphWrite, userId };
}

/**
 * GET with auto-pagination (follows @odata.nextLink).
 * @param {Function} graphFetch - Authenticated fetch from getGraphClient
 * @param {string} path - Initial Graph path
 * @param {object} [params] - Query params
 * @param {object|number} [options] - { maxPages?, headers? } or maxPages number
 * @returns {Promise<any[]>} Combined value array
 */
export async function graphGet(graphFetch, path, params = {}, options = {}) {
  const maxPages = typeof options === 'number' ? options : (options.maxPages ?? 5);
  const headers = typeof options === 'object' ? (options.headers ?? {}) : {};

  const { default: fetch } = await import('node-fetch');
  const results = [];
  let data = await graphFetch(path, params, Object.keys(headers).length ? { headers } : false);

  if (data.value !== undefined) {
    results.push(...data.value);
    let page = 1;
    while (data['@odata.nextLink'] && page < maxPages) {
      const resp = await fetch(data['@odata.nextLink'], {
        headers: { ...graphFetch._headers, ...headers },
      });
      data = await resp.json();
      if (data.value) results.push(...data.value);
      page++;
    }
    return results;
  }

  return data;
}
