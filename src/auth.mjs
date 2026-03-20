/**
 * Credential loading for mcp-server-m365.
 *
 * Priority order:
 *   1. M365_ACCOUNTS_FILE env var → JSON config file
 *   2. ~/.mcp-server-m365/accounts.json (default path)
 *   3. Individual env vars: M365_TENANT_ID, M365_CLIENT_ID, M365_CLIENT_SECRET, M365_USER_ID
 *
 * Config file format: see m365-accounts.example.json
 */

import { readFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';

const DEFAULT_CONFIG_PATH = join(homedir(), '.mcp-server-m365', 'accounts.json');

/**
 * Load credentials for a named account.
 * @param {string} accountName - Key in m365-accounts.json (default: 'default')
 * @returns {Promise<{tenantId: string, clientId: string, clientSecret: string, userId: string}>}
 */
export async function getCredentials(accountName = 'default') {
  const configPath = process.env.M365_ACCOUNTS_FILE ?? DEFAULT_CONFIG_PATH;

  let accounts;
  try {
    const raw = await readFile(configPath, 'utf8');
    accounts = JSON.parse(raw);
  } catch (err) {
    if (err.code === 'ENOENT') {
      return loadFromEnv();
    }
    throw new Error(`Failed to read accounts config at ${configPath}: ${err.message}`);
  }

  const creds = accounts[accountName];
  if (!creds) {
    const available = Object.keys(accounts).join(', ');
    throw new Error(
      `Account "${accountName}" not found in ${configPath}. Available: ${available}. ` +
      `To add it, copy m365-accounts.example.json and fill in your Azure AD app credentials.`
    );
  }

  validateCreds(creds, accountName, configPath);
  return creds;
}

function loadFromEnv() {
  const { M365_TENANT_ID, M365_CLIENT_ID, M365_CLIENT_SECRET, M365_USER_ID } = process.env;
  if (!M365_TENANT_ID || !M365_CLIENT_ID || !M365_CLIENT_SECRET || !M365_USER_ID) {
    throw new Error(
      'No accounts config found. Either:\n' +
      '  1. Set M365_ACCOUNTS_FILE=/path/to/m365-accounts.json\n' +
      '  2. Create ~/.mcp-server-m365/accounts.json\n' +
      '  3. Set env vars: M365_TENANT_ID, M365_CLIENT_ID, M365_CLIENT_SECRET, M365_USER_ID'
    );
  }
  return {
    tenantId: M365_TENANT_ID,
    clientId: M365_CLIENT_ID,
    clientSecret: M365_CLIENT_SECRET,
    userId: M365_USER_ID,
  };
}

function validateCreds(creds, accountName, source) {
  const required = ['tenantId', 'clientId', 'clientSecret', 'userId'];
  const missing = required.filter(k => !creds[k]);
  if (missing.length) {
    throw new Error(
      `Account "${accountName}" in ${source} is missing required fields: ${missing.join(', ')}`
    );
  }
}
