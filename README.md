# mcp-server-m365

A standalone MCP server for Microsoft 365 via Graph API. Works with Claude Code, Cursor, Windsurf, or any MCP client.

**26 tools:** To-Do (6), Calendar (6), Planner read (4), Email (5), Teams Chats read (2), Files search (1), Snapshot (1)

---

## Prerequisites

- **Node.js 18+** — [nodejs.org](https://nodejs.org)
- A **Microsoft 365 account** (work/school or personal with M365 subscription)
- An **Entra ID (Azure AD) App Registration** — see setup below
- Admin consent on the app registration (or a Global Admin who can grant it)

---

## Step 1: Node.js dependencies

```bash
git clone https://github.com/Backline-Strategic/mcp-server-m365.git
cd mcp-server-m365
npm install
```

This installs two packages:

| Package | Version | Purpose |
|---------|---------|---------|
| `@azure/msal-node` | ^2.16.3 | MSAL client credentials flow — acquires Graph API tokens |
| `node-fetch` | ^3.3.2 | HTTP client for Graph API calls |

No other runtime dependencies. No build step.

---

## Step 2: Entra App Registration

This server uses **app-only (client credentials) auth** — it runs without a signed-in user. That means you need an Entra app registration with application permissions (not delegated) and admin consent.

### Option A: Create it yourself (portal)

1. Go to [portal.azure.com](https://portal.azure.com)
2. Search for **Microsoft Entra ID** (formerly Azure Active Directory)
3. Go to **App registrations** → **New registration**
4. Fill in:
   - **Name:** `mcp-server-m365` (or anything you like)
   - **Supported account types:** Accounts in this organizational directory only (single tenant)
   - **Redirect URI:** leave blank
5. Click **Register**
6. Note the **Application (client) ID** and **Directory (tenant) ID** from the Overview page

### Option B: Ask your IT admin

Send them this:

> "I need an Entra app registration for a local MCP tool that reads my M365 data (calendar, email, To-Do, Teams chats, OneDrive files) using app-only auth. The app needs the Graph application permissions listed below, with admin consent granted. Please share the tenant ID, client ID, and a client secret."

---

## Step 3: API permissions

In the app registration, go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**.

Add all of the following:

| Permission | Type | Used by | Notes |
|------------|------|---------|-------|
| `Calendars.ReadWrite` | Application | calendar tools | Read + create/update/delete events |
| `Tasks.ReadWrite.All` | Application | todo tools | Read + write To-Do tasks and lists |
| `Tasks.Read.All` | Application | planner tools | Read Planner tasks, plans, buckets |
| `Mail.ReadWrite` | Application | email tools | Read, flag, move, draft emails |
| `Chat.Read.All` | Application | chat tools | Read Teams chats and messages |
| `Files.Read.All` | Application | files_search | Search OneDrive files |
| `GroupMember.Read.All` | Application | planner_list_plans | Resolve which Planner plans user can access |
| `User.Read.All` | Application | all tools | Resolve user by UPN/email address |

After adding all permissions, click **Grant admin consent for [your tenant]** and confirm. All permissions should show a green checkmark.

> **Why application permissions?** This server runs as a background process without a browser session. Delegated permissions require an interactive login flow. Application permissions + a client secret let the server authenticate silently.

> **`Mail.ReadWrite` vs `Mail.Read`:** `Mail.Read` is sufficient if you only need `email_search`, `email_get`, and `email_list_unread`. You need `Mail.ReadWrite` for `email_draft_create` and `email_update` (flagging, moving).

> **`Tasks.ReadWrite.All` vs `Tasks.Read.All`:** Both permissions cover both To-Do and Planner read access. `Tasks.ReadWrite.All` is required for any To-Do write operation (create, update, delete). Planner write is not currently implemented in this server.

---

## Step 4: Client secret

1. In the app registration, go to **Certificates & secrets** → **Client secrets** → **New client secret**
2. Set a description (e.g. `mcp-server-m365`) and expiry (24 months recommended)
3. Click **Add**
4. **Copy the secret value immediately** — you cannot view it again after leaving this page

---

## Step 5: Find your userId

The `userId` in your config is the UPN (User Principal Name) of the account whose data you want to access. This is typically your work email address: `you@yourdomain.com`.

To confirm: go to [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **Users** → find your account → copy the **User principal name**.

---

## Step 6: Configure credentials

### Option A — Config file (recommended for multi-account)

```bash
mkdir -p ~/.mcp-server-m365
cp m365-accounts.example.json ~/.mcp-server-m365/accounts.json
```

Edit `~/.mcp-server-m365/accounts.json`:

```json
{
  "default": {
    "tenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "clientId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "clientSecret": "your-client-secret-value",
    "userId": "you@yourdomain.com"
  }
}
```

For multiple accounts (e.g. personal + work):

```json
{
  "personal": {
    "tenantId": "...",
    "clientId": "...",
    "clientSecret": "...",
    "userId": "you@personal.com"
  },
  "work": {
    "tenantId": "...",
    "clientId": "...",
    "clientSecret": "...",
    "userId": "you@company.com"
  }
}
```

Pass `"account": "work"` in tool arguments to target a specific account.

### Option B — Environment variables (single account)

```bash
export M365_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
export M365_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
export M365_CLIENT_SECRET=your-client-secret-value
export M365_USER_ID=you@yourdomain.com
```

---

## Step 7: Register in `.mcp.json`

### Config file auth

```json
{
  "mcp-server-m365": {
    "command": "node",
    "args": ["/absolute/path/to/mcp-server-m365/src/index.mjs"],
    "env": {
      "M365_ACCOUNTS_FILE": "/Users/you/.mcp-server-m365/accounts.json"
    }
  }
}
```

### Env var auth

```json
{
  "mcp-server-m365": {
    "command": "node",
    "args": ["/absolute/path/to/mcp-server-m365/src/index.mjs"],
    "env": {
      "M365_TENANT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "M365_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "M365_CLIENT_SECRET": "your-secret",
      "M365_USER_ID": "you@yourdomain.com"
    }
  }
}
```

> `.mcp.json` lives at the root of your project, or use `~/.claude/mcp.json` for Claude Code global config.

---

## Testing

```bash
# List all tools (no credentials needed)
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | node src/index.mjs 2>/dev/null | node -e "const d=require('fs').readFileSync('/dev/stdin','utf8'); console.log('Tools:', JSON.parse(d).result.tools.length)"

# List To-Do tasks (requires credentials)
echo '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"todo_list_tasks","arguments":{}}}' | node src/index.mjs

# Create a task
echo '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"todo_create_task","arguments":{"title":"Test from MCP"}}}' | node src/index.mjs

# List today's calendar events
echo '{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"calendar_list_events","arguments":{}}}' | node src/index.mjs

# Full snapshot (calendar + todo + email + planner + chats)
echo '{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"m365_snapshot","arguments":{}}}' | node src/index.mjs

# Targeted snapshot (just calendar and todo)
echo '{"jsonrpc":"2.0","id":6,"method":"tools/call","params":{"name":"m365_snapshot","arguments":{"include":["calendar","todo"]}}}' | node src/index.mjs
```

---

## Multi-account usage

Pass `"account"` to any tool to target a specific account from your config file:

```json
{ "name": "todo_list_tasks", "arguments": { "account": "work" } }
{ "name": "calendar_list_events", "arguments": { "account": "personal", "startDate": "2026-03-20" } }
```

Omit `account` to use the `"default"` entry, or when using env vars.

---

## Tool reference

### To-Do
| Tool | Required | Optional |
|------|----------|----------|
| `todo_list_tasks` | — | `listName`, `status`, `limit` |
| `todo_get_task` | `taskId` OR `title` | `listName` |
| `todo_create_task` | `title` OR `titles` | `dueDate`, `importance`, `listName` |
| `todo_update_task` | `taskId` OR `title` | `status`, `newTitle`, `dueDate`, `importance`, `listName` |
| `todo_delete_task` | `taskId` OR `title` | `listName` |
| `todo_list_lists` | — | — |

### Calendar
| Tool | Required | Optional |
|------|----------|----------|
| `calendar_list_events` | — | `startDate`, `endDate`, `timeZone`, `limit` |
| `calendar_get_event` | `eventId` | — |
| `calendar_search_events` | `query` | `startDate`, `limit` |
| `calendar_create_event` | `subject`, `start`, `end` | `timeZone`, `location`, `body`, `isAllDay`, `isOnlineMeeting`, `attendees` |
| `calendar_update_event` | `eventId` | `subject`, `start`, `end`, `timeZone`, `location`, `body`, `attendees` |
| `calendar_delete_event` | `eventId` | — |

### Planner (read-only)
| Tool | Required | Optional |
|------|----------|----------|
| `planner_list_tasks` | — | `planId`, `bucketId`, `limit` |
| `planner_get_task` | `taskId` | — |
| `planner_list_plans` | — | `limit` |
| `planner_list_buckets` | `planId` | — |

### Email
| Tool | Required | Optional |
|------|----------|----------|
| `email_search` | `query` | `limit`, `folder` |
| `email_get` | `messageId` | `includeBody` |
| `email_list_unread` | — | `limit`, `folder` |
| `email_draft_create` | `subject`, `toRecipients` | `body`, `ccRecipients`, `importance` |
| `email_update` | `messageId` | `isRead`, `flag`, `destinationFolder` |

### Chats (read-only)
| Tool | Required | Optional |
|------|----------|----------|
| `chat_list` | — | `limit`, `since` |
| `chat_get_messages` | `chatId` | `limit`, `since` |

### Files
| Tool | Required | Optional |
|------|----------|----------|
| `files_search` | `query` | `limit` |

### Snapshot
| Tool | Required | Optional |
|------|----------|----------|
| `m365_snapshot` | — | `include` (array), `calendarHours` |

`include` values: `"calendar"`, `"todo"`, `"planner"`, `"email"`, `"chats"`

---

## Troubleshooting

**`No Keychain entry found`** — This server does not use macOS Keychain. Make sure `M365_ACCOUNTS_FILE` points to a valid JSON file, or set the four `M365_*` env vars.

**`403 Forbidden` on any tool** — Admin consent has not been granted, or the permission is missing. Go to the app registration → API permissions → confirm all permissions show a green "Granted" checkmark.

**`401 Unauthorized`** — Check `tenantId`, `clientId`, and `clientSecret`. Client secrets expire — generate a new one if it's past its expiry date.

**`404 Not Found` on planner tools** — The user may not have any Planner plans, or `GroupMember.Read.All` is missing.

**`chat_list` returns empty** — `Chat.Read.All` requires admin consent in most tenants. Check with your IT admin.

**`files_search` returns empty** — `drive/recent` is not supported for app-only auth (this is a Graph API limitation). `files_search` uses `drive/root/search` which does work, but the user's OneDrive must be provisioned.

---

## Security

- **Never commit credentials.** Use `~/.mcp-server-m365/accounts.json` outside your project directory, or use env vars injected at runtime.
- Client secrets are sensitive — treat them like passwords. Rotate them annually or sooner.
- This server uses app-only auth, meaning it can access the configured `userId`'s data without their active session. Limit the app registration to only the users/data it needs.
- Consider using **certificate-based auth** instead of client secrets for production or shared environments (not yet implemented in this server).

---

## .gitignore

If you fork or extend this project, ensure the following are in your `.gitignore`:

```gitignore
# Credentials — never commit these
m365-accounts.json
*.credentials.json
.env
.env.*
!.env.example

# Node
node_modules/
npm-debug.log*
yarn-debug.log*
yarn-error.log*

# macOS
.DS_Store
.AppleDouble
.LSOverride

# Editor
.vscode/
.idea/
*.swp
*.swo

# Runtime
*.log
```

---

## License

MIT
