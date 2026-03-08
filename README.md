# m365-cli

[![npm version](https://img.shields.io/npm/v/m365-cli.svg)](https://www.npmjs.com/package/m365-cli)
[![CI](https://github.com/zrosenfield/m365-cli/actions/workflows/ci.yml/badge.svg)](https://github.com/zrosenfield/m365-cli/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

> **Disclaimer:** This is not an officially supported Microsoft product and is not affiliated with, endorsed by, or sponsored by Microsoft Corporation.

**Give your AI assistant access to your Microsoft 365 files, lists, mail, and calendar.**

m365-cli is a command-line tool designed for AI agents (like Claude). You install it once and configure it — then your AI can read and write your SharePoint files, manage lists, send mail, and work with your calendar on your behalf. You don't need to type these commands yourself; your AI does.

## Prerequisites

- [Node.js](https://nodejs.org/) version 18 or later
- A Microsoft 365 account (work, school, or personal with M365 subscription)
- An Azure AD app registration (see [App Registration](#app-registration) below — it's a one-time setup)

## Install

```bash
npm install -g m365-cli
```

Requires Node.js ≥ 18.

## Getting Started

### Step 1: Create an Azure AD app registration

Before you can authenticate, you need to register an app in Azure. This is a one-time step that gives m365-cli an identity in your Microsoft tenant.

See [App Registration](#app-registration) below. If you're just getting started for personal use, use the **Delegated** option — it requires no admin approval and accesses only what your own account can see.

### Step 2: Configure authentication

**Delegated / device code (recommended for personal use):**
```bash
m365 auth setup \
  --tenant-id "<Directory (tenant) ID from Azure AD>" \
  --client-id "<Application (client) ID from Azure AD>"
m365 auth login    # Prints a code; open the URL in a browser to complete sign-in
```
This acts as your user account. Tokens expire; re-run `m365 auth login` when they do.

**Service principal / client credentials (for headless agents or automation):**
```bash
export SP_CLI_CLIENT_SECRET="<your client secret value>"
m365 auth setup \
  --tenant-id "<Directory (tenant) ID>" \
  --client-id "<Application (client) ID>" \
  --tenant-url "https://contoso.sharepoint.com"
```
No interactive login needed. Requires admin-consented application permissions (see App Registration).

**Raw access token (escape hatch):**
```bash
export SP_CLI_ACCESS_TOKEN="eyJ..."
```

### Step 3: Find your site and drive

```bash
m365 sites list | jq '.data[] | {id, displayName}'
m365 drives list --site <site-id> | jq '.data[] | {id, name}'
m365 config set --site <site-id> --drive <drive-id>
```

### Step 4: Try it out

```bash
# Files
m365 files list
m365 files upload ./report.xlsx --remote-path /Documents/report.xlsx
m365 files search "quarterly report"

# Lists
m365 lists list
m365 lists items list <list-id>

# Mail
m365 mail list --top 5
m365 mail send --to user@example.com --subject "Hello" --body "Hi there"

# Calendar
m365 calendar events list --top 5
```

## Command Reference

| Area | Commands |
|---|---|
| Auth | `m365 auth setup\|login\|logout\|token` |
| Config | `m365 config get\|set` |
| Sites | `m365 sites list\|get` |
| Drives | `m365 drives list\|get` |
| Files | `m365 files list\|get\|upload\|download\|copy\|move\|rename\|delete\|search` |
| Lists | `m365 lists list\|get\|create\|update\|delete` |
| List items | `m365 lists items list\|get\|create\|update\|delete` |
| Columns | `m365 lists columns list\|get\|add\|update\|remove` |
| Permissions | `m365 permissions list\|get\|grant\|update\|revoke` |
| Mail | `m365 mail list\|get\|send\|reply\|delete\|folders` |
| Calendar | `m365 calendar list\|get\|events list\|get\|create\|update\|delete\|rsvp` |

Run `m365 <command> --help` for detailed options.

## Output Format

All commands output JSON to stdout:
```json
{ "data": <result> }
```

Errors go to stderr:
```json
{ "error": { "code": "...", "message": "...", "status": 404 } }
```

This structured output is what allows AI agents to parse and act on results reliably.

## For AI Agents

See [`CLAUDE.md`](CLAUDE.md) for the entry point. Full command documentation is in the skill files under `skills/`:

- [`skills/m365-shared/SKILL.md`](skills/m365-shared/SKILL.md) — Auth & config (read first)
- [`skills/m365-files/SKILL.md`](skills/m365-files/SKILL.md) — File operations
- [`skills/m365-lists/SKILL.md`](skills/m365-lists/SKILL.md) — List operations
- [`skills/m365-mail/SKILL.md`](skills/m365-mail/SKILL.md) — Mail operations
- [`skills/m365-calendar/SKILL.md`](skills/m365-calendar/SKILL.md) — Calendar operations

## App Registration

m365-cli authenticates through an **Azure AD app registration** — a one-time identity you create in the Azure portal that represents the CLI in your Microsoft 365 tenant.

### Which option should I use?

| | Delegated (device code) | Sites.Selected (app-only) | Sites.ReadWrite.All (app-only) |
|---|---|---|---|
| **Access scope** | Whatever your user account can see | Only specific sites you grant | Every site in the tenant |
| **Admin approval needed?** | No | Yes (SharePoint Admin) | Yes (Global or SharePoint Admin) |
| **Best for** | Personal use, getting started | Agents with controlled access | Tenant-wide automation |

**If you're not sure, start with Delegated.**

---

### Create the app registration (all options)

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Give it a name (e.g. "m365-cli"), leave defaults, click **Register**
3. On the overview page, copy:
   - **Application (client) ID** → your `--client-id`
   - **Directory (tenant) ID** → your `--tenant-id`

---

### Delegated permissions (recommended starting point)

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Add: `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `offline_access`
3. No admin consent needed — you consent when you run `m365 auth login`
4. Go to **Authentication** → **Advanced settings** → set **Allow public client flows** to **Yes** and save (required for device code login)
5. No client secret needed

---

### Sites.Selected — app-only, scoped to specific sites

1. Go to **API permissions** → **Microsoft Graph** → **Application permissions** → add `Sites.Selected`
2. Click **Grant admin consent**
3. A SharePoint admin must then grant the app access to each site:
   ```
   POST https://graph.microsoft.com/v1.0/sites/{siteId}/permissions
   { "roles": ["write"], "grantedToIdentities": [{ "application": { "id": "<client-id>", "displayName": "m365-cli" } }] }
   ```
4. Create a client secret under **Certificates & secrets**, copy the value, and set `SP_CLI_CLIENT_SECRET` before running `m365 auth setup`

---

### Sites.ReadWrite.All — app-only, full tenant

1. Go to **API permissions** → **Microsoft Graph** → **Application permissions**
2. Add `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Lists.ReadWrite.All`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`
3. Click **Grant admin consent** (requires Global Admin or SharePoint Admin)
4. Create a client secret and set `SP_CLI_CLIENT_SECRET` before running `m365 auth setup`

---

## Building from Source

```bash
git clone https://github.com/zrosenfield/m365-cli
cd m365-cli
npm install
npm run build    # compiles TypeScript → dist/
npm link         # makes `m365` available globally
```

## License

Apache 2.0 — see [LICENSE](LICENSE).
