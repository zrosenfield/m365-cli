# sp-cli

SharePoint Online CLI for [openclaw](https://github.com/openclaw) agents. Exposes SharePoint file and list operations as a structured `sp` command that AI agents can call programmatically.

## Install

```bash
npm install -g @openclaw/sp-cli
```

Or from source:

```bash
git clone <repo>
cd sp-cli
npm install
npm run build
npm link
```

Requires Node.js ≥ 18.

## Quickstart

### 1. Configure authentication

These values come from an Azure AD app registration — see [App Registration](#app-registration) below for how to create one and which permission model to choose.

**Delegated / device code (recommended — acts as your user account):**
```bash
sp auth setup \
  --tenant-id "<Directory (tenant) ID from Azure AD>" \
  --client-id "<Application (client) ID from Azure AD>"
sp auth login    # Prints a code; open the URL in a browser to complete sign-in
```
The app can only access what your account can access. Tokens expire; re-run `sp auth login` when they do.

**Service principal / client credentials (headless agents, broad access):**
```bash
sp auth setup \
  --tenant-id "<Directory (tenant) ID from Azure AD>" \
  --client-id "<Application (client) ID from Azure AD>" \
  --client-secret "<client secret value>" \
  --tenant-url "https://contoso.sharepoint.com"
```
No interactive login needed, but requires admin-consented application permissions (see App Registration).

**Raw access token (escape hatch):**
```bash
export SP_CLI_ACCESS_TOKEN="eyJ..."
```

### 2. Discover your site and drive

```bash
sp sites list | jq '.data[] | {id, displayName}'
sp drives list --site <site-id> | jq '.data[] | {id, name}'
sp config set --site <site-id> --drive <drive-id>
```

### 3. Use it

```bash
# Files
sp files list
sp files upload ./report.xlsx --remote-path /Documents/report.xlsx
sp files download <item-id> --output ./report.xlsx
sp files search "quarterly report"

# Lists
sp lists list
sp lists items create <list-id> --fields '{"Title":"Task 1"}'
sp lists items list <list-id>
```

## Command Reference

| Area | Commands |
|---|---|
| Auth | `sp auth setup\|login\|logout\|token` |
| Config | `sp config get\|set` |
| Sites | `sp sites list\|get` |
| Drives | `sp drives list\|get` |
| Files | `sp files list\|get\|upload\|download\|copy\|move\|rename\|delete\|search` |
| Lists | `sp lists list\|get\|create\|update\|delete` |
| List items | `sp lists items list\|get\|create\|update\|delete` |
| Columns | `sp lists columns list\|get\|add\|update\|remove` |
| Permissions | `sp permissions list\|get\|grant\|update\|revoke` |

Run `sp <command> --help` for detailed options.

## Output Format

All commands output JSON to stdout:
```json
{ "data": <result> }
```

Errors go to stderr:
```json
{ "error": { "code": "...", "message": "...", "status": 404 } }
```

## For Agents

See [`CLAUDE.md`](CLAUDE.md) for the entry point, and the skill files under `skills/` for full command documentation:

- [`skills/sp-shared/SKILL.md`](skills/sp-shared/SKILL.md) — Auth & config (read first)
- [`skills/sp-files/SKILL.md`](skills/sp-files/SKILL.md) — File operations
- [`skills/sp-lists/SKILL.md`](skills/sp-lists/SKILL.md) — List operations

## App Registration

sp-cli authenticates through an **Azure AD app registration** — an identity that represents the CLI in your Microsoft 365 tenant. You need to create one before running `sp auth setup`.

**Step 1: Create the app registration**

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Give it a name (e.g. "sp-cli"), leave defaults, click **Register**
3. On the overview page, copy:
   - **Application (client) ID** → this is your `--client-id`
   - **Directory (tenant) ID** → this is your `--tenant-id`

**Step 2: Choose a permission model**

There are three options, ordered from least to most access:

| Model | Access scope | Admin consent | Best for |
|---|---|---|---|
| **Delegated** (device code) | Whatever your user account can access | Not required | Personal use, least privilege |
| **`Sites.Selected`** (app-only) | Specific sites you explicitly grant | Required for the permission itself; SharePoint admin grants per-site | Headless agents with controlled blast radius |
| **`Sites.ReadWrite.All`** (app-only) | Every site in the tenant | Required (Global/SharePoint Admin) | Tenant-wide automation |

**Delegated permissions (recommended starting point):**
1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Add: `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `offline_access`
   - Note: `Lists.ReadWrite.All` does not exist as a delegated permission — list access is covered by `Sites.ReadWrite.All`
3. No admin consent needed — you consent for yourself when you run `sp auth login`
4. No client secret needed
5. Go to **Authentication** → **Advanced settings** → set **Allow public client flows** to **Yes** and save — required for device code login

**`Sites.Selected` — app-only, scoped to specific sites:**
1. Go to **API permissions** → **Microsoft Graph** → **Application permissions** → add `Sites.Selected`
2. Click **Grant admin consent** (unlocks the permission but grants no site access yet)
3. A SharePoint admin must then explicitly grant the app access to each site:
   ```
   POST https://graph.microsoft.com/v1.0/sites/{siteId}/permissions
   { "roles": ["write"], "grantedToIdentities": [{ "application": { "id": "<client-id>", "displayName": "sp-cli" } }] }
   ```
4. Create a client secret (Certificates & secrets) and use `sp auth setup --client-secret`

**`Sites.ReadWrite.All` — app-only, full tenant:**
1. Go to **API permissions** → **Microsoft Graph** → **Application permissions**
2. Add `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Lists.ReadWrite.All`
3. Click **Grant admin consent** (requires Global Admin or SharePoint Admin)
4. Create a client secret and use `sp auth setup --client-secret`

After completing setup, run `sp auth setup` with the values you copied.

## License

MIT
