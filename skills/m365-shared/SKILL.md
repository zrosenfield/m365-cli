---
name: m365-shared
version: 0.1.0
description: "M365 CLI: Auth, config, global flags, and site/drive discovery. Prerequisite for all m365-* skills."
metadata:
  openclaw:
    category: "productivity"
    requires:
      bins: ["m365"]
    cliHelp: "m365 --help"
---

# m365-shared — Auth, Config & Global Concepts

## Prerequisites

Install the CLI:
```bash
npm install -g github:zrosenfield/m365-cli
# or from source:
git clone https://github.com/zrosenfield/m365-cli && cd m365-cli && npm install && npm run build && npm link
```

Verify: `m365 --version`

---

## Authentication

m365 supports three auth methods, tried in this order on every command:

1. `SP_CLI_ACCESS_TOKEN` env var
2. Client credentials (service principal) — if `clientSecret` is configured
3. Stored delegated token — set by `m365 auth login`

### Permission models — choose one

**Delegated (device code) — recommended, least privilege**

The app acts as your signed-in user account. It can only access sites your account can access. No admin consent required.

```bash
# .env (no SP_CLI_CLIENT_SECRET)
SP_CLI_TENANT_ID=...
SP_CLI_CLIENT_ID=...

m365 auth login    # one-time browser sign-in; token stored in OS keychain
m365 sites list    # works, scoped to your account's access
```

App registration prerequisite: **Authentication → Advanced settings → Allow public client flows → Yes** (required for device code).

Tokens expire (typically 1 hour access / 90-day refresh). Re-run `m365 auth login` if you get auth errors.

**`Sites.Selected` — app-only, scoped to specific sites**

Best for headless agents. The app has no access until a SharePoint admin explicitly grants it per-site. Requires a client secret; no interactive login.

```bash
# .env
SP_CLI_TENANT_ID=...
SP_CLI_CLIENT_ID=...
SP_CLI_CLIENT_SECRET=...

m365 sites list    # works only for sites the admin has granted
```

Admin grants access once per site via Graph API (see README App Registration section).

**`Sites.ReadWrite.All` — app-only, full tenant access**

Easiest to set up but broadest blast radius. Requires Global/SharePoint Admin consent. Avoid unless you genuinely need tenant-wide access.

### Auth commands
```
m365 auth setup    [--tenant-id <id>] [--client-id <id>] [--tenant-url <url>]
m365 auth login    # Device code flow; requires tenantId + clientId configured
m365 auth logout   # Delete stored keychain token
m365 auth token    # Print current access token; --raw for bare string
```

---

## Configuration

Config is stored at `~/.sp-cli/config.json` (mode 0600).

```
m365 config set [--tenant <url>] [--site <siteId>] [--drive <driveId>] [--tenant-id <id>] [--client-id <id>]
m365 config get
```

**Keys:**
| Key | Description |
|---|---|
| `tenantId` | Azure AD tenant GUID or domain |
| `clientId` | App registration client ID |
| `clientSecret` | Client secret (omit for device code) |
| `tenantUrl` | Default SharePoint URL (e.g. `https://contoso.sharepoint.com`) |
| `defaultSiteId` | Default site ID used when `--site` is omitted |
| `defaultDriveId` | Default drive ID used when `--drive` is omitted |

---

## Global Option Precedence

Most commands accept `--site` and `--drive` flags that override config defaults:

1. CLI flag (`--site`, `--drive`, `--tenant`)
2. Config file defaults (`defaultSiteId`, `defaultDriveId`, `tenantUrl`)

---

## Output Format

All commands output JSON to **stdout**:
```json
{ "data": <result> }
```
Errors go to **stderr**:
```json
{ "error": { "code": "...", "message": "...", "status": 404 } }
```

Use `--raw` (on `m365 auth token`) or pipe through `jq .data` to extract values.

---

## Discover Sites and Drives

Before using file or list commands, discover the IDs you need:

```bash
# List all sites in tenant
m365 sites list | jq '.data[] | {id, displayName, webUrl}'

# Get root site
m365 sites get | jq '.data.id'

# Get site by URL
m365 sites get --url https://contoso.sharepoint.com/sites/mysite

# List drives (document libraries) in a site
m365 drives list --site <site-id> | jq '.data[] | {id, name}'

# Save defaults to avoid repeating flags
m365 config set --site <site-id> --drive <drive-id>
```

---

## Required App Registration Permissions

**Delegated (device code) — no admin consent needed:**
- `Sites.ReadWrite.All` (delegated) — covers lists and list items too
- `Files.ReadWrite.All` (delegated)
- `Mail.ReadWrite` (delegated)
- `Mail.Send` (delegated)
- `Calendars.ReadWrite` (delegated)
- `Calendars.ReadWrite.Shared` (delegated)
- `offline_access` (delegated)

**`Sites.Selected` — app-only, scoped:**
- `Sites.Selected` (application) — admin consents this, then grants per-site via Graph API

**`Sites.ReadWrite.All` — app-only, full tenant:**
- `Sites.ReadWrite.All` (application) — requires Global/SharePoint Admin consent
- `Files.ReadWrite.All` (application)
- `Lists.ReadWrite.All` (application)

---

## Security Rules

- Never log or expose the access token in command output (use `m365 auth token --raw | ...` and pipe directly).
- Config file is written with mode 0600; never commit `~/.sp-cli/config.json`.
- Client secrets should use `m365 auth setup` rather than being placed in environment variables in scripts that get committed.
