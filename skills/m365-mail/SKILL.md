---
name: m365-mail
version: 0.1.0
description: "Microsoft 365 Mail: Read, send, reply to, and delete email messages; manage mail folders."
metadata:
  openclaw:
    category: "productivity"
    requires:
      bins: ["m365"]
      skills: ["m365-shared"]
    cliHelp: "m365 mail --help"
---

# m365-mail — Mail Operations

## Prerequisites

Complete auth setup from `m365-shared` SKILL.md. Mail commands require delegated auth (device code login) with Mail permissions:

```bash
# Ensure you have Mail scopes — re-login if needed
m365 auth login
```

Required app permissions (delegated):
- `Mail.ReadWrite` — read and manage messages
- `Mail.Send` — send messages

---

## Command Reference

### List messages

```bash
m365 mail list [--folder <name|id>] [--filter <odata>] [--select <fields>] [--top <n>]

# Examples
m365 mail list                                # Inbox, latest 25
m365 mail list --top 10
m365 mail list --folder sentitems --top 5     # Sent items
m365 mail list --folder drafts
m365 mail list --filter "isRead eq false"     # Unread only
m365 mail list --select "id,subject,from,receivedDateTime"
```

Well-known folder names (case-insensitive): `inbox`, `sentitems`, `drafts`, `deleteditems`, `junkemail`, `archive`, `outbox`

Output: array of Message objects with `id`, `subject`, `from`, `toRecipients`, `body`, `receivedDateTime`, `isRead`.

### Get a message

```bash
m365 mail get <messageId>

m365 mail get AAMkAGI2...
```

### Send a message

```bash
m365 mail send --to <emails> --subject <str> --body <str> \
  [--cc <emails>] [--bcc <emails>] [--importance low|normal|high] [--html]

# Examples
m365 mail send --to alice@example.com --subject "Hello" --body "Hi there"
m365 mail send --to "alice@example.com,bob@example.com" --subject "Team update" \
  --body "<h1>Update</h1><p>See attached.</p>" --html
m365 mail send --to alice@example.com --subject "Urgent" --body "Call me" --importance high
```

- `--to`, `--cc`, `--bcc` accept comma-separated email addresses
- `--html` sets `body.contentType` to `HTML`; omitting it sends plain text
- Message is saved to Sent Items

### Reply to a message

```bash
m365 mail reply <messageId> --body <str> [--reply-all] [--html]

m365 mail reply AAMkAGI2... --body "Thanks, noted."
m365 mail reply AAMkAGI2... --body "Looping in the team." --reply-all
m365 mail reply AAMkAGI2... --body "<p>See below.</p>" --html
```

### Delete a message

```bash
m365 mail delete <messageId>

m365 mail delete AAMkAGI2...
```

Moves to Deleted Items (not permanently deleted).

### List mail folders

```bash
m365 mail folders list

m365 mail folders list | jq '.data[] | {id, displayName, totalItemCount}'
```

### Get a folder

```bash
m365 mail folders get <folderId>

m365 mail folders get AAMkAGI2...
```

### Track mail changes (delta query)

**Prefer this over polling `mail list --filter "isRead eq false"` for detecting new mail.** Delta queries are efficient and stateful — each call returns only what changed since the last.

```bash
# Initialize (first run): drain all pages, save delta link, emit nothing
m365 mail delta --init-quiet

# Subsequent calls: emit only changed messages as NDJSON since last run
m365 mail delta

# Pipe to jq to get subjects of new (non-deleted) messages
m365 mail delta | jq -r 'select(.["@removed"] == null) | .subject'

# Track a specific folder
m365 mail delta --folder sentItems --init-quiet
m365 mail delta --folder sentItems

# Reset state (next run does full sync)
m365 mail delta --reset
```

**Flags:**

| Flag | Default | Description |
|---|---|---|
| `--folder <name\|id>` | `inbox` | Well-known folder name or raw folder ID |
| `--state-file <path>` | `$XDG_STATE_HOME/m365-cli/mail-delta-<folder>.link` | Where the delta link is persisted |
| `--change-type <created\|updated\|deleted>` | (all) | Filter by change type (initial sync only) |
| `--select <fields>` | `internetMessageId,from,subject,receivedDateTime,isRead` | Fields to return; `internetMessageId` always appended |
| `--max-page-size <n>` | `50` | Max messages per page |
| `--reset` | — | Delete state file and exit |
| `--init-quiet` | — | First run: drain all pages silently, just save the delta link |
| `--format ndjson\|json` | `ndjson` | Output format (NDJSON by default; `json` gives standard `{ "data": [...] }` envelope) |

**Output format:** NDJSON by default (one JSON object per line). Each object is the raw Graph Message. Deleted messages arrive with `"@removed": {"reason": "deleted"}` — pass them through, don't filter.

**Required permissions (delegated):** `Mail.Read`

**State file location:** `$XDG_STATE_HOME/m365-cli/mail-delta-<folder>.link` (falls back to `~/.local/state/m365-cli/...`). Delete it or use `--reset` to force a full re-sync.

---

## Common Patterns

### Read unread messages from inbox

```bash
m365 mail list --filter "isRead eq false" --top 10 \
  | jq '.data[] | {id, subject, from: .from.emailAddress.address}'
```

### Send and verify delivery

```bash
m365 mail send --to me@example.com --subject "Test" --body "Hello"
m365 mail list --folder sentitems --top 1 | jq '.data[0].subject'
```

### Find and delete messages matching a subject

```bash
m365 mail list --filter "subject eq 'Old newsletter'" --select "id" \
  | jq -r '.data[].id' | xargs -I{} m365 mail delete {}
```

### Reply to the latest message in inbox

```bash
msg_id=$(m365 mail list --top 1 | jq -r '.data[0].id')
m365 mail reply "$msg_id" --body "Acknowledged, thank you."
```

---

## API Resources

- `GET /me/mailFolders/{folder}/messages` — list messages
- `GET /me/messages/{id}` — get message
- `POST /me/sendMail` — send message
- `POST /me/messages/{id}/reply` — reply
- `POST /me/messages/{id}/replyAll` — reply all
- `DELETE /me/messages/{id}` — delete message
- `GET /me/mailFolders` — list folders
- `GET /me/mailFolders/{id}` — get folder

---

## Discovering Commands

```bash
m365 mail --help
m365 mail send --help
m365 mail folders --help
```
