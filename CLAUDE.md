# m365-cli — Microsoft 365 CLI

## Prerequisite Skill

Before using any `m365` commands, read: [`skills/m365-shared/SKILL.md`](skills/m365-shared/SKILL.md)

This teaches you:
- How to authenticate (service principal, device code, env var)
- How to configure defaults (tenant URL, site ID, drive ID)
- How to discover site and drive IDs
- Output format and error handling

## Available Skills

| Skill | Path | Covers |
|---|---|---|
| `m365-shared` | `skills/m365-shared/SKILL.md` | Auth, config, sites, drives — **read first** |
| `m365-files` | `skills/m365-files/SKILL.md` | File upload/download/copy/move/delete/search |
| `m365-lists` | `skills/m365-lists/SKILL.md` | List CRUD, list items, column management |
| `m365-mail` | `skills/m365-mail/SKILL.md` | Read/send/reply/delete mail, folders |
| `m365-calendar` | `skills/m365-calendar/SKILL.md` | Calendar CRUD, events, RSVP, shared calendars |

## Quick Reference

```bash
m365 auth setup --tenant-id <id> --client-id <id> --client-secret <secret>
m365 sites list
m365 drives list --site <id>
m365 config set --site <id> --drive <id>
m365 files list
m365 lists list
m365 mail list --top 5
m365 calendar events list --top 5
```

## Development

```bash
npm install
npm run build        # compiles TypeScript → dist/
npm link             # makes `m365` available globally
```

Source: `src/index.ts` → commands in `src/commands/`, library in `src/lib/`.
