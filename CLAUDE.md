# sp-cli — SharePoint Online CLI

## Prerequisite Skill

Before using any `sp` commands, read: [`skills/sp-shared/SKILL.md`](skills/sp-shared/SKILL.md)

This teaches you:
- How to authenticate (service principal, device code, env var)
- How to configure defaults (tenant URL, site ID, drive ID)
- How to discover site and drive IDs
- Output format and error handling

## Available Skills

| Skill | Path | Covers |
|---|---|---|
| `sp-shared` | `skills/sp-shared/SKILL.md` | Auth, config, sites, drives — **read first** |
| `sp-files` | `skills/sp-files/SKILL.md` | File upload/download/copy/move/delete/search |
| `sp-lists` | `skills/sp-lists/SKILL.md` | List CRUD, list items, column management |

## Quick Reference

```bash
sp auth setup --tenant-id <id> --client-id <id> --client-secret <secret>
sp sites list
sp drives list --site <id>
sp config set --site <id> --drive <id>
sp files list
sp lists list
```

## Development

```bash
npm install
npm run build        # compiles TypeScript → dist/
npm link             # makes `sp` available globally
```

Source: `src/index.ts` → commands in `src/commands/`, library in `src/lib/`.
