---
name: m365-files
version: 0.1.0
description: "SharePoint Files: Upload, download, copy, move, rename, delete, and search files in SharePoint document libraries."
metadata:
  openclaw:
    category: "productivity"
    requires:
      bins: ["m365"]
      skills: ["m365-shared"]
    cliHelp: "m365 files --help"
---

# m365-files — SharePoint File Operations

## Prerequisites

Complete auth setup from `m365-shared` SKILL.md, then discover your site and drive IDs:

```bash
m365 sites list | jq '.data[] | {id, displayName}'
m365 drives list --site <site-id> | jq '.data[] | {id, name}'
m365 config set --site <site-id> --drive <drive-id>
```

---

## Command Reference

### List files
```bash
m365 files list [--site <id>] [--drive <id>] [--path <folder-path>]

# Examples
m365 files list                              # Root of default drive
m365 files list --path /Documents            # Subfolder
m365 files list --path "/Shared Documents"
```

Output: array of DriveItem objects with `id`, `name`, `size`, `webUrl`, `file`/`folder`.

### Get file metadata
```bash
m365 files get <item-id> [--site <id>] [--drive <id>]

m365 files get 01ABC123...
```

### Upload a file
```bash
m365 files upload <local-path> [--site <id>] [--drive <id>] [--remote-path <path>]

m365 files upload ./report.xlsx --remote-path /Documents/report.xlsx
m365 files upload ./data.csv    # Uploads to drive root as data.csv
```

- Uses PUT to `/drives/{id}/root:{remotePath}:/content`
- Overwrites if file already exists at that path
- For files >4 MB, consider chunked upload (not yet implemented — split large files first)

### Download a file
```bash
m365 files download <item-id> [--site <id>] [--drive <id>] [--output <local-path>]

m365 files download 01ABC123... --output ./downloaded.xlsx
m365 files download 01ABC123...              # Saves with original filename
```

### Copy a file
```bash
m365 files copy <item-id> --dest-path <path> [--dest-drive <id>] [--name <new-name>] [--site <id>] [--drive <id>]

m365 files copy 01ABC123... --dest-path /Archive
m365 files copy 01ABC123... --dest-path /Archive --name report-copy.xlsx
m365 files copy 01ABC123... --dest-path /Other --dest-drive 01DEF456...
```

- Graph copy is async; response may be a 202 with a monitor URL.

### Move a file
```bash
m365 files move <item-id> --dest-path <path> [--dest-drive <id>] [--site <id>] [--drive <id>]

m365 files move 01ABC123... --dest-path /Archive
```

### Rename a file or folder
```bash
m365 files rename <item-id> --name <new-name> [--site <id>] [--drive <id>]

m365 files rename 01ABC123... --name renamed-report.xlsx
```

### Delete a file or folder
```bash
m365 files delete <item-id> [--site <id>] [--drive <id>]

m365 files delete 01ABC123...
```

Deletes to the site recycle bin (recoverable for 30–93 days).

### Search for files
```bash
m365 files search <query> [--site <id>] [--drive <id>]

m365 files search "quarterly report"
m365 files search budget.xlsx
```

---

## Common Patterns

### Find and download the latest version of a file
```bash
item_id=$(m365 files search "budget 2025" | jq -r '.data[0].id')
m365 files download "$item_id" --output ./budget-2025.xlsx
```

### Upload and get the new item's ID
```bash
m365 files upload ./contract.pdf --remote-path /Legal/contract.pdf \
  | jq -r '.data.id'
```

### Bulk-delete files matching a pattern
```bash
m365 files list --path /Temp | jq -r '.data[] | select(.name | startswith("tmp_")) | .id' \
  | xargs -I{} m365 files delete {}
```

---

## API Resources

All commands wrap the MS Graph Drive API:
- `GET /drives/{driveId}/root/children` — list root
- `GET /drives/{driveId}/root:{path}:/children` — list folder
- `GET /drives/{driveId}/items/{itemId}` — get metadata
- `PUT /drives/{driveId}/root:{path}:/content` — upload
- `GET @microsoft.graph.downloadUrl` from item metadata — download
- `POST /drives/{driveId}/items/{itemId}/copy` — copy
- `PATCH /drives/{driveId}/items/{itemId}` — move/rename
- `DELETE /drives/{driveId}/items/{itemId}` — delete
- `GET /drives/{driveId}/root/search(q='...')` — search

---

## Discovering Commands

```bash
m365 files --help
m365 files upload --help
m365 files copy --help
```
