---
name: m365-lists
version: 0.1.0
description: "SharePoint Lists: Create and manage metadata lists, list items, and columns."
metadata:
  openclaw:
    category: "productivity"
    requires:
      bins: ["m365"]
      skills: ["m365-shared"]
    cliHelp: "m365 lists --help"
---

# m365-lists — SharePoint List Operations

## Lists vs Document Libraries

SharePoint has two distinct object types that both appear as "lists" in the Graph API:

| Type | Template | Purpose | Managed via |
|---|---|---|---|
| **List** | `generic` | Pure metadata: rows and columns, no file storage | `m365 lists` |
| **Document Library** | `documentLibrary` | File storage with metadata columns | `m365 files` / `m365 drives` |

**Always use `m365 lists create` for metadata lists only.**
To create a document library, use the SharePoint UI or `m365 drives`.

`m365 lists list` returns both types by default. Use `--type generic` to show only metadata lists.

## Prerequisites

Complete auth setup from `m365-shared` SKILL.md, then set a default site:

```bash
m365 sites list | jq '.data[] | {id, displayName}'
m365 config set --site <site-id>
```

---

## List CRUD

### List all lists in a site
```bash
m365 lists list [--site <id>] [--type generic|documentLibrary]

# Only metadata lists (not document libraries)
m365 lists list --type generic | jq '.data[] | {id, displayName}'

# All lists (includes document libraries)
m365 lists list | jq '.data[] | {id, displayName, list}'
```

### Get a list
```bash
m365 lists get <list-id> [--site <id>]

m365 lists get 01ABC123-...
```

### Create a list
```bash
# Creates a generic metadata list (NOT a document library)
m365 lists create --name <name> [--site <id>]

m365 lists create --name "Project Tasks"
```

### Update a list
```bash
m365 lists update <list-id> --name <new-name> [--site <id>]

m365 lists update 01ABC123-... --name "Renamed Tasks"
```

### Delete a list
```bash
m365 lists delete <list-id> [--site <id>]

m365 lists delete 01ABC123-...
```

---

## List Items

### List items
```bash
m365 lists items list <list-id> [--site <id>] [--filter <odata>] [--select <fields>]

m365 lists items list 01ABC123-...
m365 lists items list 01ABC123-... --filter "fields/Status eq 'Active'"
m365 lists items list 01ABC123-... --select "fields/Title,fields/Status"
```

Output: array of listItem objects with expanded `fields`.

### Get an item
```bash
m365 lists items get <list-id> <item-id> [--site <id>]

m365 lists items get 01ABC123-... 42
```

### Create an item
```bash
m365 lists items create <list-id> --fields '<json>' [--site <id>]

m365 lists items create 01ABC123-... --fields '{"Title":"New Task","Status":"Active"}'
```

Field names must match the list's internal column names.

### Update an item
```bash
m365 lists items update <list-id> <item-id> --fields '<json>' [--site <id>]

m365 lists items update 01ABC123-... 42 --fields '{"Status":"Completed"}'
```

Only the provided fields are updated (PATCH semantics).

### Delete an item
```bash
m365 lists items delete <list-id> <item-id> [--site <id>]

m365 lists items delete 01ABC123-... 42
```

---

## Columns

Columns apply to generic metadata lists. (Document library columns can also be managed this way, but use `m365 lists list --type documentLibrary` to find those list IDs.)

### List columns
```bash
m365 lists columns list <list-id> [--site <id>]

m365 lists columns list 01ABC123-... | jq '.data[] | {id, name, columnGroup}'
```

### Get a column
```bash
m365 lists columns get <list-id> <column-id> [--site <id>]
```

### Add a column
```bash
m365 lists columns add <list-id> --name <name> --type <type> [--required] [--site <id>]
```

**Supported types:**
| Type | Description |
|---|---|
| `text` | Single line of text |
| `number` | Numeric value |
| `boolean` | Yes/No checkbox |
| `dateTime` | Date and time picker |
| `choice` | Choice dropdown (choices editable via Graph) |
| `person` | Person or group picker |
| `lookup` | Lookup to another list |
| `hyperlink` | URL/hyperlink field |
| `currency` | Currency value (en-US locale) |

```bash
m365 lists columns add 01ABC123-... --name "Priority" --type choice
m365 lists columns add 01ABC123-... --name "DueDate" --type dateTime --required
m365 lists columns add 01ABC123-... --name "Owner" --type person
```

### Update a column
```bash
m365 lists columns update <list-id> <column-id> [--name <name>] [--required <bool>] [--site <id>]

m365 lists columns update 01ABC123-... col-id-... --name "Deadline"
m365 lists columns update 01ABC123-... col-id-... --required true
```

### Remove a column
```bash
m365 lists columns remove <list-id> <column-id> [--site <id>]

m365 lists columns remove 01ABC123-... col-id-...
```

---

## Common Patterns

### Create a task tracker list with custom columns
```bash
list_id=$(m365 lists create --name "Sprint Tasks" | jq -r '.data.id')
m365 lists columns add "$list_id" --name "Status" --type choice --required
m365 lists columns add "$list_id" --name "Assignee" --type person
m365 lists columns add "$list_id" --name "DueDate" --type dateTime

m365 lists items create "$list_id" --fields '{"Title":"Implement login","Status":"Active"}'
```

### Query items with OData filter
```bash
# Items where Status is 'Completed'
m365 lists items list "$list_id" --filter "fields/Status eq 'Completed'" \
  | jq '.data[] | .fields | {id, Title, Status}'
```

### Bulk create items from JSON array
```bash
cat tasks.json | jq -c '.[]' | while read item; do
  m365 lists items create "$list_id" --fields "$item"
done
```

---

## API Resources

- `GET /sites/{siteId}/lists` — list all lists
- `GET /sites/{siteId}/lists/{listId}` — get list
- `POST /sites/{siteId}/lists` — create list
- `PATCH /sites/{siteId}/lists/{listId}` — update list
- `DELETE /sites/{siteId}/lists/{listId}` — delete list
- `GET /sites/{siteId}/lists/{listId}/items?expand=fields` — list items
- `GET /sites/{siteId}/lists/{listId}/items/{itemId}?expand=fields` — get item
- `POST /sites/{siteId}/lists/{listId}/items` — create item
- `PATCH /sites/{siteId}/lists/{listId}/items/{itemId}/fields` — update item fields
- `DELETE /sites/{siteId}/lists/{listId}/items/{itemId}` — delete item
- `GET /sites/{siteId}/lists/{listId}/columns` — list columns
- `POST /sites/{siteId}/lists/{listId}/columns` — add column
- `PATCH /sites/{siteId}/lists/{listId}/columns/{columnId}` — update column
- `DELETE /sites/{siteId}/lists/{listId}/columns/{columnId}` — remove column

---

## Discovering Commands

```bash
m365 lists --help
m365 lists items --help
m365 lists columns --help
m365 lists columns add --help
```
