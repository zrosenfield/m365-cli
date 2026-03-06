---
name: sp-lists
version: 0.1.0
description: "SharePoint Lists: Create and manage lists, list items, and columns (including document library columns)."
metadata:
  openclaw:
    category: "productivity"
    requires:
      bins: ["sp"]
      skills: ["sp-shared"]
    cliHelp: "sp lists --help"
---

# sp-lists — SharePoint List Operations

## Prerequisites

Complete auth setup from `sp-shared` SKILL.md, then set a default site:

```bash
sp sites list | jq '.data[] | {id, displayName}'
sp config set --site <site-id>
```

---

## List CRUD

### List all lists in a site
```bash
sp lists list [--site <id>]

sp lists list | jq '.data[] | {id, displayName, list}'
```

### Get a list
```bash
sp lists get <list-id> [--site <id>]

sp lists get 01ABC123-...
```

### Create a list
```bash
sp lists create --name <name> [--template generic|documentLibrary] [--site <id>]

sp lists create --name "Project Tasks"
sp lists create --name "Assets" --template documentLibrary
```

### Update a list
```bash
sp lists update <list-id> --name <new-name> [--site <id>]

sp lists update 01ABC123-... --name "Renamed Tasks"
```

### Delete a list
```bash
sp lists delete <list-id> [--site <id>]

sp lists delete 01ABC123-...
```

---

## List Items

### List items
```bash
sp lists items list <list-id> [--site <id>] [--filter <odata>] [--select <fields>]

sp lists items list 01ABC123-...
sp lists items list 01ABC123-... --filter "fields/Status eq 'Active'"
sp lists items list 01ABC123-... --select "fields/Title,fields/Status"
```

Output: array of listItem objects with expanded `fields`.

### Get an item
```bash
sp lists items get <list-id> <item-id> [--site <id>]

sp lists items get 01ABC123-... 42
```

### Create an item
```bash
sp lists items create <list-id> --fields '<json>' [--site <id>]

sp lists items create 01ABC123-... --fields '{"Title":"New Task","Status":"Active"}'
```

Field names must match the list's internal column names.

### Update an item
```bash
sp lists items update <list-id> <item-id> --fields '<json>' [--site <id>]

sp lists items update 01ABC123-... 42 --fields '{"Status":"Completed"}'
```

Only the provided fields are updated (PATCH semantics).

### Delete an item
```bash
sp lists items delete <list-id> <item-id> [--site <id>]

sp lists items delete 01ABC123-... 42
```

---

## Columns

Columns apply to both generic lists and document libraries. To manage document library columns, first find the library's list ID:

```bash
# Get the list ID for a document library named "Documents"
sp lists list | jq '.data[] | select(.displayName=="Documents") | .id'
```

Then use `sp lists columns *` with that list ID.

### List columns
```bash
sp lists columns list <list-id> [--site <id>]

sp lists columns list 01ABC123-... | jq '.data[] | {id, name, columnGroup}'
```

### Get a column
```bash
sp lists columns get <list-id> <column-id> [--site <id>]
```

### Add a column
```bash
sp lists columns add <list-id> --name <name> --type <type> [--required] [--site <id>]
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
sp lists columns add 01ABC123-... --name "Priority" --type choice
sp lists columns add 01ABC123-... --name "DueDate" --type dateTime --required
sp lists columns add 01ABC123-... --name "Owner" --type person
```

### Update a column
```bash
sp lists columns update <list-id> <column-id> [--name <name>] [--required <bool>] [--site <id>]

sp lists columns update 01ABC123-... col-id-... --name "Deadline"
sp lists columns update 01ABC123-... col-id-... --required true
```

### Remove a column
```bash
sp lists columns remove <list-id> <column-id> [--site <id>]

sp lists columns remove 01ABC123-... col-id-...
```

---

## Common Patterns

### Create a task tracker list with custom columns
```bash
list_id=$(sp lists create --name "Sprint Tasks" | jq -r '.data.id')
sp lists columns add "$list_id" --name "Status" --type choice --required
sp lists columns add "$list_id" --name "Assignee" --type person
sp lists columns add "$list_id" --name "DueDate" --type dateTime

sp lists items create "$list_id" --fields '{"Title":"Implement login","Status":"Active"}'
```

### Query items with OData filter
```bash
# Items where Status is 'Completed'
sp lists items list "$list_id" --filter "fields/Status eq 'Completed'" \
  | jq '.data[] | .fields | {id, Title, Status}'
```

### Bulk create items from JSON array
```bash
cat tasks.json | jq -c '.[]' | while read item; do
  sp lists items create "$list_id" --fields "$item"
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
sp lists --help
sp lists items --help
sp lists columns --help
sp lists columns add --help
```
