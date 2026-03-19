---
name: m365-docs
version: 0.1.0
description: "SharePoint Docs: Create, co-edit, convert, and share .md/.txt/.docx documents collaboratively in chat."
metadata:
  openclaw:
    category: "productivity"
    requires:
      bins: ["m365", "pandoc"]
      skills: ["m365-shared", "m365-files"]
    cliHelp: "m365 files --help"
---

# m365-docs — Collaborative Document Editing

## Overview

This skill enables a full document co-editing loop entirely in chat:
1. Create a new doc from a topic, or open an existing one
2. Iterate on content with the user (suggest edits, rewrites, tone changes, etc.)
3. Save back to SharePoint as `.md`, `.txt`, or `.docx`
4. Share with someone via a link or permission grant

---

## Workflow

### 1. Start a session

**From a topic (new doc):**
```
User: "Make a new doc on topic Foo"
```
- Draft the document content in markdown
- Choose a filename: `foo.md` (slugify the topic, lowercase, hyphens)
- Write content to a temp file, upload to SharePoint, capture the item ID
- Confirm to the user: "Created foo.md — let's work on it. What would you like to change?"

```bash
# Write draft to temp file
cat > /tmp/foo.md << 'EOF'
# Foo

...content...
EOF

# Upload and capture item ID
item_id=$(m365 files upload /tmp/foo.md --remote-path /foo.md | jq -r '.data.id')
```

**From an existing file:**
```
User: "Work with me on bar.md"
```
- Search for the file, get its item ID
- Read its content into context

```bash
item_id=$(m365 files search "bar.md" | jq -r '.data[0].id')
content=$(m365 files read "$item_id")
```

---

### 2. Edit loop

Hold the current document content in context as a markdown string. Each turn:

- Show proposed changes inline (quote the before/after, or describe what changed)
- Ask: "Does this look right, or would you like to adjust anything?"
- On approval, update the in-context content
- Only save back to SharePoint when the user says they're done or explicitly asks to save

**Example prompts that trigger edits:**
- "Make it more formal"
- "Shorten the intro"
- "Add a section on X"
- "Rewrite the third paragraph"
- "Fix the tone — it sounds too aggressive"

Do NOT save after every turn. Save only when asked, or when the user says "done", "looks good", "save it", etc.

---

### 3. Save back

**As markdown (default — no conversion needed):**
```bash
printf '%s' "$content" > /tmp/foo.md
m365 files upload /tmp/foo.md --remote-path /foo.md
```

**As Word document (requires pandoc):**
```
User: "Make it a Word doc"
```
```bash
printf '%s' "$content" > /tmp/foo.md
pandoc /tmp/foo.md -o /tmp/foo.docx
docx_item_id=$(m365 files upload /tmp/foo.docx --remote-path /foo.docx | jq -r '.data.id')
```

Use the docx item ID for any subsequent sharing steps.

---

### 4. Share

**Create a shareable link:**
```
User: "Share it with me" / "Get me a link"
```
```bash
# View-only link, org-scoped (default)
m365 permissions create-link "$item_id" | jq -r '.data.link.webUrl'

# Editable link
m365 permissions create-link "$item_id" --type edit | jq -r '.data.link.webUrl'

# Anonymous link (anyone with the link)
m365 permissions create-link "$item_id" --scope anonymous | jq -r '.data.link.webUrl'
```
Return the URL to the user in chat.

**Grant access to a specific person:**
```
User: "Share it with alice@contoso.com"
```
```bash
m365 permissions grant "$item_id" --emails alice@contoso.com --role read
```

---

## State to track during a session

| Variable | What it holds |
|---|---|
| `$item_id` | SharePoint item ID of the active document |
| `$remote_path` | Path in the drive (e.g. `/foo.md`) |
| `$content` | Current markdown content (in context) |
| `$format` | `md`, `txt`, or `docx` |

---

## Common patterns

### Create, iterate, export to Word, and share — full example
```bash
# 1. Create
printf '%s' "$content" > /tmp/proposal.md
item_id=$(m365 files upload /tmp/proposal.md --remote-path /proposal.md | jq -r '.data.id')

# 2. (iterate in chat — no shell commands needed)

# 3. Export to Word
printf '%s' "$content" > /tmp/proposal.md
pandoc /tmp/proposal.md -o /tmp/proposal.docx
docx_id=$(m365 files upload /tmp/proposal.docx --remote-path /proposal.docx | jq -r '.data.id')

# 4. Share
m365 permissions create-link "$docx_id" --type edit | jq -r '.data.link.webUrl'
```

### Open an existing .txt or .md and continue editing
```bash
item_id=$(m365 files search "meeting-notes.md" | jq -r '.data[0].id')
content=$(m365 files read "$item_id")
# Now edit content in context, save back when done
```

---

## Notes

- `.md` and `.txt` round-trip perfectly — no conversion loss
- `.docx` via pandoc preserves headings, bold, italics, bullets, and tables; complex Word styles are not preserved
- `permissions create-link` with `--scope anonymous` requires the SharePoint tenant to allow anonymous sharing (admin setting)
- Item IDs change if a file is deleted and re-uploaded; prefer overwriting in place with `files upload` to the same `--remote-path`
