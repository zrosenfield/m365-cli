---
name: m365-calendar
version: 0.1.0
description: "Microsoft 365 Calendar: Create, read, update, and delete events; RSVP to invitations; access shared calendars."
metadata:
  openclaw:
    category: "productivity"
    requires:
      bins: ["m365"]
      skills: ["sp-shared"]
    cliHelp: "m365 calendar --help"
---

# m365-calendar — Calendar Operations

## Prerequisites

Complete auth setup from `sp-shared` SKILL.md. Calendar commands require delegated auth with Calendar permissions:

```bash
m365 auth login
```

Required app permissions (delegated):
- `Calendars.ReadWrite` — read and manage your calendar
- `Calendars.ReadWrite.Shared` — access shared/other users' calendars (requires `--user` flag)

---

## Command Reference

### List calendars

```bash
m365 calendar list [--user <email|id>]

m365 calendar list
m365 calendar list --user colleague@example.com
```

Output: array of Calendar objects with `id`, `name`, `color`, `isDefaultCalendar`, `canEdit`.

### Get a calendar

```bash
m365 calendar get <calendarId> [--user <email|id>]

m365 calendar get AAMkAGI2...
```

---

### List events

```bash
m365 calendar events list [--calendar <id>] [--filter <odata>] [--select <fields>] [--top <n>] [--user <email|id>]

m365 calendar events list --top 10
m365 calendar events list --calendar AAMkAGI2... --top 5
m365 calendar events list --filter "showAs eq 'busy'"
m365 calendar events list --select "id,subject,start,end,location"
m365 calendar events list --user colleague@example.com --top 5
```

### View events in a date range

```bash
m365 calendar events view --start <datetime> --end <datetime> [--calendar <id>] [--user <email|id>]

# Current week
m365 calendar events view --start 2026-03-06T00:00:00 --end 2026-03-13T00:00:00

# With timezone context (ISO datetime)
m365 calendar events view --start "2026-03-10T00:00:00" --end "2026-03-11T00:00:00"
```

Uses `calendarView` endpoint — returns expanded recurring instances within the range.

### Get an event

```bash
m365 calendar events get <eventId> [--user <email|id>]

m365 calendar events get AAMkAGI2...
```

---

### Create an event

```bash
m365 calendar events create \
  --subject <str> --start <datetime> --end <datetime> \
  [--timezone <tz>] [--all-day] \
  [--body <str>] [--html] \
  [--location <str>] \
  [--attendees <csv>] [--attendee-type required|optional] \
  [--show-as free|tentative|busy|oof|workingElsewhere] \
  [--reminder <minutes>] [--no-reminder] \
  [--online-meeting] \
  [--importance low|normal|high] \
  [--sensitivity normal|personal|private|confidential] \
  [--recurrence daily|weekly|monthly|yearly] \
  [--recur-interval <n>] [--recur-days <csv>] \
  [--recur-end <date> | --recur-count <n>] \
  [--calendar <id>] [--user <email|id>]
```

**Required for create:** `--subject`, `--start`, `--end`

**Examples:**

```bash
# Simple 30-minute meeting
m365 calendar events create \
  --subject "Sync with Alice" \
  --start 2026-03-10T10:00:00 \
  --end 2026-03-10T10:30:00 \
  --timezone "Eastern Standard Time"

# Full-featured meeting
m365 calendar events create \
  --subject "Quarterly review" \
  --start 2026-03-15T14:00:00 \
  --end 2026-03-15T15:00:00 \
  --timezone "Pacific Standard Time" \
  --attendees "alice@example.com,bob@example.com" \
  --location "Conference Room A" \
  --body "Please review the Q1 report before this meeting." \
  --show-as busy \
  --reminder 15 \
  --online-meeting

# All-day event
m365 calendar events create \
  --subject "Company Holiday" \
  --start 2026-03-17T00:00:00 \
  --end 2026-03-17T00:00:00 \
  --all-day \
  --show-as free

# Weekly recurring (Mon/Wed/Fri for 8 occurrences)
m365 calendar events create \
  --subject "Daily standup" \
  --start 2026-03-09T09:00:00 \
  --end 2026-03-09T09:15:00 \
  --timezone "Eastern Standard Time" \
  --recurrence weekly \
  --recur-days mon,wed,fri \
  --recur-count 8
```

### Update an event

```bash
m365 calendar events update <eventId> [same flags as create, all optional]

# Change subject and add a reminder
m365 calendar events update AAMkAGI2... \
  --subject "Updated: Quarterly review" \
  --reminder 30

# Move time slot
m365 calendar events update AAMkAGI2... \
  --start 2026-03-15T15:00:00 \
  --end 2026-03-15T16:00:00

# Disable reminder
m365 calendar events update AAMkAGI2... --no-reminder
```

### Delete an event

```bash
m365 calendar events delete <eventId> [--user <email|id>]

m365 calendar events delete AAMkAGI2...
```

---

### RSVP to invitations

```bash
# Accept
m365 calendar events accept <eventId> [--comment <str>] [--user <email|id>]

# Decline
m365 calendar events decline <eventId> [--comment <str>] [--user <email|id>]

# Tentative
m365 calendar events tentative <eventId> [--comment <str>] [--user <email|id>]
```

**Examples:**

```bash
m365 calendar events accept AAMkAGI2...
m365 calendar events decline AAMkAGI2... --comment "Conflict with another meeting"
m365 calendar events tentative AAMkAGI2... --comment "Will try to join"
```

All RSVP commands send the response to the organizer (`sendResponse: true`).

---

## Shared Calendar Access

Use `--user <email|upn>` on any calendar command to access another user's calendar (requires `Calendars.ReadWrite.Shared` and appropriate calendar sharing settings):

```bash
# View colleague's upcoming events
m365 calendar events view \
  --start 2026-03-06T00:00:00 --end 2026-03-13T00:00:00 \
  --user colleague@example.com

# List colleague's calendars
m365 calendar list --user colleague@example.com

# Create event on colleague's calendar
m365 calendar events create \
  --subject "Meeting" --start 2026-03-10T10:00:00 --end 2026-03-10T10:30:00 \
  --user colleague@example.com
```

---

## Event Flags Reference

| Flag | Description | Default |
|------|-------------|---------|
| `--subject <str>` | Event title | required (create) |
| `--start <datetime>` | Start (ISO8601) | required (create) |
| `--end <datetime>` | End (ISO8601) | required (create) |
| `--timezone <tz>` | Windows timezone name | `UTC` |
| `--all-day` | All-day event | off |
| `--body <str>` | Description/body text | |
| `--html` | Body is HTML | plain text |
| `--location <str>` | Location display name | |
| `--attendees <csv>` | Comma-separated email addresses | |
| `--attendee-type` | `required` or `optional` | `required` |
| `--show-as` | `free`, `tentative`, `busy`, `oof`, `workingElsewhere` | |
| `--reminder <n>` | Enable reminder N minutes before | |
| `--no-reminder` | Disable reminder | |
| `--online-meeting` | Add Teams meeting link | off |
| `--importance` | `low`, `normal`, `high` | |
| `--sensitivity` | `normal`, `personal`, `private`, `confidential` | |
| `--recurrence` | `daily`, `weekly`, `monthly`, `yearly` | |
| `--recur-interval <n>` | Every N periods | `1` |
| `--recur-days <csv>` | Days: `mon,tue,wed,thu,fri,sat,sun` | |
| `--recur-end <date>` | Recurrence end date (ISO date) | |
| `--recur-count <n>` | Number of occurrences | (no end) |
| `--calendar <id>` | Specific calendar | default calendar |
| `--user <email\|id>` | Shared calendar user | signed-in user |

**Timezone names** use Windows format (not IANA): `"Eastern Standard Time"`, `"Pacific Standard Time"`, `"UTC"`, `"GMT Standard Time"`, etc.

---

## Common Patterns

### Get this week's schedule

```bash
m365 calendar events view \
  --start "$(date -u +%Y-%m-%dT00:00:00)" \
  --end "$(date -u -d '+7 days' +%Y-%m-%dT00:00:00)" \
  | jq '.data[] | {subject, start: .start.dateTime, end: .end.dateTime}'
```

### Create and confirm an event

```bash
event_id=$(m365 calendar events create \
  --subject "Test" --start 2026-03-10T10:00:00 --end 2026-03-10T10:30:00 \
  | jq -r '.data.id')
m365 calendar events get "$event_id" | jq '.data | {id, subject, start}'
```

### Accept all pending invitations

```bash
m365 calendar events list \
  --filter "responseStatus/response eq 'notResponded'" \
  --select "id,subject" \
  | jq -r '.data[].id' | xargs -I{} m365 calendar events accept {}
```

### Delete an event after confirming it

```bash
m365 calendar events get AAMkAGI2... | jq '.data | {subject, start}'
m365 calendar events delete AAMkAGI2...
```

---

## API Resources

- `GET /me/calendars` — list calendars
- `GET /me/calendars/{id}` — get calendar
- `GET /me/calendar/events` — list events (default calendar)
- `GET /me/calendarView?startDateTime=...&endDateTime=...` — date-range view
- `GET /me/events/{id}` — get event
- `POST /me/calendar/events` — create event
- `PATCH /me/events/{id}` — update event
- `DELETE /me/events/{id}` — delete event
- `POST /me/events/{id}/accept` — accept invitation
- `POST /me/events/{id}/decline` — decline invitation
- `POST /me/events/{id}/tentativelyAccept` — tentative accept
- Replace `/me` with `/users/{email}` for shared calendar access

---

## Discovering Commands

```bash
m365 calendar --help
m365 calendar events --help
m365 calendar events create --help
m365 calendar events view --help
```
