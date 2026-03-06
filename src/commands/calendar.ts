import { Command } from "commander";
import { graph, validateId } from "../lib/graph.js";
import { outputData, handleCommandError } from "../lib/output.js";

const DAY_MAP: Record<string, string> = {
  sun: "sunday",
  mon: "monday",
  tue: "tuesday",
  wed: "wednesday",
  thu: "thursday",
  fri: "friday",
  sat: "saturday",
};

function calBase(opts: { user?: string }): string {
  return opts.user ? `/users/${encodeURIComponent(opts.user)}` : "/me";
}

function parseEmailCsv(csv: string, fieldName: string): string[] {
  return csv.split(",").map((e) => {
    const trimmed = e.trim();
    if (!trimmed.includes("@")) {
      throw new Error(`Invalid email address in ${fieldName}: "${trimmed}"`);
    }
    return trimmed;
  });
}

function buildRecurrence(opts: Record<string, unknown>): object | undefined {
  if (!opts.recurrence) return undefined;
  const typeMap: Record<string, string> = {
    daily: "daily",
    weekly: "weekly",
    monthly: "absoluteMonthly",
    yearly: "absoluteYearly",
  };
  const range =
    opts.recurEnd
      ? { type: "endDate", endDate: opts.recurEnd as string }
      : opts.recurCount
      ? { type: "numbered", numberOfOccurrences: parseInt(opts.recurCount as string) }
      : { type: "noEnd" };
  const startDate =
    typeof opts.start === "string" ? opts.start.split("T")[0] : "";
  return {
    pattern: {
      type: typeMap[opts.recurrence as string],
      interval: parseInt((opts.recurInterval as string) ?? "1"),
      daysOfWeek: opts.recurDays
        ? (opts.recurDays as string)
            .split(",")
            .map((d) => DAY_MAP[d.trim()])
            .filter(Boolean)
        : [],
    },
    range: {
      ...range,
      startDate,
      recurrenceTimeZone: (opts.timezone as string) ?? "UTC",
    },
  };
}

function buildEventBody(opts: Record<string, unknown>): Record<string, unknown> {
  const body: Record<string, unknown> = {};

  if (opts.subject !== undefined) body.subject = opts.subject;

  const tz = (opts.timezone as string) ?? "UTC";
  if (opts.start !== undefined) body.start = { dateTime: opts.start, timeZone: tz };
  if (opts.end !== undefined) body.end = { dateTime: opts.end, timeZone: tz };
  if (opts.allDay) body.isAllDay = true;

  if (opts.body !== undefined) {
    body.body = {
      contentType: opts.html ? "HTML" : "Text",
      content: opts.body,
    };
  }

  if (opts.location !== undefined) {
    body.location = { displayName: opts.location };
  }

  if (opts.attendees !== undefined) {
    const type = (opts.attendeeType as string) ?? "required";
    body.attendees = parseEmailCsv(opts.attendees as string, "--attendees").map(
      (addr) => ({ emailAddress: { address: addr }, type })
    );
  }

  if (opts.showAs !== undefined) body.showAs = opts.showAs;
  if (opts.importance !== undefined) body.importance = opts.importance;
  if (opts.sensitivity !== undefined) body.sensitivity = opts.sensitivity;

  if (opts.onlineMeeting) {
    body.isOnlineMeeting = true;
    body.onlineMeetingProvider = "teamsForBusiness";
  }

  // opts.reminder: true = not set (Commander default for --no-reminder pattern)
  //                false = --no-reminder was passed
  //                string = --reminder <minutes> was passed
  if (opts.reminder === false) {
    body.isReminderOn = false;
  } else if (typeof opts.reminder === "string") {
    body.isReminderOn = true;
    body.reminderMinutesBeforeStart = parseInt(opts.reminder);
  }

  const recurrence = buildRecurrence(opts);
  if (recurrence) body.recurrence = recurrence;

  return body;
}

function addEventFlags(cmd: Command): Command {
  return cmd
    .option("--subject <str>", "Event subject")
    .option("--start <datetime>", "Start datetime (ISO8601)")
    .option("--end <datetime>", "End datetime (ISO8601)")
    .option("--timezone <tz>", "Timezone (default: UTC)")
    .option("--all-day", "All-day event")
    .option("--body <str>", "Event body/description")
    .option("--html", "Treat body as HTML")
    .option("--location <str>", "Location display name")
    .option("--attendees <csv>", "Attendee emails, comma-separated")
    .option("--attendee-type <type>", "Attendee type: required or optional (default: required)")
    .option("--show-as <status>", "Show as: free, tentative, busy, oof, workingElsewhere")
    .option("--reminder <minutes>", "Reminder minutes before start")
    .option("--no-reminder", "Disable reminder")
    .option("--online-meeting", "Add Teams online meeting link")
    .option("--importance <level>", "Importance: low, normal, high")
    .option("--sensitivity <level>", "Sensitivity: normal, personal, private, confidential")
    .option("--recurrence <type>", "Recurrence pattern: daily, weekly, monthly, yearly")
    .option("--recur-interval <n>", "Recurrence interval (default: 1)")
    .option("--recur-days <csv>", "Days of week for weekly recurrence: mon,tue,wed,thu,fri,sat,sun")
    .option("--recur-end <date>", "Recurrence end date (ISO date, e.g. 2026-12-31)")
    .option("--recur-count <n>", "Number of occurrences")
    .option("--calendar <id>", "Target calendar ID")
    .option("--user <email|id>", "Target user for shared calendar access");
}

function calendarPath(opts: Record<string, unknown>): string {
  if (opts.calendar) {
    validateId(opts.calendar as string, "calendar ID");
    return `/calendars/${opts.calendar as string}`;
  }
  return "/calendar";
}

export function registerCalendarCommands(program: Command): void {
  const calendar = program.command("calendar").description("Calendar operations");

  // m365 calendar list
  calendar
    .command("list")
    .description("List calendars")
    .option("--user <email|id>", "Target user")
    .action(async (opts) => {
      try {
        const base = calBase(opts);
        const result = await graph.get<{ value: unknown[] }>(`${base}/calendars`);
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 calendar get <calendarId>
  calendar
    .command("get <calendarId>")
    .description("Get a calendar by ID")
    .option("--user <email|id>", "Target user")
    .action(async (calendarId, opts) => {
      try {
        validateId(calendarId, "calendar ID");
        const base = calBase(opts);
        const result = await graph.get<unknown>(`${base}/calendars/${calendarId}`);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // events subcommand group
  const events = calendar.command("events").description("Calendar event operations");

  // m365 calendar events list
  events
    .command("list")
    .description("List events in a calendar")
    .option("--calendar <id>", "Calendar ID")
    .option("--filter <odata>", "OData filter expression")
    .option("--select <fields>", "Comma-separated fields to include")
    .option("--top <n>", "Max number of events to return", "25")
    .option("--user <email|id>", "Target user")
    .action(async (opts) => {
      try {
        const base = calBase(opts);
        const calPath = calendarPath(opts);
        const params = new URLSearchParams();
        params.set("$top", opts.top);
        if (opts.filter) params.set("$filter", opts.filter);
        if (opts.select) params.set("$select", opts.select);
        const result = await graph.get<{ value: unknown[] }>(
          `${base}${calPath}/events?${params.toString()}`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 calendar events view
  events
    .command("view")
    .description("View events in a date/time range (calendarView)")
    .requiredOption("--start <datetime>", "Range start (ISO8601)")
    .requiredOption("--end <datetime>", "Range end (ISO8601)")
    .option("--calendar <id>", "Calendar ID")
    .option("--user <email|id>", "Target user")
    .action(async (opts) => {
      try {
        const base = calBase(opts);
        const calPath = calendarPath(opts);
        const params = new URLSearchParams({
          startDateTime: opts.start,
          endDateTime: opts.end,
        });
        const result = await graph.get<{ value: unknown[] }>(
          `${base}${calPath}/calendarView?${params.toString()}`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 calendar events get <eventId>
  events
    .command("get <eventId>")
    .description("Get an event by ID")
    .option("--user <email|id>", "Target user")
    .action(async (eventId, opts) => {
      try {
        validateId(eventId, "event ID");
        const base = calBase(opts);
        const result = await graph.get<unknown>(`${base}/events/${eventId}`);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 calendar events create
  addEventFlags(events.command("create").description("Create a new calendar event")).action(
    async (opts) => {
      try {
        if (!opts.subject) throw new Error("--subject is required");
        if (!opts.start) throw new Error("--start is required");
        if (!opts.end) throw new Error("--end is required");
        const base = calBase(opts);
        const calPath = calendarPath(opts);
        const body = buildEventBody(opts as Record<string, unknown>);
        const result = await graph.post<unknown>(`${base}${calPath}/events`, body);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    }
  );

  // m365 calendar events update <eventId>
  addEventFlags(
    events.command("update <eventId>").description("Update an existing calendar event")
  ).action(async (eventId, opts) => {
    try {
      validateId(eventId, "event ID");
      const base = calBase(opts);
      const body = buildEventBody(opts as Record<string, unknown>);
      const result = await graph.patch<unknown>(`${base}/events/${eventId}`, body);
      outputData(result);
    } catch (err) {
      handleCommandError(err);
    }
  });

  // m365 calendar events delete <eventId>
  events
    .command("delete <eventId>")
    .description("Delete an event")
    .option("--user <email|id>", "Target user")
    .action(async (eventId, opts) => {
      try {
        validateId(eventId, "event ID");
        const base = calBase(opts);
        await graph.delete(`${base}/events/${eventId}`);
        outputData({ message: `Event ${eventId} deleted.` });
      } catch (err) {
        handleCommandError(err);
      }
    });

  // RSVP commands: accept, decline, tentative
  const rsvpCommands: [string, string, string][] = [
    ["accept", "accept", "accepted"],
    ["decline", "decline", "declined"],
    ["tentative", "tentativelyAccept", "tentatively accepted"],
  ];

  for (const [subCmd, endpoint, label] of rsvpCommands) {
    events
      .command(`${subCmd} <eventId>`)
      .description(`Mark event as ${label}`)
      .option("--comment <str>", "Optional response comment")
      .option("--user <email|id>", "Target user")
      .action(async (eventId, opts) => {
        try {
          validateId(eventId, "event ID");
          const base = calBase(opts);
          await graph.post<void>(`${base}/events/${eventId}/${endpoint}`, {
            comment: opts.comment ?? "",
            sendResponse: true,
          });
          outputData({ message: `Event ${eventId} ${label}.` });
        } catch (err) {
          handleCommandError(err);
        }
      });
  }
}
