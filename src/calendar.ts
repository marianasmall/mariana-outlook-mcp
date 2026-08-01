import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { getAccessToken } from "./auth.js";
import { graphFetch } from "./graph.js";
import { logAction } from "./logger.js";

const EVENT_FIELDS =
  "id,subject,start,end,location,bodyPreview,attendees,onlineMeeting,showAs,isCancelled,organizer";

function formatEvent(event: any) {
  return {
    id: event.id,
    title: event.subject || "(no title)",
    start: event.start?.dateTime ? `${event.start.dateTime} (${event.start.timeZone})` : "",
    end: event.end?.dateTime ? `${event.end.dateTime} (${event.end.timeZone})` : "",
    location: event.location?.displayName || "",
    description: event.bodyPreview || "",
    attendees: (event.attendees || []).map((a: any) => ({
      email: a.emailAddress?.address || "",
      name: a.emailAddress?.name || "",
      status: a.status?.response || "",
    })),
    videoLink: event.onlineMeeting?.joinUrl || "",
    status: event.isCancelled ? "cancelled" : event.showAs || "",
  };
}

/** Convert user-supplied ISO 8601 (with or without offset) to a Graph dateTimeTimeZone in UTC. */
function toGraphTime(iso: string): { dateTime: string; timeZone: string } {
  const d = new Date(iso);
  if (isNaN(d.getTime())) throw new Error(`Invalid date/time: "${iso}" (use ISO 8601, e.g. 2026-08-01T14:00:00-07:00)`);
  return { dateTime: d.toISOString().replace(/\.\d{3}Z$/, ""), timeZone: "UTC" };
}

export function registerCalendarTools(server: McpServer): void {
  // --- calendar_list ---
  server.tool(
    "calendar_list",
    "List upcoming calendar events. Times are returned in UTC.",
    {
      days_ahead: z.coerce.number().optional().default(7).describe("Days to look ahead (default 7)"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ days_ahead, account }) => {
      const { accessToken } = await getAccessToken(account);
      const now = new Date();
      const future = new Date(now.getTime() + days_ahead * 24 * 60 * 60 * 1000);

      const res = await graphFetch(
        accessToken,
        `/me/calendarView?startDateTime=${now.toISOString()}&endDateTime=${future.toISOString()}&$orderby=start/dateTime&$top=50&$select=${EVENT_FIELDS}`
      );

      const events = (res.value || []).map(formatEvent);
      return { content: [{ type: "text" as const, text: JSON.stringify(events, null, 2) }] };
    }
  );

  // --- calendar_search ---
  server.tool(
    "calendar_search",
    "Search calendar events by keyword in the title. Times are returned in UTC.",
    {
      query: z.string().describe("Search keyword (matched against event titles)"),
      time_min: z.string().optional().describe("Start of range (ISO 8601, default 1 year back)"),
      time_max: z.string().optional().describe("End of range (ISO 8601, default 1 year ahead)"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ query, time_min, time_max, account }) => {
      const { accessToken } = await getAccessToken(account);
      const start = time_min || new Date(Date.now() - 365 * 24 * 60 * 60 * 1000).toISOString();
      const end = time_max || new Date(Date.now() + 365 * 24 * 60 * 60 * 1000).toISOString();

      // Filter client-side: Graph's $filter on calendarView subject 500s
      // whenever the range contains an event with an empty subject
      // (documented Graph bug), which real mailboxes routinely do.
      const res = await graphFetch(
        accessToken,
        `/me/calendarView?startDateTime=${start}&endDateTime=${end}&$orderby=start/dateTime&$top=250&$select=${EVENT_FIELDS}`
      );

      const q = query.toLowerCase();
      const events = (res.value || [])
        .filter((e: any) => (e.subject || "").toLowerCase().includes(q))
        .slice(0, 50)
        .map(formatEvent);
      return { content: [{ type: "text" as const, text: JSON.stringify(events, null, 2) }] };
    }
  );

  // --- calendar_get ---
  server.tool(
    "calendar_get",
    "Get full details of a specific calendar event.",
    {
      event_id: z.string().describe("Event ID"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ event_id, account }) => {
      const { accessToken } = await getAccessToken(account);
      const res = await graphFetch(accessToken, `/me/events/${event_id}?$select=${EVENT_FIELDS}`);
      return { content: [{ type: "text" as const, text: JSON.stringify(formatEvent(res), null, 2) }] };
    }
  );

  // --- calendar_create ---
  server.tool(
    "calendar_create",
    "Create a calendar event. By default attendees are NOT invited (their names go into the description instead); set send_invites=true to actually invite them — Outlook sends invitation emails to every attendee on the event.",
    {
      title: z.string().describe("Event title"),
      start: z.string().describe("Start time (ISO 8601, e.g. '2026-08-01T14:00:00-07:00')"),
      end: z.string().describe("End time (ISO 8601)"),
      description: z.string().optional().describe("Event description"),
      location: z.string().optional().describe("Event location"),
      attendees: z.array(z.string()).optional().describe("Attendee email addresses"),
      send_invites: z.boolean().optional().default(false).describe("Actually invite attendees (default: false — Outlook emails invitations when true)"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ title, start, end, description, location, attendees, send_invites, account }) => {
      const { accessToken, accountName } = await getAccessToken(account);

      let bodyText = description || "";
      const inviteAttendees = send_invites && attendees && attendees.length > 0;
      if (!send_invites && attendees && attendees.length > 0) {
        bodyText = `${bodyText}${bodyText ? "\n\n" : ""}Planned attendees (not invited): ${attendees.join(", ")}`;
      }

      const res = await graphFetch(accessToken, "/me/events", {
        method: "POST",
        body: {
          subject: title,
          start: toGraphTime(start),
          end: toGraphTime(end),
          ...(bodyText ? { body: { contentType: "Text", content: bodyText } } : {}),
          ...(location ? { location: { displayName: location } } : {}),
          ...(inviteAttendees
            ? { attendees: attendees!.map((email) => ({ emailAddress: { address: email }, type: "required" })) }
            : {}),
        },
      });

      await logAction({
        timestamp: new Date().toISOString(),
        tool: "calendar_create",
        account: accountName,
        summary: `Event created: '${title}' at ${start}${inviteAttendees ? " (invites sent)" : " (no invites)"}`,
        ids: [res.id || "unknown"],
      });

      return {
        content: [
          {
            type: "text" as const,
            text: `Event created: "${title}" (ID: ${res.id}).${inviteAttendees ? " Invites sent." : " No invites sent."}`,
          },
        ],
      };
    }
  );

  // --- calendar_update ---
  server.tool(
    "calendar_update",
    "Modify an existing calendar event. CAUTION: if the event has invited attendees, Outlook notifies them of changes automatically.",
    {
      event_id: z.string().describe("Event ID to update"),
      title: z.string().optional().describe("New title"),
      start: z.string().optional().describe("New start time (ISO 8601)"),
      end: z.string().optional().describe("New end time (ISO 8601)"),
      description: z.string().optional().describe("New description"),
      location: z.string().optional().describe("New location"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ event_id, title, start, end, description, location, account }) => {
      const { accessToken, accountName } = await getAccessToken(account);

      const patch: Record<string, unknown> = {};
      if (title !== undefined) patch.subject = title;
      if (start !== undefined) patch.start = toGraphTime(start);
      if (end !== undefined) patch.end = toGraphTime(end);
      if (description !== undefined) patch.body = { contentType: "Text", content: description };
      if (location !== undefined) patch.location = { displayName: location };

      const res = await graphFetch(accessToken, `/me/events/${event_id}`, {
        method: "PATCH",
        body: patch,
      });

      const changes = [title && "title", start && "start", end && "end", description && "description", location && "location"]
        .filter(Boolean)
        .join(", ");

      await logAction({
        timestamp: new Date().toISOString(),
        tool: "calendar_update",
        account: accountName,
        summary: `Event updated (${changes}): '${res.subject}'`,
        ids: [event_id],
      });

      return {
        content: [{ type: "text" as const, text: `Event updated: "${res.subject}" (changed: ${changes})` }],
      };
    }
  );

  // --- calendar_flag_delete ---
  server.tool(
    "calendar_flag_delete",
    "Soft-delete: prepend 'DELETE - ' to event title. Does NOT actually delete the event.",
    {
      event_id: z.string().describe("Event ID to flag"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ event_id, account }) => {
      const { accessToken, accountName } = await getAccessToken(account);

      const current = await graphFetch(accessToken, `/me/events/${event_id}?$select=subject`);
      const originalTitle = current.subject || "(no title)";

      if (originalTitle.startsWith("DELETE - ")) {
        return {
          content: [{ type: "text" as const, text: `Event "${originalTitle}" is already flagged for deletion.` }],
        };
      }

      await graphFetch(accessToken, `/me/events/${event_id}`, {
        method: "PATCH",
        body: { subject: `DELETE - ${originalTitle}` },
      });

      await logAction({
        timestamp: new Date().toISOString(),
        tool: "calendar_flag_delete",
        account: accountName,
        summary: `Flagged for deletion: '${originalTitle}'`,
        ids: [event_id],
      });

      return {
        content: [
          { type: "text" as const, text: `Event flagged: "DELETE - ${originalTitle}". Review and delete manually in Outlook Calendar.` },
        ],
      };
    }
  );

  // --- calendar_availability ---
  server.tool(
    "calendar_availability",
    "Check free/busy status for a date or date range (derived from events not marked 'free').",
    {
      date_start: z.string().describe("Start date/time (ISO 8601)"),
      date_end: z.string().optional().describe("End date/time (ISO 8601, defaults to end of start day)"),
      account: z.string().optional().describe("Account name (default account if omitted)"),
    },
    async ({ date_start, date_end, account }) => {
      const { accessToken } = await getAccessToken(account);

      const startDate = new Date(date_start);
      // Default end-of-day in UTC: date-only strings parse as UTC midnight,
      // so local-time getters here would silently shrink the window by the
      // host's UTC offset (e.g. a ~7-hour "day" on a Pacific machine).
      const endDate = date_end
        ? new Date(date_end)
        : new Date(
            Date.UTC(startDate.getUTCFullYear(), startDate.getUTCMonth(), startDate.getUTCDate(), 23, 59, 59)
          );

      const res = await graphFetch(
        accessToken,
        `/me/calendarView?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}&$orderby=start/dateTime&$top=100&$select=subject,start,end,showAs`
      );

      const busy = (res.value || [])
        .filter((e: any) => e.showAs && e.showAs !== "free")
        .map((e: any) => ({
          start: e.start?.dateTime ? `${e.start.dateTime}Z` : "",
          end: e.end?.dateTime ? `${e.end.dateTime}Z` : "",
          showAs: e.showAs,
        }));

      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              { date_start, date_end: endDate.toISOString(), busy_blocks: busy, total_busy: busy.length },
              null,
              2
            ),
          },
        ],
      };
    }
  );
}
