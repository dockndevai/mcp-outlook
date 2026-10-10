import { z } from "zod";
import type { ToolDef } from "./types.js";
import { jsonResult } from "./types.js";

const folder = z
  .string()
  .optional()
  .describe("Mail folder id or well-known name (inbox, sentitems, drafts, archive, deleteditems). Defaults to the whole mailbox.");

export const readTools: ToolDef[] = [
  {
    name: "whoami",
    capability: "read",
    config: {
      title: "Who am I",
      description:
        "Return the signed-in mailbox's identity (display name, user principal name, id). Use this to confirm which " +
        "mailbox the server is operating on before reading or writing.",
      inputSchema: {},
    },
    handler: async (_a, { client, policy }) => {
      policy.guard({ tool: "whoami", capability: "read" });
      const me = (await client.me()) as Record<string, unknown>;
      return jsonResult({
        id: me.id,
        displayName: me.displayName,
        userPrincipalName: me.userPrincipalName ?? me.mail,
        mail: me.mail,
      });
    },
  },
  {
    name: "list_folders",
    capability: "read",
    config: {
      title: "List mail folders",
      description:
        "List the mailbox's top-level mail folders with their ids, display names, and unread/total counts. Use a folder " +
        "id or well-known name (inbox, sentitems, drafts, archive) with list_messages.",
      inputSchema: {},
    },
    handler: async (_a, { client, policy }) => {
      policy.guard({ tool: "list_folders", capability: "read" });
      return jsonResult(await client.listFolders());
    },
  },
  {
    name: "list_messages",
    capability: "read",
    config: {
      title: "List messages",
      description:
        "List messages, most recent first. Optionally scope to a folder and/or only unread messages. Returns compact " +
        "summaries (id, subject, from, received, isRead, hasAttachments, preview). Fetch full bodies with get_message.",
      inputSchema: {
        folder,
        unread_only: z.boolean().optional().describe("Only return unread messages"),
        top: z.number().int().min(1).max(200).optional().describe("Maximum messages to return (capped by OUTLOOK_MAX_RESULTS)"),
      },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "list_messages", capability: "read", folder: a.folder as string | undefined });
      return jsonResult(
        await client.listMessages({
          folder: a.folder as string | undefined,
          unreadOnly: a.unread_only as boolean | undefined,
          top: a.top as number | undefined,
        }),
      );
    },
  },
  {
    name: "search_messages",
    capability: "read",
    config: {
      title: "Search messages",
      description:
        "Full-text search across the mailbox (subject, body, sender, recipients) using Microsoft Graph $search. " +
        'Example queries: "invoice", "from:alice@contoso.com", "subject:release". Returns compact summaries.',
      inputSchema: {
        query: z.string().min(1).describe('Search text, e.g. "quarterly report" or "from:bob@contoso.com"'),
        top: z.number().int().min(1).max(200).optional().describe("Maximum results (capped by OUTLOOK_MAX_RESULTS)"),
      },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "search_messages", capability: "read" });
      return jsonResult(await client.searchMessages(a.query as string, a.top as number | undefined));
    },
  },
  {
    name: "get_message",
    capability: "read",
    config: {
      title: "Get message",
      description:
        "Fetch a single message by id with its full body, recipients (to/cc/bcc), importance, and conversation id. " +
        "Get ids from list_messages or search_messages.",
      inputSchema: { id: z.string().describe("Message id (from list_messages or search_messages)") },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "get_message", capability: "read" });
      return jsonResult(await client.getMessage(a.id as string));
    },
  },
  {
    name: "list_attachments",
    capability: "read",
    config: {
      title: "List attachments",
      description:
        "List a message's attachments as metadata only (id, name, contentType, size, isInline). Attachment bytes are " +
        "intentionally not returned.",
      inputSchema: { id: z.string().describe("Message id (from list_messages or search_messages)") },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "list_attachments", capability: "read" });
      return jsonResult(await client.listAttachments(a.id as string));
    },
  },
  {
    name: "list_contacts",
    capability: "read",
    config: {
      title: "List contacts",
      description: "List the mailbox's personal contacts (display name, email addresses, company).",
      inputSchema: {
        top: z.number().int().min(1).max(200).optional().describe("Maximum contacts to return (capped by OUTLOOK_MAX_RESULTS)"),
      },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "list_contacts", capability: "read" });
      return jsonResult(await client.listContacts(a.top as number | undefined));
    },
  },

  // ===== OneDrive files =====
  {
    name: "list_drive_items",
    capability: "read",
    config: {
      title: "List OneDrive items",
      description:
        "List the files and folders in a OneDrive folder. Give a `path` relative to the drive root (e.g. " +
        '"/Documents"), or an `item_id` from a previous result. Omit both for the root. Returns id, name, ' +
        "isFolder, size, childCount, lastModified, webUrl.",
      inputSchema: {
        path: z.string().optional().describe('Folder path from the drive root, e.g. "/Documents/Reports". Omit for root.'),
        item_id: z.string().optional().describe("Folder item id (from a previous list). Alternative to path."),
      },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "list_drive_items", capability: "read" });
      return jsonResult(await client.listDriveItems({ path: a.path as string | undefined, id: a.item_id as string | undefined }));
    },
  },
  {
    name: "get_drive_item",
    capability: "read",
    config: {
      title: "Get OneDrive item",
      description: "Fetch metadata for a single OneDrive file or folder by `path` or `item_id`.",
      inputSchema: {
        path: z.string().optional().describe("Item path from the drive root."),
        item_id: z.string().optional().describe("Item id (from list_drive_items). Alternative to path."),
      },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "get_drive_item", capability: "read" });
      return jsonResult(await client.getDriveItem({ path: a.path as string | undefined, id: a.item_id as string | undefined }));
    },
  },
  {
    name: "search_drive_files",
    capability: "read",
    config: {
      title: "Search OneDrive",
      description: "Search across OneDrive for files and folders matching a query (name/content).",
      inputSchema: {
        query: z.string().min(1).describe("Search text."),
        top: z.number().int().min(1).max(200).optional().describe("Maximum results (capped by OUTLOOK_MAX_RESULTS)"),
      },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "search_drive_files", capability: "read" });
      const top = Math.min(client.maxResults, (a.top as number) ?? client.maxResults);
      return jsonResult(await client.searchDrive(a.query as string, top));
    },
  },
  {
    name: "download_drive_file",
    capability: "read",
    config: {
      title: "Download a OneDrive file",
      description:
        "Download a OneDrive file's contents as text (capped at ~1 MB). Best for text/markdown/JSON/CSV; binary " +
        "files come back as best-effort UTF-8, so prefer webUrl for those. Give `path` or `item_id`.",
      inputSchema: {
        path: z.string().optional().describe("File path from the drive root, e.g. /Documents/notes.md."),
        item_id: z.string().optional().describe("File item id (from list_drive_items). Alternative to path."),
      },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "download_drive_file", capability: "read" });
      return jsonResult(
        await client.downloadFile({ path: a.path as string | undefined, id: a.item_id as string | undefined }, 1_000_000),
      );
    },
  },

  // --- Calendar ---
  {
    name: "list_calendars",
    capability: "read",
    config: {
      title: "List calendars",
      description: "List the calendars in the mailbox (id, name, owner, whether you can edit). Needs the Calendars scope.",
      inputSchema: {},
    },
    handler: async (_a, { client, policy }) => {
      policy.guard({ tool: "list_calendars", capability: "read" });
      return jsonResult(await client.listCalendars());
    },
  },
  {
    name: "list_events",
    capability: "read",
    config: {
      title: "List calendar events",
      description:
        "List events, newest window first. Give `start` and `end` (ISO 8601) to get a time window with recurring " +
        "events expanded; omit them to list upcoming/raw events. Optionally scope to one calendar.",
      inputSchema: {
        start: z.string().optional().describe("Window start, ISO 8601 (e.g. 2026-10-10T00:00:00Z). Pair with `end`."),
        end: z.string().optional().describe("Window end, ISO 8601. Pair with `start`."),
        calendar_id: z.string().optional().describe("Calendar id (from list_calendars). Defaults to the default calendar."),
        limit: z.number().int().min(1).max(200).optional().describe("Max events (capped by OUTLOOK_MAX_RESULTS)."),
      },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "list_events", capability: "read" });
      return jsonResult(
        await client.listEvents({
          start: a.start as string | undefined,
          end: a.end as string | undefined,
          calendarId: a.calendar_id as string | undefined,
          top: a.limit as number | undefined,
        }),
      );
    },
  },
  {
    name: "get_event",
    capability: "read",
    config: {
      title: "Get a calendar event",
      description: "Fetch one event by id, with its full body, attendees, recurrence and online-meeting details.",
      inputSchema: { id: z.string().min(1).describe("Event id (from list_events).") },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "get_event", capability: "read" });
      return jsonResult(await client.getEvent(a.id as string));
    },
  },
  {
    name: "get_free_busy",
    capability: "read",
    config: {
      title: "Get free/busy schedule",
      description:
        "Check availability (free/busy) for one or more people over a time window — useful before proposing a " +
        "meeting time. Returns each address's availability view for the window.",
      inputSchema: {
        schedules: z.array(z.string()).min(1).describe("Email addresses to check (your own and/or others)."),
        start: z.string().describe("Window start, ISO 8601 (UTC)."),
        end: z.string().describe("Window end, ISO 8601 (UTC)."),
        interval_minutes: z.number().int().min(5).max(1440).optional().describe("Availability slot size in minutes (default 30)."),
      },
    },
    handler: async (a, { client, policy }) => {
      policy.guard({ tool: "get_free_busy", capability: "read" });
      return jsonResult(
        await client.getSchedule(
          a.schedules as string[],
          a.start as string,
          a.end as string,
          (a.interval_minutes as number | undefined) ?? 30,
        ),
      );
    },
  },
];
