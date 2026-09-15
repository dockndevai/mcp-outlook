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
];
