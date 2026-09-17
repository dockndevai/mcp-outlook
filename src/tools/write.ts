import { z } from "zod";
import type { Recipient } from "../graph/client.js";
import type { ToolDef } from "./types.js";
import { jsonResult, textResult } from "./types.js";

/**
 * Write tools. Registered only in `read-write` mode and up.
 *
 * Sending mail (send_mail, reply_mail, forward_mail) is doubly gated: it needs OUTLOOK_ALLOW_SEND=true
 * *and* pauses for a human confirmation via elicitation, because a sent message can't be unsent. The
 * non-sending writes (create_draft, move_message, mark_read) are reversible and only need read-write.
 */

const recipient = z.object({
  address: z.string().describe("Email address, e.g. alice@contoso.com"),
  name: z.string().optional().describe("Display name"),
});

function toRecipients(v: unknown): Recipient[] {
  if (!Array.isArray(v)) return [];
  return v as Recipient[];
}

function recipientList(rs: Recipient[]): string {
  return rs.map((r) => r.address).join(", ");
}

export const writeTools: ToolDef[] = [
  {
    name: "create_draft",
    capability: "write",
    config: {
      title: "Create a draft",
      description:
        "Create a draft message in the Drafts folder without sending it. Safe to use freely — nothing leaves the mailbox " +
        "until a human sends the draft from their client. Returns the created draft (including its id).",
      inputSchema: {
        subject: z.string().describe("Subject line"),
        body: z.string().describe("Message body"),
        to: z.array(recipient).min(1).describe("Primary recipients"),
        cc: z.array(recipient).optional().describe("Cc recipients"),
        bcc: z.array(recipient).optional().describe("Bcc recipients"),
        html: z.boolean().optional().describe("Treat the body as HTML instead of plain text"),
      },
    },
    handler: async (args, { client, policy }) => {
      const { dryRun } = policy.guard({ tool: "create_draft", capability: "write" });
      const to = toRecipients(args.to);
      if (dryRun) return textResult(`[dry-run] Would create draft "${args.subject as string}" to ${recipientList(to)}.`);
      return jsonResult(
        await client.createDraft({
          subject: args.subject as string,
          body: args.body as string,
          contentType: (args.html as boolean) ? "HTML" : "Text",
          to,
          cc: toRecipients(args.cc),
          bcc: toRecipients(args.bcc),
        }),
      );
    },
  },
  {
    name: "send_mail",
    capability: "write",
    requiresSend: true,
    config: {
      title: "Send an email",
      description:
        "Compose and send a new email immediately. Gated by OUTLOOK_ALLOW_SEND=true and a human confirmation prompt — a " +
        "sent message cannot be recalled. Use create_draft instead if you only want to prepare a message.",
      inputSchema: {
        subject: z.string().describe("Subject line"),
        body: z.string().describe("Message body"),
        to: z.array(recipient).min(1).describe("Primary recipients"),
        cc: z.array(recipient).optional().describe("Cc recipients"),
        bcc: z.array(recipient).optional().describe("Bcc recipients"),
        html: z.boolean().optional().describe("Treat the body as HTML instead of plain text"),
      },
    },
    handler: async (args, { client, policy, confirm }) => {
      const to = toRecipients(args.to);
      const { dryRun } = policy.guard({ tool: "send_mail", capability: "write", requiresSend: true });
      if (dryRun) return textResult(`[dry-run] Would send "${args.subject as string}" to ${recipientList(to)}.`);
      const ok = await confirm.confirm({
        action: "send email",
        target: recipientList(to),
        details: { subject: args.subject as string, cc: recipientList(toRecipients(args.cc)) || undefined },
      });
      if (!ok.approved) return textResult(`Send cancelled — ${ok.reason}.`);
      await client.sendMail({
        subject: args.subject as string,
        body: args.body as string,
        contentType: (args.html as boolean) ? "HTML" : "Text",
        to,
        cc: toRecipients(args.cc),
        bcc: toRecipients(args.bcc),
      });
      return textResult(`Sent "${args.subject as string}" to ${recipientList(to)}.`);
    },
  },
  {
    name: "reply_mail",
    capability: "write",
    requiresSend: true,
    config: {
      title: "Reply to a message",
      description:
        "Reply to an existing message (optionally reply-all) with a comment. Gated by OUTLOOK_ALLOW_SEND=true and a human " +
        "confirmation. Get the message id from list_messages or search_messages.",
      inputSchema: {
        id: z.string().describe("Message id to reply to"),
        comment: z.string().describe("Your reply text"),
        reply_all: z.boolean().optional().describe("Reply to all recipients instead of just the sender"),
      },
    },
    handler: async (args, { client, policy, confirm }) => {
      const id = args.id as string;
      const replyAll = Boolean(args.reply_all);
      const { dryRun } = policy.guard({ tool: "reply_mail", capability: "write", requiresSend: true });
      if (dryRun) return textResult(`[dry-run] Would ${replyAll ? "reply-all" : "reply"} to message ${id}.`);
      const ok = await confirm.confirm({ action: replyAll ? "reply-all to message" : "reply to message", target: id });
      if (!ok.approved) return textResult(`Reply cancelled — ${ok.reason}.`);
      await client.replyMessage(id, args.comment as string, replyAll);
      return textResult(`${replyAll ? "Replied to all" : "Replied"} on message ${id}.`);
    },
  },
  {
    name: "forward_mail",
    capability: "write",
    requiresSend: true,
    config: {
      title: "Forward a message",
      description:
        "Forward an existing message to new recipients with an optional comment. Gated by OUTLOOK_ALLOW_SEND=true and a " +
        "human confirmation.",
      inputSchema: {
        id: z.string().describe("Message id to forward"),
        to: z.array(recipient).min(1).describe("Recipients to forward to"),
        comment: z.string().optional().describe("Optional note added above the forwarded message"),
      },
    },
    handler: async (args, { client, policy, confirm }) => {
      const id = args.id as string;
      const to = toRecipients(args.to);
      const { dryRun } = policy.guard({ tool: "forward_mail", capability: "write", requiresSend: true });
      if (dryRun) return textResult(`[dry-run] Would forward message ${id} to ${recipientList(to)}.`);
      const ok = await confirm.confirm({ action: "forward message", target: recipientList(to), details: { message: id } });
      if (!ok.approved) return textResult(`Forward cancelled — ${ok.reason}.`);
      await client.forwardMessage(id, to, (args.comment as string) ?? "");
      return textResult(`Forwarded message ${id} to ${recipientList(to)}.`);
    },
  },
  {
    name: "mark_read",
    capability: "write",
    idempotent: true,
    config: {
      title: "Mark read / unread",
      description: "Mark a message as read or unread. Reversible.",
      inputSchema: {
        id: z.string().describe("Message id"),
        read: z.boolean().describe("true to mark read, false to mark unread"),
      },
    },
    handler: async (args, { client, policy }) => {
      const id = args.id as string;
      const read = Boolean(args.read);
      const { dryRun } = policy.guard({ tool: "mark_read", capability: "write" });
      if (dryRun) return textResult(`[dry-run] Would mark message ${id} as ${read ? "read" : "unread"}.`);
      await client.markRead(id, read);
      return textResult(`Marked message ${id} as ${read ? "read" : "unread"}.`);
    },
  },
  {
    name: "move_message",
    capability: "write",
    config: {
      title: "Move a message",
      description:
        "Move a message to another mail folder (by id or well-known name such as archive, deleteditems). Reversible: the " +
        "message keeps its content and can be moved back. The destination folder is subject to the folder allowlist.",
      inputSchema: {
        id: z.string().describe("Message id to move"),
        destination: z.string().describe("Destination folder id or well-known name (e.g. archive, deleteditems)"),
      },
    },
    handler: async (args, { client, policy }) => {
      const id = args.id as string;
      const destination = args.destination as string;
      const { dryRun } = policy.guard({ tool: "move_message", capability: "write", folder: destination });
      if (dryRun) return textResult(`[dry-run] Would move message ${id} to ${destination}.`);
      return jsonResult(await client.moveMessage(id, destination));
    },
  },

  // ===== OneDrive files =====
  {
    name: "upload_drive_file",
    capability: "write",
    config: {
      title: "Upload / write a OneDrive file",
      description:
        "Create or overwrite a text file in OneDrive at the given path (≤4 MB). Parent folders in the path are " +
        "created as needed. Reversible-ish: an overwritten file's previous version is kept in OneDrive version history.",
      inputSchema: {
        path: z.string().min(1).describe('File path from the drive root, e.g. "/Documents/notes.md".'),
        content: z.string().describe("Full UTF-8 text contents to write."),
      },
    },
    handler: async (args, { client, policy }) => {
      const path = args.path as string;
      const { dryRun } = policy.guard({ tool: "upload_drive_file", capability: "write" });
      if (dryRun) return textResult(`[dry-run] Would write OneDrive file ${path}.`);
      return jsonResult(await client.uploadFile(path, args.content as string));
    },
  },
  {
    name: "create_drive_folder",
    capability: "write",
    config: {
      title: "Create a OneDrive folder",
      description: "Create a new folder inside a OneDrive folder (by parent `path` or `item_id`; omit both for the root).",
      inputSchema: {
        name: z.string().min(1).describe("New folder name."),
        parent_path: z.string().optional().describe("Parent folder path from the drive root. Omit for root."),
        parent_id: z.string().optional().describe("Parent folder item id. Alternative to parent_path."),
      },
    },
    handler: async (args, { client, policy }) => {
      const { dryRun } = policy.guard({ tool: "create_drive_folder", capability: "write" });
      if (dryRun) return textResult(`[dry-run] Would create folder '${args.name as string}'.`);
      return jsonResult(
        await client.createFolder(
          { path: args.parent_path as string | undefined, id: args.parent_id as string | undefined },
          args.name as string,
        ),
      );
    },
  },
  {
    name: "move_drive_item",
    capability: "write",
    config: {
      title: "Move / rename a OneDrive item",
      description:
        "Move a OneDrive file or folder to another folder and/or rename it. Give the item id, and a new parent " +
        "folder id and/or a new name. Reversible.",
      inputSchema: {
        item_id: z.string().min(1).describe("Item id to move/rename (from list_drive_items)."),
        new_parent_id: z.string().optional().describe("Destination folder item id (to move it)."),
        new_name: z.string().optional().describe("New name (to rename it)."),
      },
    },
    handler: async (args, { client, policy }) => {
      const id = args.item_id as string;
      const { dryRun } = policy.guard({ tool: "move_drive_item", capability: "write" });
      if (dryRun) return textResult(`[dry-run] Would move/rename item ${id}.`);
      return jsonResult(
        await client.moveDriveItem(id, {
          newParentId: args.new_parent_id as string | undefined,
          newName: args.new_name as string | undefined,
        }),
      );
    },
  },
];
