import { z } from "zod";
import type { ToolDef } from "./types.js";
import { jsonResult, textResult } from "./types.js";

/**
 * Admin tools. Registered only in `admin` mode, and gated behind OUTLOOK_ALLOW_DELETE=true, plus a
 * human confirmation via elicitation.
 *
 * Graph's DELETE on a message moves it to Deleted Items (recoverable from there) rather than
 * hard-purging it, which is the safe default — but it is still treated as destructive.
 */
export const adminTools: ToolDef[] = [
  {
    name: "delete_message",
    capability: "admin",
    destructive: true,
    config: {
      title: "Delete a message",
      description:
        "Delete a message by id. It is moved to the Deleted Items folder (recoverable from there), not permanently " +
        "purged. Requires admin mode and OUTLOOK_ALLOW_DELETE=true, and prompts for human confirmation.",
      inputSchema: { id: z.string().min(1).describe("Message id (from list_messages or search_messages).") },
    },
    handler: async (args, { client, policy, confirm }) => {
      const id = args.id as string;
      const { dryRun } = policy.guard({ tool: "delete_message", capability: "admin", destructive: true });
      if (dryRun) return textResult(`[dry-run] Would delete message ${id} (to Deleted Items).`);
      const ok = await confirm.confirm({ action: "delete message", target: id });
      if (!ok.approved) return textResult(`Deletion cancelled — ${ok.reason}.`);
      await client.deleteMessage(id);
      return jsonResult({ deleted: id, note: "Moved to Deleted Items (recoverable)." });
    },
  },
  {
    name: "delete_drive_item",
    capability: "admin",
    destructive: true,
    config: {
      title: "Delete a OneDrive item",
      description:
        "Delete a OneDrive file or folder by id. It goes to the OneDrive recycle bin (recoverable), not a permanent " +
        "purge. Requires admin mode and OUTLOOK_ALLOW_DELETE=true, and prompts for human confirmation.",
      inputSchema: { item_id: z.string().min(1).describe("Item id (from list_drive_items or search_drive_files).") },
    },
    handler: async (args, { client, policy, confirm }) => {
      const id = args.item_id as string;
      const { dryRun } = policy.guard({ tool: "delete_drive_item", capability: "admin", destructive: true });
      if (dryRun) return textResult(`[dry-run] Would delete OneDrive item ${id} (to recycle bin).`);
      const ok = await confirm.confirm({ action: "delete OneDrive item", target: id });
      if (!ok.approved) return textResult(`Deletion cancelled — ${ok.reason}.`);
      await client.deleteDriveItem(id);
      return jsonResult({ deleted: id, note: "Moved to the OneDrive recycle bin (recoverable)." });
    },
  },

  // --- Calendar (destructive) ---
  {
    name: "cancel_event",
    capability: "admin",
    destructive: true,
    config: {
      title: "Cancel a meeting",
      description:
        "Cancel a meeting you organize: a cancellation is sent to all attendees and the event is removed from their " +
        "calendars. Requires admin mode and OUTLOOK_ALLOW_DELETE=true, and prompts for confirmation. Organizer only — " +
        "use respond_to_event=decline for meetings you don't own.",
      inputSchema: {
        id: z.string().min(1).describe("Event id (from list_events)."),
        comment: z.string().optional().describe("Optional message included in the cancellation."),
      },
    },
    handler: async (args, { client, policy, confirm }) => {
      const id = args.id as string;
      const { dryRun } = policy.guard({ tool: "cancel_event", capability: "admin", destructive: true });
      if (dryRun) return textResult(`[dry-run] Would cancel meeting ${id} and notify attendees.`);
      const ok = await confirm.confirm({ action: "cancel meeting", target: id });
      if (!ok.approved) return textResult(`Cancellation stopped — ${ok.reason}.`);
      await client.cancelEvent(id, (args.comment as string | undefined) ?? "");
      return jsonResult({ cancelled: id, note: "Cancellation sent to attendees." });
    },
  },
  {
    name: "delete_event",
    capability: "admin",
    destructive: true,
    config: {
      title: "Delete a calendar event",
      description:
        "Delete an event from the calendar. If you organize it, Graph sends cancellations to attendees. Requires admin " +
        "mode and OUTLOOK_ALLOW_DELETE=true, and prompts for confirmation.",
      inputSchema: { id: z.string().min(1).describe("Event id (from list_events).") },
    },
    handler: async (args, { client, policy, confirm }) => {
      const id = args.id as string;
      const { dryRun } = policy.guard({ tool: "delete_event", capability: "admin", destructive: true });
      if (dryRun) return textResult(`[dry-run] Would delete event ${id}.`);
      const ok = await confirm.confirm({ action: "delete calendar event", target: id });
      if (!ok.approved) return textResult(`Deletion cancelled — ${ok.reason}.`);
      await client.deleteEvent(id);
      return jsonResult({ deleted: id });
    },
  },
];
