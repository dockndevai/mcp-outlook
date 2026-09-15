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
];
