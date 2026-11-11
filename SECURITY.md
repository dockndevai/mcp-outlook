# Security

`mcp-outlook` gives an AI agent access to a Microsoft 365 mailbox — messages, contacts, and the
ability (in higher modes) to send mail on your behalf. Treat it like any other privileged automation
and grant it the least access it needs.

## Principles

- **Start read-only.** Leave `OUTLOOK_MODE=read-only` until you need to write. In read-only mode only
  the read tools are registered — draft/send/move/delete tools are not exposed to the model at all.
- **The Graph token is the primary control.** These flags are defence in depth. The real boundary is
  the OAuth token the server authenticates with: it carries its own delegated scopes (or application
  permissions) and the mailbox it can reach. Grant the narrowest scopes that work — **Mail.Read**
  alone means even a bug or a prompt injection cannot send or delete.
- **Sign-in happens in the browser.** In the default interactive mode the server runs the OAuth
  authorization-code + PKCE flow: your credentials go only to Microsoft, never to the server. The
  cached access/refresh token is written user-only (`~/.mcp-outlook/token.json`, mode 0600) and never
  logged.
- **Capabilities are gated by mode.** Every tool declares a capability (`read` / `write` / `admin`).
  A tool is registered only if the mode allows its capability, and each call is re-checked at runtime
  (`src/security.ts`).
- **Sending is doubly gated.** `send_mail`, `reply_mail` and `forward_mail` require `read-write` mode
  **and** `OUTLOOK_ALLOW_SEND=true`, because a sent message cannot be recalled. When the client
  supports MCP elicitation they also pause for a human to approve the exact recipients first.
- **Require a human.** Set `OUTLOOK_REQUIRE_HUMAN_APPROVAL=true` so every confirmed operation (`send_mail`,
  `reply_mail`, `forward_mail`, `create_event`, `update_event`, `respond_to_event`, `cancel_event`,
  `delete_message`, `delete_drive_item`, `delete_event`, and laya-guard `confirm` verdicts) is refused when
  the client cannot show an approval prompt. Each approval is bound to the exact tool and arguments
  (fingerprint). For existing objects it is also bound to the target's version: the Graph `changeKey` of
  the message or event, or the `eTag` of the OneDrive item. If that changed while the human was deciding,
  the action is refused. Sends, replies, forwards, event creation and invitation responses have no existing
  object to version; they are bound to the fingerprint of the exact arguments shown. Approval receipts are
  written to the audit log (`"audit":"approval"`, with `approvedBy` = `human` or `flag-only`).
- **Protect folders.** Folders listed in `OUTLOOK_PROTECTED_FOLDERS` (e.g. `sentitems`, `archive`)
  can be read but never written to or moved into. `OUTLOOK_FOLDER_ALLOWLIST` confines writes/moves to
  named folders.
- **Gate deletes.** `delete_message` requires both `admin` mode and `OUTLOOK_ALLOW_DELETE=true`, and
  moves the message to Deleted Items (recoverable) rather than hard-purging it.
- **Preview with dry-run.** `OUTLOOK_DRY_RUN=true` validates and logs write intent without executing.
- **Attachment bytes stay in the mailbox.** `list_attachments` returns metadata only (name, type,
  size) — the model never receives attachment content.

## Limitations

- Folder allowlist/protection is enforced on the folder a call names (the destination for
  `move_message`, or a scoped `list_messages`). Search spans the whole mailbox as the token allows.
- The approval prompt shows at most 1500 characters of the arguments; the recipients and subject of a
  send are listed separately, but a message body longer than that is only partly visible to the human
  (the fingerprint still covers all of it).
- Version pins are read from Graph before and after the prompt, not passed to the mutating call, so a change
  in the instant between the second read and the request is not caught. A message's `changeKey` also moves
  when it is marked read, and an event's when an attendee responds, so those cases fail closed and the
  operation must be re-requested. `upload_drive_file`, `move_message`, `move_drive_item` and
  `create_drive_folder` do not prompt at all.
- `send_mail` and friends trust the recipients and body you pass; there is no outbound content
  filtering beyond the human-confirmation prompt. Keep `OUTLOOK_ALLOW_SEND` off unless you need it.

## Reporting a vulnerability

Please open a private security advisory on the GitHub repository rather than a public issue.
