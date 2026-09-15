# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-09-15

### Added
- Initial release: a safe-by-default MCP server for Microsoft Outlook mail over Microsoft Graph.
  14 tools across read/read-write/admin: `whoami`, `list_folders`, `list_messages`,
  `search_messages`, `get_message`, `list_attachments`, `list_contacts`, `create_draft`,
  `send_mail`, `reply_mail`, `forward_mail`, `mark_read`, `move_message`, and `delete_message`.
- **Browser sign-in** (default): OAuth authorization-code + PKCE with a loopback redirect. The
  server opens the Microsoft sign-in page, caches the access/refresh token on disk (0600), and
  refreshes silently — it never handles your password. App-only (client-credentials) and static
  bearer-token auth are also supported and auto-detected.
- Security model: access modes (read-only/read-write/admin), a dedicated **send gate**
  (`OUTLOOK_ALLOW_SEND`) on top of read-write for send/reply/forward, folder allowlist and protected
  folders, delete opt-in (`OUTLOOK_ALLOW_DELETE`), dry-run, JSON audit logging, and attachment-byte
  redaction.
- **Human-in-the-loop confirmation** via MCP [elicitation](https://modelcontextprotocol.io/specification/draft/client/elicitation):
  sending mail and deleting a message pause and ask the human to approve the exact action; clients
  that can't elicit fall back to the `*_ALLOW_*` flag gates.
- MCP tool annotations derived from each tool's capability, with a test keeping them consistent.
