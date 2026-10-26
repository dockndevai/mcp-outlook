# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.3] - 2026-10-26

### Changed
- Dependency/security maintenance: ran `npm audit fix` to clear a high-severity `body-parser` DoS advisory (GHSA-27r3-jxmh-p828, body-parser 2.3.0 → 2.4.0) and moderate `fast-uri` advisories (GHSA-w74x-v6f9-j3xg, GHSA-xhq7-xmqj-769h, GHSA-wjr9-j795-jxcr, GHSA-333x-w4r6-6crp, fast-uri 3.1.8 → 3.1.9), both transitive. Took in-range patch bump for `@modelcontextprotocol/sdk` (1.32.1 → 1.32.2). `npm audit` now reports 0 vulnerabilities. No feature or behavior changes; the Graph OAuth/PKCE auth module (`src/graph/auth.ts`) was not touched.

## [0.4.2] - 2026-10-19

### Changed
- Routine dependency maintenance: `npm audit --audit-level=high` reported 0 advisories (no high/critical vulnerabilities to fix this pass). Took in-range patch bumps for `vitest` (5.0.3 → 5.0.5) and `zod` (4.6.6 → 4.6.7). No feature or behavior changes.

## [0.4.1] - 2026-10-12

### Maintenance
- Routine dependency maintenance: updated `zod`, `@types/node`, `tsx` and `vitest` to their latest
  in-range versions. No behavior change; `npm audit` reports 0 vulnerabilities before and after.

## [0.4.0] - 2026-10-10

### Added
- **Calendar support (Microsoft Graph).** Nine tools across read/read-write/admin: `list_calendars`,
  `list_events` (a start/end window expands recurring events via calendarView), `get_event`,
  `get_free_busy` (availability for one or more people); `create_event`, `update_event`,
  `respond_to_event` (accept/decline/tentative); and `cancel_event`, `delete_event`.
- Calendar **writes are gated** the same way as mail: creating/updating/responding notify
  attendees/organizers, so they need `OUTLOOK_ALLOW_SEND=true` plus a human confirmation; cancelling
  and deleting need admin mode plus `OUTLOOK_ALLOW_DELETE=true`.
- Added the `Calendars.ReadWrite` delegated scope to the defaults. **Access is governed by the Microsoft
  account's granted scopes** — if the token isn't consented for calendar, Graph returns 403; the server's
  modes/flags are defence in depth on top. Override with `OUTLOOK_SCOPES` to downscope.

## [0.3.0] - 2026-09-29

### Added
- **Optional AI risk guard (laya-guard).** Mutating tools can consult a local [laya-guard](https://github.com/dockndevai/laya-guard) daemon that classifies each operation allow / confirm / block before it runs. Runs after the deterministic policy gates and can only *tighten*, never grant; fails closed. Off by default (`OUTLOOK_GUARD_MODE=monitor|enforce`).

## [0.2.0] - 2026-09-17

### Added
- **OneDrive file management** (Microsoft Graph Files API), 8 new tools: `list_drive_items`,
  `get_drive_item`, `search_drive_files`, `download_drive_file` (read); `upload_drive_file`,
  `create_drive_folder`, `move_drive_item` (read-write); and `delete_drive_item` (admin +
  `OUTLOOK_ALLOW_DELETE`, → recycle bin, with human confirmation). Items are addressed by a path
  relative to the drive root or by item id. The interactive login now also requests the
  `Files.ReadWrite` scope. 22 tools total.

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
