# mcp-outlook

[![npm](https://img.shields.io/npm/v/@dockndevai/mcp-outlook)](https://www.npmjs.com/package/@dockndevai/mcp-outlook)
[![CI](https://github.com/dockndevai/mcp-outlook/actions/workflows/ci.yml/badge.svg)](https://github.com/dockndevai/mcp-outlook/actions/workflows/ci.yml)
[![licence](https://img.shields.io/badge/licence-MIT-blue)](LICENSE)

A **safe-by-default** [Model Context Protocol](https://modelcontextprotocol.io) server for **Microsoft Outlook** mail, over [Microsoft Graph](https://learn.microsoft.com/graph/overview). It lets an agent read and operate a mailbox — list folders and messages, full-text **search**, read bodies and attachment metadata, list contacts, and (in higher modes) create drafts, **send / reply / forward**, mark read, move messages, and delete. It also manages **OneDrive files** — browse, search, download, upload, create folders, move/rename, and delete.

**Browser sign-in:** on first run it opens your browser to the Microsoft sign-in page, then caches the token and refreshes it silently — the server never sees your password.

Part of the [dockndevai MCP server suite](https://dockndevai.github.io/) — one governance model across all of them.

## What it gives an agent

The server starts **read-only** (see [Safe by default](#safe-by-default)); higher-capability tools are only registered when you raise the mode.

| Tool | For | Needs mode |
|---|---|---|
| `whoami` | confirm which mailbox is in use | read-only |
| `list_folders` | mail folders with unread/total counts | read-only |
| `list_messages` | recent messages (by folder, unread-only) | read-only |
| `search_messages` | full-text search across the mailbox | read-only |
| `get_message` | one message with full body + recipients | read-only |
| `list_attachments` | attachment metadata (bytes not returned) | read-only |
| `list_contacts` | personal contacts | read-only |
| `create_draft` | prepare a draft without sending | read-write |
| `send_mail` | compose & send a new email | read-write + `OUTLOOK_ALLOW_SEND` |
| `reply_mail` / `forward_mail` | reply (all) / forward a message | read-write + `OUTLOOK_ALLOW_SEND` |
| `mark_read` | mark read / unread (reversible) | read-write |
| `move_message` | move to another folder (reversible) | read-write |
| `delete_message` | delete (to Deleted Items) | admin + `OUTLOOK_ALLOW_DELETE` |
| `list_drive_items` / `get_drive_item` | browse OneDrive files & folders | read-only |
| `search_drive_files` | search OneDrive | read-only |
| `download_drive_file` | read a OneDrive file's text | read-only |
| `upload_drive_file` | create/overwrite a OneDrive file | read-write |
| `create_drive_folder` / `move_drive_item` | create folder / move-rename | read-write |
| `delete_drive_item` | delete a OneDrive item (→ recycle bin) | admin + `OUTLOOK_ALLOW_DELETE` |

## Install

```bash
npx -y @dockndevai/mcp-outlook
```

You need an **Entra (Azure AD) app registration**. For the default browser sign-in, register a **public client** and add the redirect URI `http://localhost` (platform: *Mobile and desktop applications*), then use its **Application (client) ID** as `OUTLOOK_CLIENT_ID`. Grant delegated **Mail.Read** (and **Mail.Send** / **Mail.ReadWrite** if you want to send or organize). No client secret is needed for interactive use.

Prefer automation? Use **app-only** auth instead: set `OUTLOOK_CLIENT_SECRET` + `OUTLOOK_TENANT_ID` + `OUTLOOK_USER` (see [Authentication](#authentication)).

## Configure

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "@dockndevai/mcp-outlook"],
      "env": {
        "OUTLOOK_CLIENT_ID": "00000000-0000-0000-0000-000000000000",
        "OUTLOOK_TENANT_ID": "common",
        "OUTLOOK_MODE": "read-only"
      }
    }
  }
}
```

On first use the server opens your browser to sign in and caches the token at `~/.mcp-outlook/token.json` (0600); later runs refresh silently.

See [docs/CLIENTS.md](docs/CLIENTS.md) for Claude Code / Cursor / Codex / VS Code / Windsurf snippets, and [.env.example](.env.example) for every supported variable.

## Authentication

Auth mode is chosen automatically (override with `OUTLOOK_AUTH`):

- **interactive** (default) — only `OUTLOOK_CLIENT_ID` set. Authorization-code + PKCE with a loopback redirect: the browser opens, you approve once, and the access + refresh token are cached on disk. Operates on the signed-in user's mailbox (`/me`). The server never handles your password.
- **client-credentials** (app-only) — `OUTLOOK_CLIENT_SECRET` present. The server fetches an app token itself; requires `OUTLOOK_TENANT_ID` and `OUTLOOK_USER` (the mailbox to act on, since an app token has no signed-in user). Grant the app **application** Mail permissions with admin consent.
- **token** — `OUTLOOK_TOKEN` set to a pre-obtained Graph bearer token. You manage its lifetime.

## Safe by default

The access model is enforced by [`src/security.ts`](src/security.ts) — defence in depth on top of the Graph token's own scopes/roles:

- **`OUTLOOK_MODE`** — `read-only` (default) → `read-write` → `admin`. A tool is registered only if the mode allows its capability. Read-only exposes the 7 read tools; drafts/moves need `read-write`; deletes need `admin`.
- **`OUTLOOK_ALLOW_SEND`** — sending mail (send / reply / forward) can't be undone, so on top of `read-write` it also requires this flag. Drafting is always allowed in read-write; nothing leaves the mailbox until sent.
- **`OUTLOOK_ALLOW_DELETE`** — deletes require this flag on top of `admin` mode.
- **`OUTLOOK_FOLDER_ALLOWLIST` / `OUTLOOK_PROTECTED_FOLDERS`** — confine which folders can be written to / moved into; mark folders (e.g. `sentitems`, `archive`) that may be read but never modified.
- **Interactive confirmation** — when the client supports MCP elicitation, sending mail and deleting a message pause and ask the **human** to approve before running; clients that can't elicit fall back to the `OUTLOOK_ALLOW_SEND` / `OUTLOOK_ALLOW_DELETE` gates.
- **`OUTLOOK_DRY_RUN`** — validate and log writes without executing them.
- **`OUTLOOK_AUDIT_LOG`** — a JSON audit line per guarded operation, on stderr (default on).
- **Attachment bytes are never returned** — `list_attachments` returns metadata only.

See [SECURITY.md](SECURITY.md).

## Developing

```bash
npm install
npm run build
# introspect the tool list without signing in (uses a fake token, no network):
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}' | OUTLOOK_TOKEN=x node dist/index.js
npm test
```

## Licence

MIT
