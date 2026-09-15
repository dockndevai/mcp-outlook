# Installing `mcp-outlook` in your MCP client

`mcp-outlook` is a **stdio** MCP server. Any MCP-compatible agent can run it.

- **From npm (recommended):** `npx -y @dockndevai/mcp-outlook`
- **From source:** `node /ABSOLUTE/PATH/TO/mcp-outlook/dist/index.js` after `npm install && npm run build`.

> You need an Entra (Azure AD) **app registration**. For the default browser sign-in, register a **public client** with redirect URI `http://localhost` and use its client id as `OUTLOOK_CLIENT_ID`. **Start in `read-only` mode** and raise it deliberately. See [`.env.example`](../.env.example) for every variable.

On first use the server opens your browser to sign in, then caches the token at `~/.mcp-outlook/token.json` and refreshes it silently.

## Claude Code (CLI)

```bash
claude mcp add outlook \
  -e OUTLOOK_CLIENT_ID="00000000-0000-0000-0000-000000000000" \
  -e OUTLOOK_TENANT_ID="common" \
  -e OUTLOOK_MODE="read-only" \
  -- npx -y @dockndevai/mcp-outlook
```

Add `-s user` to install it for all your projects, or `-s project` for a shared `.mcp.json`. List with `claude mcp list`, remove with `claude mcp remove outlook`.

## Claude Desktop

Edit `claude_desktop_config.json` (macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`) and merge:

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

Restart Claude Desktop. The server appears under the tools (🔨) menu.

## Cursor

Create `.cursor/mcp.json` (or `~/.cursor/mcp.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "@dockndevai/mcp-outlook"],
      "env": { "OUTLOOK_CLIENT_ID": "00000000-0000-0000-0000-000000000000", "OUTLOOK_TENANT_ID": "common", "OUTLOOK_MODE": "read-only" }
    }
  }
}
```

Then enable it in **Cursor Settings → MCP**.

## OpenAI Codex CLI

Edit `~/.codex/config.toml`:

```toml
[mcp_servers.outlook]
command = "npx"
args = ["-y", "@dockndevai/mcp-outlook"]
env = { OUTLOOK_CLIENT_ID = "00000000-0000-0000-0000-000000000000", OUTLOOK_TENANT_ID = "common", OUTLOOK_MODE = "read-only" }
```

## Windsurf

Edit `~/.codeium/windsurf/mcp_config.json` with the same `mcpServers` block as Cursor, then **Refresh** in the Windsurf MCP panel.

## VS Code (GitHub Copilot / Agent mode)

Create `.vscode/mcp.json` (top-level key is `servers`):

```json
{
  "servers": {
    "outlook": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "@dockndevai/mcp-outlook"],
      "env": { "OUTLOOK_CLIENT_ID": "00000000-0000-0000-0000-000000000000", "OUTLOOK_TENANT_ID": "common", "OUTLOOK_MODE": "read-only" }
    }
  }
}
```

> **Headless / remote hosts:** interactive sign-in needs a browser on the same machine. On a server, either run once locally to populate `~/.mcp-outlook/token.json` and copy it over, or use app-only auth (`OUTLOOK_CLIENT_SECRET` + `OUTLOOK_TENANT_ID` + `OUTLOOK_USER`).

## Verify

On startup the server logs a line to **stderr** like:

```
outlook-mcp connected [auth=interactive, mailbox=/me, mode=read-only, tools=7: whoami, list_folders, ...]
```

If credentials are missing it exits with the fix printed to stderr. Ask your agent to *"list the Outlook tools"* to confirm.
