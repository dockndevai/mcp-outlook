#!/usr/bin/env node
/**
 * MCP server for Microsoft Outlook (Microsoft Graph mail).
 *
 * Lets an agent read and operate an Outlook / Microsoft 365 mailbox — list folders and messages,
 * full-text search, read message bodies and attachment metadata, list contacts, and (in higher
 * modes) create drafts, send/reply/forward mail, mark read, move messages, and delete.
 *
 * Auth is browser-first: on first run with only OUTLOOK_CLIENT_ID set, the server opens your
 * browser to the Microsoft sign-in page (authorization-code + PKCE), caches the token, and refreshes
 * it silently thereafter. App-only (client-credentials) and static-token modes are also supported.
 *
 * Safe by default: starts in read-only mode, so only the read tools are registered. Writes need
 * OUTLOOK_MODE=read-write; sending mail additionally needs OUTLOOK_ALLOW_SEND=true; deletes need
 * admin mode plus OUTLOOK_ALLOW_DELETE=true. The access model (src/security.ts) is defence in depth
 * over the Graph token's own scopes.
 */
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { hasCredentials, loadConfig, missingCredentialHint } from "./config.js";
import { buildServer } from "./server.js";

const config = loadConfig();

if (!hasCredentials(config.connection)) {
  process.stderr.write(`mcp-outlook: no usable credentials. ${missingCredentialHint(config.connection)}\n`);
  process.exit(1);
}

const { server, enabled } = buildServer(config);

const transport = new StdioServerTransport();
await server.connect(transport);
// stdout carries the protocol; diagnostics go to stderr.
process.stderr.write(
  `outlook-mcp connected [auth=${config.connection.auth.mode}, mailbox=${config.connection.userPath}, ` +
    `mode=${config.security.mode}, tools=${enabled.length}: ${enabled.join(", ")}]\n`,
);
