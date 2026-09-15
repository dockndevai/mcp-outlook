/**
 * Configuration from environment variables.
 *
 * Auth is Microsoft Graph. Three ways to authenticate, chosen automatically in this priority order:
 *
 *   1. **interactive** (default for a personal / work mailbox) — no secret configured. On first use
 *      the server opens your browser to the Microsoft sign-in page, you approve once, and it caches
 *      the resulting token (with a refresh token) on disk. Subsequent runs refresh silently. This is
 *      the `az login` model: the server never sees your password. Needs only OUTLOOK_CLIENT_ID (a
 *      public-client app registration with the redirect URI http://localhost).
 *   2. **client-credentials** (app-only, for automation) — OUTLOOK_TENANT_ID + OUTLOOK_CLIENT_ID +
 *      OUTLOOK_CLIENT_SECRET. The server fetches and caches an app token itself. Requires OUTLOOK_USER.
 *   3. **static token** — OUTLOOK_TOKEN, a pre-obtained Graph bearer token. Simplest, but you manage
 *      its lifetime.
 *
 * The mailbox is `/me` for interactive/delegated tokens, or `/users/{OUTLOOK_USER}` when OUTLOOK_USER
 * is set (required for app-only, since app tokens have no signed-in user).
 */
import { homedir } from "node:os";
import { join } from "node:path";
import type { AccessMode, SecurityConfig } from "./security.js";

export type AuthMode = "interactive" | "client-credentials" | "token";

/** Default delegated scopes for interactive login. offline_access yields a refresh token. */
const DEFAULT_SCOPES = [
  "offline_access",
  "User.Read",
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "MailboxSettings.Read",
];

export interface GraphAuth {
  mode: AuthMode;
  /** Static bearer token (mode "token"). */
  token?: string;
  /** Directory tenant: a tenant id, "organizations", "consumers", or "common" (default). */
  tenantId: string;
  /** App (client) id — required for interactive and client-credentials. */
  clientId?: string;
  /** App secret (mode "client-credentials" only). */
  clientSecret?: string;
  /** Delegated scopes requested at interactive login. */
  scopes: string[];
  /** Where the cached interactive token/refresh token is stored. */
  tokenCachePath: string;
  /** Fixed loopback port for the login redirect, or 0 to pick a free one. */
  redirectPort: number;
}

export interface OutlookConnection {
  graphBaseUrl: string;
  /** OAuth authority base, e.g. https://login.microsoftonline.com */
  authorityBase: string;
  auth: GraphAuth;
  /** Mailbox path segment: "/me" or "/users/<id-or-upn>". */
  userPath: string;
  timeoutMs: number;
  maxResults: number;
}

export interface AppConfig {
  connection: OutlookConnection;
  security: SecurityConfig;
}

function bool(name: string, fallback: boolean): boolean {
  const v = process.env[name];
  if (v === undefined || v === "") return fallback;
  return ["1", "true", "yes", "on"].includes(v.toLowerCase());
}

function list(name: string): string[] {
  const v = process.env[name];
  if (!v) return [];
  return v.split(",").map((s) => s.trim()).filter(Boolean);
}

function parseMode(): AccessMode {
  const raw = (process.env.OUTLOOK_MODE ?? "read-only").toLowerCase();
  if (raw === "read-only" || raw === "read-write" || raw === "admin") return raw;
  throw new Error(`Invalid OUTLOOK_MODE '${raw}'. Expected one of: read-only, read-write, admin.`);
}

/** Pick the auth mode from what is present, honouring an explicit OUTLOOK_AUTH override. */
function resolveAuthMode(): AuthMode {
  const explicit = process.env.OUTLOOK_AUTH?.trim().toLowerCase();
  if (explicit === "interactive" || explicit === "client-credentials" || explicit === "token") {
    return explicit;
  }
  if (explicit) {
    throw new Error(`Invalid OUTLOOK_AUTH '${explicit}'. Expected one of: interactive, client-credentials, token.`);
  }
  if (process.env.OUTLOOK_TOKEN?.trim()) return "token";
  if (process.env.OUTLOOK_CLIENT_SECRET?.trim()) return "client-credentials";
  return "interactive";
}

export function loadConfig(): AppConfig {
  const user = process.env.OUTLOOK_USER?.trim();
  const scopes = list("OUTLOOK_SCOPES");
  return {
    connection: {
      graphBaseUrl: (process.env.OUTLOOK_GRAPH_BASE || "https://graph.microsoft.com/v1.0").replace(/\/$/, ""),
      authorityBase: (process.env.OUTLOOK_AUTHORITY || "https://login.microsoftonline.com").replace(/\/$/, ""),
      auth: {
        mode: resolveAuthMode(),
        token: process.env.OUTLOOK_TOKEN?.trim() || undefined,
        tenantId: process.env.OUTLOOK_TENANT_ID?.trim() || "common",
        clientId: process.env.OUTLOOK_CLIENT_ID?.trim() || undefined,
        clientSecret: process.env.OUTLOOK_CLIENT_SECRET?.trim() || undefined,
        scopes: scopes.length ? scopes : DEFAULT_SCOPES,
        tokenCachePath:
          process.env.OUTLOOK_TOKEN_CACHE?.trim() || join(homedir(), ".mcp-outlook", "token.json"),
        redirectPort: Math.max(0, Number(process.env.OUTLOOK_REDIRECT_PORT ?? 0)),
      },
      userPath: user ? `/users/${encodeURIComponent(user)}` : "/me",
      timeoutMs: Number(process.env.OUTLOOK_TIMEOUT_MS ?? 30000),
      maxResults: Math.max(1, Math.min(200, Number(process.env.OUTLOOK_MAX_RESULTS ?? 25))),
    },
    security: {
      mode: parseMode(),
      folderAllowlist: list("OUTLOOK_FOLDER_ALLOWLIST"),
      protectedFolders: list("OUTLOOK_PROTECTED_FOLDERS"),
      allowSend: bool("OUTLOOK_ALLOW_SEND", false),
      allowDelete: bool("OUTLOOK_ALLOW_DELETE", false),
      dryRun: bool("OUTLOOK_DRY_RUN", false),
      auditLog: bool("OUTLOOK_AUDIT_LOG", true),
    },
  };
}

/**
 * True when the configuration is complete enough to attempt authentication.
 * Interactive and client-credentials both need a client id; interactive can then log in on demand.
 */
export function hasCredentials(c: OutlookConnection): boolean {
  const a = c.auth;
  switch (a.mode) {
    case "token":
      return Boolean(a.token);
    case "client-credentials":
      return Boolean(a.tenantId && a.clientId && a.clientSecret);
    case "interactive":
      return Boolean(a.clientId);
  }
}

/** A human-readable hint about what is missing, for the startup diagnostic. */
export function missingCredentialHint(c: OutlookConnection): string {
  switch (c.auth.mode) {
    case "token":
      return "Set OUTLOOK_TOKEN to a Microsoft Graph bearer token.";
    case "client-credentials":
      return "Set OUTLOOK_TENANT_ID, OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET (and OUTLOOK_USER).";
    case "interactive":
      return "Set OUTLOOK_CLIENT_ID to a public-client app registration (redirect URI http://localhost) so the server can open a browser to sign you in.";
  }
}
