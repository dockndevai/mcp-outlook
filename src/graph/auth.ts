/**
 * Microsoft Graph token acquisition.
 *
 * A `TokenProvider` hands the HTTP client a fresh bearer token on demand. Three implementations,
 * one per configured auth mode:
 *
 *   - **StaticTokenProvider** — returns OUTLOOK_TOKEN verbatim.
 *   - **ClientCredentialsProvider** — app-only; POSTs client_credentials to the token endpoint and
 *     caches the token in memory until shortly before it expires.
 *   - **InteractiveProvider** — authorization-code + PKCE with a loopback redirect. On first use it
 *     opens the system browser to the Microsoft sign-in page, runs a one-request localhost server to
 *     catch the redirect, exchanges the code for an access + refresh token, and caches both to disk
 *     (0600). Later runs load the cache and refresh silently; the browser only reappears if the
 *     refresh token is missing or rejected.
 *
 * The server never handles the user's password — sign-in happens entirely in the browser against
 * Microsoft. Tokens are secrets: the disk cache is written user-only and never logged.
 */
import { spawn } from "node:child_process";
import { createHash, randomBytes } from "node:crypto";
import { chmodSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { createServer } from "node:http";
import { dirname } from "node:path";
import { AddressInfo } from "node:net";
import type { GraphAuth } from "../config.js";

export class AuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AuthError";
  }
}

export interface TokenProvider {
  /** A valid bearer token, refreshing or signing in as needed. */
  getToken(): Promise<string>;
}

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  token_type?: string;
  error?: string;
  error_description?: string;
}

interface CachedToken {
  accessToken: string;
  refreshToken?: string;
  /** Epoch ms when the access token expires. */
  expiresAt: number;
  scopes: string[];
}

const EXPIRY_SKEW_MS = 60_000; // refresh a minute early

function base64url(buf: Buffer): string {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function tokenEndpoint(auth: GraphAuth, authorityBase: string): string {
  return `${authorityBase}/${encodeURIComponent(auth.tenantId)}/oauth2/v2.0/token`;
}
function authorizeEndpoint(auth: GraphAuth, authorityBase: string): string {
  return `${authorityBase}/${encodeURIComponent(auth.tenantId)}/oauth2/v2.0/authorize`;
}

/** Open a URL in the user's default browser (best effort; login can also be completed manually). */
function openBrowser(url: string): void {
  const platform = process.platform;
  const cmd = platform === "darwin" ? "open" : platform === "win32" ? "cmd" : "xdg-open";
  const args = platform === "win32" ? ["/c", "start", "", url] : [url];
  try {
    const child = spawn(cmd, args, { stdio: "ignore", detached: true });
    child.on("error", () => {});
    child.unref();
  } catch {
    /* if we can't launch a browser, the URL is printed to stderr for the user to open manually */
  }
}

async function postForm(url: string, params: Record<string, string>, timeoutMs: number): Promise<TokenResponse> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: "POST",
      signal: controller.signal,
      headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
      body: new URLSearchParams(params).toString(),
    });
    const body = (await res.json().catch(() => ({}))) as TokenResponse;
    if (!res.ok || body.error) {
      throw new AuthError(
        `Token request failed (${res.status}): ${body.error ?? res.statusText}${
          body.error_description ? ` — ${body.error_description}` : ""
        }`,
      );
    }
    return body;
  } catch (err) {
    if (err instanceof AuthError) throw err;
    if (err instanceof Error && err.name === "AbortError") {
      throw new AuthError(`Token endpoint did not respond within ${timeoutMs / 1000}s.`);
    }
    throw new AuthError(err instanceof Error ? err.message : String(err));
  } finally {
    clearTimeout(timer);
  }
}

class StaticTokenProvider implements TokenProvider {
  constructor(private readonly token: string) {}
  async getToken(): Promise<string> {
    return this.token;
  }
}

class ClientCredentialsProvider implements TokenProvider {
  private cached?: { token: string; expiresAt: number };
  constructor(
    private readonly auth: GraphAuth,
    private readonly authorityBase: string,
    private readonly timeoutMs: number,
  ) {}

  async getToken(): Promise<string> {
    if (this.cached && this.cached.expiresAt - EXPIRY_SKEW_MS > Date.now()) return this.cached.token;
    const body = await postForm(
      tokenEndpoint(this.auth, this.authorityBase),
      {
        grant_type: "client_credentials",
        client_id: this.auth.clientId!,
        client_secret: this.auth.clientSecret!,
        scope: "https://graph.microsoft.com/.default",
      },
      this.timeoutMs,
    );
    this.cached = { token: body.access_token, expiresAt: Date.now() + (body.expires_in ?? 3600) * 1000 };
    return this.cached.token;
  }
}

class InteractiveProvider implements TokenProvider {
  private mem?: CachedToken;
  /** De-dupe concurrent sign-ins/refreshes so the browser opens at most once. */
  private inflight?: Promise<string>;

  constructor(
    private readonly auth: GraphAuth,
    private readonly authorityBase: string,
    private readonly timeoutMs: number,
  ) {}

  async getToken(): Promise<string> {
    const cached = this.mem ?? this.load();
    if (cached && cached.expiresAt - EXPIRY_SKEW_MS > Date.now()) {
      this.mem = cached;
      return cached.accessToken;
    }
    if (this.inflight) return this.inflight;
    this.inflight = this.acquire(cached).finally(() => {
      this.inflight = undefined;
    });
    return this.inflight;
  }

  private async acquire(cached?: CachedToken): Promise<string> {
    if (cached?.refreshToken) {
      try {
        return await this.refresh(cached.refreshToken);
      } catch {
        // refresh token expired/revoked — fall through to a fresh interactive login
      }
    }
    return this.login();
  }

  private async refresh(refreshToken: string): Promise<string> {
    const body = await postForm(
      tokenEndpoint(this.auth, this.authorityBase),
      {
        grant_type: "refresh_token",
        client_id: this.auth.clientId!,
        refresh_token: refreshToken,
        scope: this.auth.scopes.join(" "),
      },
      this.timeoutMs,
    );
    return this.persist(body);
  }

  private async login(): Promise<string> {
    const verifier = base64url(randomBytes(32));
    const challenge = base64url(createHash("sha256").update(verifier).digest());
    const state = base64url(randomBytes(16));

    const { code, redirectUri } = await this.runLoopback(challenge, state);
    const body = await postForm(
      tokenEndpoint(this.auth, this.authorityBase),
      {
        grant_type: "authorization_code",
        client_id: this.auth.clientId!,
        code,
        redirect_uri: redirectUri,
        code_verifier: verifier,
        scope: this.auth.scopes.join(" "),
      },
      this.timeoutMs,
    );
    return this.persist(body);
  }

  /** Start a one-shot localhost server, open the browser, and resolve with the returned auth code. */
  private runLoopback(challenge: string, state: string): Promise<{ code: string; redirectUri: string }> {
    return new Promise((resolve, reject) => {
      const server = createServer((req, res) => {
        const url = new URL(req.url ?? "/", "http://localhost");
        if (url.pathname !== "/") {
          res.writeHead(404).end();
          return;
        }
        const err = url.searchParams.get("error");
        const code = url.searchParams.get("code");
        const returnedState = url.searchParams.get("state");
        res.writeHead(200, { "Content-Type": "text/html" });
        if (err) {
          res.end(page("Sign-in failed", `${err}: ${url.searchParams.get("error_description") ?? ""}`));
          cleanup();
          reject(new AuthError(`Sign-in failed: ${err}`));
        } else if (!code || returnedState !== state) {
          res.end(page("Sign-in failed", "Missing code or state mismatch."));
          cleanup();
          reject(new AuthError("Sign-in failed: missing authorization code or state mismatch."));
        } else {
          res.end(page("Signed in", "You can close this tab and return to your terminal."));
          cleanup();
          resolve({ code, redirectUri });
        }
      });

      let redirectUri = "";
      const timer = setTimeout(() => {
        cleanup();
        reject(new AuthError("Timed out waiting for browser sign-in (5 minutes)."));
      }, 5 * 60_000);

      function cleanup() {
        clearTimeout(timer);
        server.close();
      }

      server.on("error", (e) => {
        clearTimeout(timer);
        reject(new AuthError(`Could not start the local sign-in listener: ${e.message}`));
      });

      server.listen(this.auth.redirectPort, "127.0.0.1", () => {
        const { port } = server.address() as AddressInfo;
        redirectUri = `http://localhost:${port}`;
        const authorizeUrl =
          `${authorizeEndpoint(this.auth, this.authorityBase)}?` +
          new URLSearchParams({
            client_id: this.auth.clientId!,
            response_type: "code",
            redirect_uri: redirectUri,
            response_mode: "query",
            scope: this.auth.scopes.join(" "),
            state,
            code_challenge: challenge,
            code_challenge_method: "S256",
            prompt: "select_account",
          }).toString();
        process.stderr.write(`[mcp-outlook] Opening browser to sign in. If it does not open, visit:\n${authorizeUrl}\n`);
        openBrowser(authorizeUrl);
      });
    });
  }

  private persist(body: TokenResponse): string {
    const cached: CachedToken = {
      accessToken: body.access_token,
      refreshToken: body.refresh_token ?? this.mem?.refreshToken,
      expiresAt: Date.now() + (body.expires_in ?? 3600) * 1000,
      scopes: this.auth.scopes,
    };
    this.mem = cached;
    this.save(cached);
    return cached.accessToken;
  }

  private load(): CachedToken | undefined {
    try {
      const raw = readFileSync(this.auth.tokenCachePath, "utf8");
      const parsed = JSON.parse(raw) as CachedToken;
      if (parsed && typeof parsed.accessToken === "string") return parsed;
    } catch {
      /* no cache yet, or unreadable — sign in fresh */
    }
    return undefined;
  }

  private save(cached: CachedToken): void {
    try {
      mkdirSync(dirname(this.auth.tokenCachePath), { recursive: true, mode: 0o700 });
      writeFileSync(this.auth.tokenCachePath, JSON.stringify(cached), { mode: 0o600 });
      chmodSync(this.auth.tokenCachePath, 0o600);
    } catch (e) {
      process.stderr.write(
        `[mcp-outlook] Warning: could not write token cache at ${this.auth.tokenCachePath}: ${
          e instanceof Error ? e.message : String(e)
        }\n`,
      );
    }
  }
}

function page(title: string, body: string): string {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title><style>body{font-family:system-ui,sans-serif;padding:3rem;text-align:center;color:#222}h1{font-size:1.4rem}</style></head><body><h1>${title}</h1><p>${body}</p></body></html>`;
}

/** Build the right TokenProvider for the configured auth mode. */
export function createTokenProvider(auth: GraphAuth, authorityBase: string, timeoutMs: number): TokenProvider {
  switch (auth.mode) {
    case "token":
      if (!auth.token) throw new AuthError("OUTLOOK_TOKEN is not set.");
      return new StaticTokenProvider(auth.token);
    case "client-credentials":
      if (!auth.clientId || !auth.clientSecret) {
        throw new AuthError("Client-credentials mode needs OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET.");
      }
      return new ClientCredentialsProvider(auth, authorityBase, timeoutMs);
    case "interactive":
      if (!auth.clientId) throw new AuthError("Interactive mode needs OUTLOOK_CLIENT_ID (a public-client app).");
      return new InteractiveProvider(auth, authorityBase, timeoutMs);
  }
}
