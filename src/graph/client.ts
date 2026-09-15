/**
 * Thin HTTP client for Microsoft Graph (Outlook mail).
 *
 * Tokens come from a `TokenProvider` (interactive browser login, app-only client-credentials, or a
 * static bearer token) — see auth.ts. Every request asks the provider for a fresh token, so silent
 * refresh is transparent here.
 *
 * Attachment bytes (`contentBytes`) are large and often sensitive; `redactAttachment` strips them so
 * only metadata reaches the model. Fetch the raw bytes explicitly if ever needed.
 */
import type { OutlookConnection } from "../config.js";
import { createTokenProvider, type TokenProvider } from "./auth.js";

export class GraphError extends Error {
  constructor(
    readonly status: number,
    message: string,
    readonly body?: string,
  ) {
    super(message);
    this.name = "GraphError";
  }
}

export interface MessageRef {
  id: string;
  subject?: string;
  from?: string;
  to?: string[];
  received?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  preview?: string;
  webLink?: string;
}

export interface MailFolderRef {
  id: string;
  displayName: string;
  unreadItemCount?: number;
  totalItemCount?: number;
  childFolderCount?: number;
}

export interface AttachmentMeta {
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
}

type GraphRecipient = { emailAddress?: { name?: string; address?: string } };
type GraphMessage = {
  id: string;
  subject?: string;
  from?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  bodyPreview?: string;
  webLink?: string;
};

function addr(r?: GraphRecipient): string | undefined {
  return r?.emailAddress?.address;
}

/** Compact a Graph message down to the fields worth showing. */
export function summarizeMessage(m: GraphMessage): MessageRef {
  return {
    id: m.id,
    subject: m.subject,
    from: addr(m.from),
    to: (m.toRecipients ?? []).map(addr).filter((x): x is string => Boolean(x)),
    received: m.receivedDateTime,
    isRead: m.isRead,
    hasAttachments: m.hasAttachments,
    preview: m.bodyPreview,
    webLink: m.webLink,
  };
}

/** Drop attachment bytes; keep metadata only. */
export function redactAttachment(a: Record<string, unknown>): AttachmentMeta {
  return {
    id: a.id as string,
    name: a.name as string | undefined,
    contentType: a.contentType as string | undefined,
    size: a.size as number | undefined,
    isInline: a.isInline as boolean | undefined,
  };
}

/** A recipient address, optionally with a display name. */
export interface Recipient {
  address: string;
  name?: string;
}
function toRecipient(r: Recipient): GraphRecipient {
  return { emailAddress: { address: r.address, ...(r.name ? { name: r.name } : {}) } };
}

export interface DraftMessage {
  subject: string;
  body: string;
  contentType?: "Text" | "HTML";
  to: Recipient[];
  cc?: Recipient[];
  bcc?: Recipient[];
}

function buildMessage(m: DraftMessage): Record<string, unknown> {
  return {
    subject: m.subject,
    body: { contentType: m.contentType ?? "Text", content: m.body },
    toRecipients: m.to.map(toRecipient),
    ...(m.cc?.length ? { ccRecipients: m.cc.map(toRecipient) } : {}),
    ...(m.bcc?.length ? { bccRecipients: m.bcc.map(toRecipient) } : {}),
  };
}

export class GraphClient {
  private readonly base: string;
  private readonly userPath: string;
  private readonly timeoutMs: number;
  readonly maxResults: number;
  private readonly tokens: TokenProvider;

  constructor(conn: OutlookConnection) {
    this.base = conn.graphBaseUrl;
    this.userPath = conn.userPath;
    this.timeoutMs = conn.timeoutMs;
    this.maxResults = conn.maxResults;
    this.tokens = createTokenProvider(conn.auth, conn.authorityBase, conn.timeoutMs);
  }

  /** Force auth to resolve (interactive login happens here on first run). */
  async ensureAuth(): Promise<void> {
    await this.tokens.getToken();
  }

  private async request<T>(path: string, init: RequestInit = {}): Promise<T> {
    const token = await this.tokens.getToken();
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const res = await fetch(path.startsWith("http") ? path : `${this.base}${path}`, {
        ...init,
        signal: controller.signal,
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          ...(init.body ? { "Content-Type": "application/json" } : {}),
          ...(init.headers ?? {}),
        },
      });
      if (!res.ok) {
        let detail = res.statusText;
        let body: string | undefined;
        try {
          body = await res.text();
          const parsed = JSON.parse(body) as { error?: { message?: string } };
          detail = parsed.error?.message ?? detail;
        } catch {
          /* non-JSON error body */
        }
        throw new GraphError(res.status, detail, body);
      }
      if (res.status === 204) return undefined as T;
      const text = await res.text();
      return (text ? JSON.parse(text) : undefined) as T;
    } catch (err) {
      if (err instanceof GraphError) throw err;
      if (err instanceof Error && err.name === "AbortError") {
        throw new GraphError(504, `Microsoft Graph did not respond within ${this.timeoutMs / 1000}s.`);
      }
      throw new GraphError(0, err instanceof Error ? err.message : String(err));
    } finally {
      clearTimeout(timer);
    }
  }

  private mailbox(suffix: string): string {
    return `${this.userPath}${suffix}`;
  }

  // --- Identity ---
  me() {
    return this.request<Record<string, unknown>>(this.mailbox(""));
  }

  // --- Folders ---
  async listFolders(): Promise<MailFolderRef[]> {
    const data = await this.request<{ value: MailFolderRef[] }>(
      this.mailbox("/mailFolders?$top=100&$select=id,displayName,unreadItemCount,totalItemCount,childFolderCount"),
    );
    return data.value ?? [];
  }

  // --- Messages ---
  private selectFields = "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,webLink";

  async listMessages(opts: { folder?: string; top?: number; unreadOnly?: boolean }): Promise<MessageRef[]> {
    const top = Math.min(this.maxResults, opts.top ?? this.maxResults);
    const seg = opts.folder ? `/mailFolders/${encodeURIComponent(opts.folder)}/messages` : "/messages";
    const q = new URLSearchParams({
      $top: String(top),
      $select: this.selectFields,
      $orderby: "receivedDateTime desc",
    });
    if (opts.unreadOnly) q.set("$filter", "isRead eq false");
    const data = await this.request<{ value: GraphMessage[] }>(this.mailbox(`${seg}?${q.toString()}`));
    return (data.value ?? []).map(summarizeMessage);
  }

  async searchMessages(query: string, top?: number): Promise<MessageRef[]> {
    const n = Math.min(this.maxResults, top ?? this.maxResults);
    // $search must be double-quoted and cannot be combined with $orderby.
    const q = new URLSearchParams({ $top: String(n), $select: this.selectFields });
    q.set("$search", `"${query.replace(/"/g, '\\"')}"`);
    const data = await this.request<{ value: GraphMessage[] }>(this.mailbox(`/messages?${q.toString()}`));
    return (data.value ?? []).map(summarizeMessage);
  }

  async getMessage(id: string): Promise<Record<string, unknown>> {
    const q = new URLSearchParams({
      $select: `${this.selectFields},body,ccRecipients,bccRecipients,importance,conversationId`,
    });
    return this.request<Record<string, unknown>>(this.mailbox(`/messages/${encodeURIComponent(id)}?${q.toString()}`));
  }

  async listAttachments(id: string): Promise<AttachmentMeta[]> {
    const data = await this.request<{ value: Record<string, unknown>[] }>(
      this.mailbox(`/messages/${encodeURIComponent(id)}/attachments?$select=id,name,contentType,size,isInline`),
    );
    return (data.value ?? []).map(redactAttachment);
  }

  // --- Contacts ---
  async listContacts(top?: number): Promise<Array<Record<string, unknown>>> {
    const n = Math.min(this.maxResults, top ?? this.maxResults);
    const data = await this.request<{ value: Array<Record<string, unknown>> }>(
      this.mailbox(`/contacts?$top=${n}&$select=id,displayName,emailAddresses,companyName`),
    );
    return data.value ?? [];
  }

  // --- Writes ---
  createDraft(msg: DraftMessage) {
    return this.request<Record<string, unknown>>(this.mailbox("/messages"), {
      method: "POST",
      body: JSON.stringify(buildMessage(msg)),
    });
  }

  sendMail(msg: DraftMessage, saveToSentItems = true) {
    return this.request<void>(this.mailbox("/sendMail"), {
      method: "POST",
      body: JSON.stringify({ message: buildMessage(msg), saveToSentItems }),
    });
  }

  replyMessage(id: string, comment: string, replyAll = false) {
    const action = replyAll ? "replyAll" : "reply";
    return this.request<void>(this.mailbox(`/messages/${encodeURIComponent(id)}/${action}`), {
      method: "POST",
      body: JSON.stringify({ comment }),
    });
  }

  forwardMessage(id: string, to: Recipient[], comment = "") {
    return this.request<void>(this.mailbox(`/messages/${encodeURIComponent(id)}/forward`), {
      method: "POST",
      body: JSON.stringify({ comment, toRecipients: to.map(toRecipient) }),
    });
  }

  markRead(id: string, isRead: boolean) {
    return this.request<Record<string, unknown>>(this.mailbox(`/messages/${encodeURIComponent(id)}`), {
      method: "PATCH",
      body: JSON.stringify({ isRead }),
    });
  }

  moveMessage(id: string, destinationId: string) {
    return this.request<Record<string, unknown>>(this.mailbox(`/messages/${encodeURIComponent(id)}/move`), {
      method: "POST",
      body: JSON.stringify({ destinationId }),
    });
  }

  // --- Destructive ---
  deleteMessage(id: string) {
    // Graph DELETE on a message moves it to Deleted Items (recoverable), which is the safe default.
    return this.request<void>(this.mailbox(`/messages/${encodeURIComponent(id)}`), { method: "DELETE" });
  }
}
