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
import { createHash } from "node:crypto";
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

export interface DriveItemRef {
  id: string;
  name?: string;
  isFolder: boolean;
  size?: number;
  childCount?: number;
  lastModified?: string;
  webUrl?: string;
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

// --- Calendar ---

export interface CalendarRef {
  id: string;
  name?: string;
  owner?: string;
  canEdit?: boolean;
  isDefault?: boolean;
}

export interface EventRef {
  id: string;
  subject?: string;
  start?: string;
  end?: string;
  timeZone?: string;
  location?: string;
  organizer?: string;
  attendees?: string[];
  isAllDay?: boolean;
  isCancelled?: boolean;
  isOnlineMeeting?: boolean;
  joinUrl?: string;
  myResponse?: string;
  webLink?: string;
}

/** Fields a create/update accepts. Times are ISO 8601; timeZone is an IANA/Windows zone (default UTC). */
export interface EventInput {
  subject?: string;
  start?: string;
  end?: string;
  timeZone?: string;
  body?: string;
  html?: boolean;
  location?: string;
  attendees?: string[];
  isAllDay?: boolean;
  isOnlineMeeting?: boolean;
  reminderMinutesBeforeStart?: number;
}

type GraphDateTimeTz = { dateTime?: string; timeZone?: string };
type GraphEvent = {
  id: string;
  subject?: string;
  start?: GraphDateTimeTz;
  end?: GraphDateTimeTz;
  location?: { displayName?: string };
  organizer?: GraphRecipient;
  attendees?: Array<{ emailAddress?: { address?: string } }>;
  isAllDay?: boolean;
  isCancelled?: boolean;
  isOnlineMeeting?: boolean;
  onlineMeeting?: { joinUrl?: string } | null;
  responseStatus?: { response?: string };
  webLink?: string;
};

export function summarizeEvent(e: GraphEvent): EventRef {
  return {
    id: e.id,
    subject: e.subject,
    start: e.start?.dateTime,
    end: e.end?.dateTime,
    timeZone: e.start?.timeZone,
    location: e.location?.displayName,
    organizer: e.organizer?.emailAddress?.address,
    attendees: (e.attendees ?? []).map((a) => a.emailAddress?.address).filter((x): x is string => Boolean(x)),
    isAllDay: e.isAllDay,
    isCancelled: e.isCancelled,
    isOnlineMeeting: e.isOnlineMeeting,
    joinUrl: e.onlineMeeting?.joinUrl,
    myResponse: e.responseStatus?.response,
    webLink: e.webLink,
  };
}

function buildEvent(ev: EventInput): Record<string, unknown> {
  const tz = ev.timeZone ?? "UTC";
  const out: Record<string, unknown> = {};
  if (ev.subject !== undefined) out.subject = ev.subject;
  if (ev.body !== undefined) out.body = { contentType: ev.html ? "HTML" : "Text", content: ev.body };
  if (ev.start !== undefined) out.start = { dateTime: ev.start, timeZone: tz };
  if (ev.end !== undefined) out.end = { dateTime: ev.end, timeZone: tz };
  if (ev.location !== undefined) out.location = { displayName: ev.location };
  if (ev.attendees?.length) {
    out.attendees = ev.attendees.map((address) => ({ emailAddress: { address }, type: "required" }));
  }
  if (ev.isAllDay !== undefined) out.isAllDay = ev.isAllDay;
  if (ev.isOnlineMeeting !== undefined) out.isOnlineMeeting = ev.isOnlineMeeting;
  if (ev.reminderMinutesBeforeStart !== undefined) out.reminderMinutesBeforeStart = ev.reminderMinutesBeforeStart;
  return out;
}

/** Short, stable digest of an object id + its change marker, safe to show a human and to log. */
export function versionToken(id: string, marker: string | undefined): string {
  return createHash("sha256").update(`${id}|${marker ?? ""}`).digest("hex").slice(0, 12);
}

/** Context about a message, shown to the human at approval time (display only, not part of any version). */
export interface MessageContext {
  subject?: string;
  from?: string;
  to?: string[];
  cc?: string[];
  received?: string;
}

/** Context about a calendar event, shown to the human at approval time. */
export interface EventContext {
  subject?: string;
  start?: string;
  end?: string;
  organizer?: string;
  attendees?: string[];
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

  /** GET that maps 404 to `undefined` (the object does not exist) and rethrows anything else. */
  private async getOrUndefined<T>(path: string): Promise<T | undefined> {
    try {
      return await this.request<T>(path);
    } catch (err) {
      if (err instanceof GraphError && err.status === 404) return undefined;
      throw err;
    }
  }

  /**
   * Version token for a mail message, pinned across an approval prompt: the message id plus Graph's
   * `changeKey` (falling back to `@odata.etag`). `changeKey` changes whenever the message is modified,
   * including read-state, so opening the message in another client while the prompt is open makes the
   * approval fail closed (re-ask). If Graph returns neither marker the token only proves the id still exists.
   * `undefined` when the message does not exist; transport errors throw.
   */
  async messageVersion(id: string): Promise<string | undefined> {
    const m = await this.getOrUndefined<Record<string, unknown>>(
      this.mailbox(`/messages/${encodeURIComponent(id)}?$select=id,changeKey`),
    );
    if (!m) return undefined;
    return versionToken(String(m.id ?? id), (m.changeKey as string | undefined) ?? (m["@odata.etag"] as string | undefined));
  }

  /**
   * Version token for a calendar event: id + `changeKey` (falling back to `@odata.etag`). It also moves
   * when an attendee responds or the series is touched, so a busy meeting can fail closed spuriously;
   * it never fails open. `undefined` when the event does not exist.
   */
  async eventVersion(id: string): Promise<string | undefined> {
    const e = await this.getOrUndefined<Record<string, unknown>>(
      this.mailbox(`/events/${encodeURIComponent(id)}?$select=id,changeKey`),
    );
    if (!e) return undefined;
    return versionToken(String(e.id ?? id), (e.changeKey as string | undefined) ?? (e["@odata.etag"] as string | undefined));
  }

  /**
   * Version token for a OneDrive item: id + `eTag` (metadata/name/parent changes; `cTag` is content-only
   * and ignored). Whether adding children to a folder moves a folder's `eTag` is not verified against live
   * Graph, so a folder whose contents changed while the prompt was open may not be detected.
   * `undefined` when the item does not exist.
   */
  async driveItemVersion(id: string): Promise<string | undefined> {
    const i = await this.getOrUndefined<Record<string, unknown>>(
      this.mailbox(`/drive/items/${encodeURIComponent(id)}?$select=id,eTag`),
    );
    if (!i) return undefined;
    return versionToken(String(i.id ?? id), (i.eTag as string | undefined) ?? (i["@odata.etag"] as string | undefined));
  }

  /** Best-effort display context for a message id (who/what the human is acting on). Never throws. */
  async messageContext(id: string): Promise<MessageContext | undefined> {
    try {
      const q = new URLSearchParams({ $select: "subject,from,toRecipients,ccRecipients,receivedDateTime" });
      const m = await this.getOrUndefined<GraphMessage & { ccRecipients?: GraphRecipient[] }>(
        this.mailbox(`/messages/${encodeURIComponent(id)}?${q.toString()}`),
      );
      if (!m) return undefined;
      const list = (rs?: GraphRecipient[]) => (rs ?? []).map(addr).filter((x): x is string => Boolean(x));
      return { subject: m.subject, from: addr(m.from), to: list(m.toRecipients), cc: list(m.ccRecipients), received: m.receivedDateTime };
    } catch {
      return undefined;
    }
  }

  /** Best-effort display context for an event id. Never throws. */
  async eventContext(id: string): Promise<EventContext | undefined> {
    try {
      const e = await this.getOrUndefined<GraphEvent>(
        this.mailbox(`/events/${encodeURIComponent(id)}?$select=subject,start,end,organizer,attendees`),
      );
      if (!e) return undefined;
      const s = summarizeEvent(e);
      return { subject: s.subject, start: s.start, end: s.end, organizer: s.organizer, attendees: s.attendees };
    } catch {
      return undefined;
    }
  }

  /** Best-effort display context for a OneDrive item. Never throws. */
  async driveItemContext(id: string): Promise<DriveItemRef | undefined> {
    try {
      return await this.getDriveItem({ id });
    } catch {
      return undefined;
    }
  }

  /**
   * The item currently at a OneDrive path together with its version token (id + `eTag`), or `undefined`
   * when nothing exists there (404 = a new file). Used to pin an overwrite across an approval prompt.
   * Transport/permission errors throw, so a caller that cannot tell "absent" from "unreadable" fails closed.
   */
  async driveItemAtPath(path: string): Promise<{ item: DriveItemRef; version: string } | undefined> {
    const q = "$select=id,name,size,folder,file,lastModifiedDateTime,webUrl,eTag";
    const i = await this.getOrUndefined<Record<string, unknown>>(this.mailbox(`${this.driveRef({ path })}?${q}`));
    if (!i) return undefined;
    return {
      item: this.summarizeDriveItem(i),
      version: versionToken(String(i.id ?? ""), (i.eTag as string | undefined) ?? (i["@odata.etag"] as string | undefined)),
    };
  }

  /** Best-effort path (from the drive root) of a OneDrive item, for approval prompts. Never throws. */
  async driveItemPath(id: string): Promise<string | undefined> {
    try {
      const i = await this.getOrUndefined<{ name?: string; parentReference?: { path?: string } }>(
        this.mailbox(`/drive/items/${encodeURIComponent(id)}?$select=id,name,parentReference`),
      );
      if (!i) return undefined;
      // parentReference.path looks like "/drive/root:/Documents" ("/drives/<id>/root:/..." on other drives);
      // the drive's root folder itself has no parentReference.path.
      const pp = i.parentReference?.path;
      if (pp === undefined) return "/";
      return `${pp.replace(/^.*?root:/, "")}/${i.name ?? ""}`;
    } catch {
      return undefined;
    }
  }

  /** Best-effort display name of a mail folder (id or well-known name). Never throws. */
  async mailFolderName(folder: string): Promise<string | undefined> {
    try {
      const f = await this.getOrUndefined<{ displayName?: string }>(
        this.mailbox(`/mailFolders/${encodeURIComponent(folder)}?$select=displayName`),
      );
      return f?.displayName;
    } catch {
      return undefined;
    }
  }

  /** Best-effort display name of the folder a message currently sits in. Never throws. */
  async messageFolderName(id: string): Promise<string | undefined> {
    try {
      const m = await this.getOrUndefined<{ parentFolderId?: string }>(
        this.mailbox(`/messages/${encodeURIComponent(id)}?$select=parentFolderId`),
      );
      if (!m?.parentFolderId) return undefined;
      return (await this.mailFolderName(m.parentFolderId)) ?? m.parentFolderId;
    } catch {
      return undefined;
    }
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

  // --- Calendar (reads) ---
  private eventSelect =
    "id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,isOnlineMeeting,onlineMeeting,responseStatus,webLink";

  async listCalendars(): Promise<CalendarRef[]> {
    const data = await this.request<{ value: Array<Record<string, unknown>> }>(
      this.mailbox("/calendars?$select=id,name,owner,canEdit,isDefaultCalendar"),
    );
    return (data.value ?? []).map((c) => ({
      id: c.id as string,
      name: c.name as string | undefined,
      owner: (c.owner as { address?: string } | undefined)?.address,
      canEdit: c.canEdit as boolean | undefined,
      isDefault: c.isDefaultCalendar as boolean | undefined,
    }));
  }

  async listEvents(opts: { start?: string; end?: string; calendarId?: string; top?: number }): Promise<EventRef[]> {
    const top = Math.min(this.maxResults, opts.top ?? this.maxResults);
    const calSeg = opts.calendarId ? `/calendars/${encodeURIComponent(opts.calendarId)}` : "";
    // With a window, calendarView expands recurrences; otherwise list raw events.
    if (opts.start && opts.end) {
      const q = new URLSearchParams({
        startDateTime: opts.start,
        endDateTime: opts.end,
        $top: String(top),
        $select: this.eventSelect,
        $orderby: "start/dateTime",
      });
      const data = await this.request<{ value: GraphEvent[] }>(this.mailbox(`${calSeg}/calendarView?${q.toString()}`));
      return (data.value ?? []).map(summarizeEvent);
    }
    const q = new URLSearchParams({ $top: String(top), $select: this.eventSelect, $orderby: "start/dateTime" });
    const data = await this.request<{ value: GraphEvent[] }>(this.mailbox(`${calSeg}/events?${q.toString()}`));
    return (data.value ?? []).map(summarizeEvent);
  }

  getEvent(id: string): Promise<Record<string, unknown>> {
    const q = new URLSearchParams({ $select: `${this.eventSelect},body,bodyPreview,recurrence,showAs,importance` });
    return this.request<Record<string, unknown>>(this.mailbox(`/events/${encodeURIComponent(id)}?${q.toString()}`));
  }

  getSchedule(schedules: string[], start: string, end: string, intervalMinutes = 30): Promise<Record<string, unknown>> {
    return this.request<Record<string, unknown>>(this.mailbox("/calendar/getSchedule"), {
      method: "POST",
      body: JSON.stringify({
        schedules,
        startTime: { dateTime: start, timeZone: "UTC" },
        endTime: { dateTime: end, timeZone: "UTC" },
        availabilityViewInterval: intervalMinutes,
      }),
    });
  }

  // --- Calendar (writes) ---
  createEvent(ev: EventInput): Promise<Record<string, unknown>> {
    return this.request<Record<string, unknown>>(this.mailbox("/events"), {
      method: "POST",
      body: JSON.stringify(buildEvent(ev)),
    });
  }

  updateEvent(id: string, ev: EventInput): Promise<Record<string, unknown>> {
    return this.request<Record<string, unknown>>(this.mailbox(`/events/${encodeURIComponent(id)}`), {
      method: "PATCH",
      body: JSON.stringify(buildEvent(ev)),
    });
  }

  respondToEvent(
    id: string,
    response: "accept" | "decline" | "tentativelyAccept",
    comment = "",
    sendResponse = true,
  ): Promise<void> {
    return this.request<void>(this.mailbox(`/events/${encodeURIComponent(id)}/${response}`), {
      method: "POST",
      body: JSON.stringify({ comment, sendResponse }),
    });
  }

  cancelEvent(id: string, comment = ""): Promise<void> {
    // Organizer-only: notifies attendees and cancels the meeting.
    return this.request<void>(this.mailbox(`/events/${encodeURIComponent(id)}/cancel`), {
      method: "POST",
      body: JSON.stringify({ comment }),
    });
  }

  deleteEvent(id: string): Promise<void> {
    // Removes the event; if you organize it, Graph sends cancellations to attendees.
    return this.request<void>(this.mailbox(`/events/${encodeURIComponent(id)}`), { method: "DELETE" });
  }

  // ===== OneDrive files =====

  /** Address a drive item by its id, or by a path relative to the drive root ("/Documents/x.txt"). */
  private driveRef(ref: { id?: string; path?: string }): string {
    if (ref.id) return `/drive/items/${encodeURIComponent(ref.id)}`;
    const rel = (ref.path ?? "").replace(/^\/+/, "");
    if (!rel) return "/drive/root";
    return `/drive/root:/${rel.split("/").map(encodeURIComponent).join("/")}:`;
  }

  private summarizeDriveItem(i: Record<string, unknown>): DriveItemRef {
    const folder = i.folder as { childCount?: number } | undefined;
    return {
      id: i.id as string,
      name: i.name as string | undefined,
      isFolder: Boolean(folder),
      size: i.size as number | undefined,
      childCount: folder?.childCount,
      lastModified: (i.lastModifiedDateTime as string) ?? undefined,
      webUrl: i.webUrl as string | undefined,
    };
  }

  async listDriveItems(ref: { id?: string; path?: string }): Promise<DriveItemRef[]> {
    const seg = this.driveRef(ref);
    const q = "$select=id,name,size,folder,file,lastModifiedDateTime,webUrl&$top=200";
    const data = await this.request<{ value: Record<string, unknown>[] }>(this.mailbox(`${seg}/children?${q}`));
    return (data.value ?? []).map((i) => this.summarizeDriveItem(i));
  }

  async getDriveItem(ref: { id?: string; path?: string }): Promise<DriveItemRef> {
    const i = await this.request<Record<string, unknown>>(this.mailbox(this.driveRef(ref)));
    return this.summarizeDriveItem(i);
  }

  async searchDrive(query: string, top: number): Promise<DriveItemRef[]> {
    const data = await this.request<{ value: Record<string, unknown>[] }>(
      this.mailbox(`/drive/root/search(q='${encodeURIComponent(query.replace(/'/g, "''"))}')?$top=${top}`),
    );
    return (data.value ?? []).map((i) => this.summarizeDriveItem(i));
  }

  /** Download a file's content as text (capped). Binary files come back as best-effort UTF-8. */
  async downloadFile(ref: { id?: string; path?: string }, maxBytes: number): Promise<{ bytes: number; truncated: boolean; content: string }> {
    const token = await this.tokens.getToken();
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const res = await fetch(`${this.base}${this.mailbox(this.driveRef(ref))}/content`, {
        signal: controller.signal,
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!res.ok) throw new GraphError(res.status, res.statusText);
      const buf = Buffer.from(await res.arrayBuffer());
      return { bytes: buf.length, truncated: buf.length > maxBytes, content: buf.subarray(0, maxBytes).toString("utf8") };
    } catch (err) {
      if (err instanceof GraphError) throw err;
      if (err instanceof Error && err.name === "AbortError") throw new GraphError(504, "Download timed out.");
      throw new GraphError(0, err instanceof Error ? err.message : String(err));
    } finally {
      clearTimeout(timer);
    }
  }

  /** Create or overwrite a small text file at a path (PUT simple upload, ≤4 MB). */
  uploadFile(path: string, content: string) {
    return this.request<Record<string, unknown>>(this.mailbox(`${this.driveRef({ path })}/content`), {
      method: "PUT",
      body: content,
      headers: { "Content-Type": "text/plain" },
    });
  }

  createFolder(parent: { id?: string; path?: string }, name: string) {
    return this.request<Record<string, unknown>>(this.mailbox(`${this.driveRef(parent)}/children`), {
      method: "POST",
      body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "rename" }),
    });
  }

  /** Move and/or rename an item. Pass a new parent folder id and/or a new name. */
  moveDriveItem(id: string, opts: { newParentId?: string; newName?: string }) {
    const body: Record<string, unknown> = {};
    if (opts.newParentId) body.parentReference = { id: opts.newParentId };
    if (opts.newName) body.name = opts.newName;
    return this.request<Record<string, unknown>>(this.mailbox(`/drive/items/${encodeURIComponent(id)}`), {
      method: "PATCH",
      body: JSON.stringify(body),
    });
  }

  /** Delete a drive item (moves it to the OneDrive recycle bin — recoverable). */
  deleteDriveItem(id: string) {
    return this.request<void>(this.mailbox(`/drive/items/${encodeURIComponent(id)}`), { method: "DELETE" });
  }
}
