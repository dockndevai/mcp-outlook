import type { DriveItemRef, EventContext, MessageContext } from "../graph/client.js";

/**
 * Extra lines for the human approval prompt describing the object an id points at. Display only: the
 * binding is the confirmer's argument fingerprint (+ version probe where there is an existing object).
 * The lookup is best-effort, so a failed lookup shows "unavailable" rather than blocking the prompt.
 */
type Details = Record<string, string | number | boolean | undefined>;

const join = (xs: string[] | undefined): string | undefined => (xs && xs.length ? xs.join(", ") : undefined);

export function messageDetails(prefix: string, m: MessageContext | undefined): Details {
  if (!m) return { [prefix]: "unavailable" };
  return {
    [`${prefix} subject`]: m.subject,
    [`${prefix} from`]: m.from,
    [`${prefix} to`]: join(m.to),
    [`${prefix} cc`]: join(m.cc),
  };
}

export function eventDetails(e: EventContext | undefined): Details {
  if (!e) return { event: "unavailable" };
  return {
    "event subject": e.subject,
    "event when": e.start ? `${e.start} – ${e.end ?? "?"}` : undefined,
    "event organizer": e.organizer,
    "event attendees": join(e.attendees),
  };
}

export function driveItemDetails(i: DriveItemRef | undefined): Details {
  if (!i) return { item: "unavailable" };
  return {
    "item name": i.name,
    "item type": i.isFolder ? `folder (${i.childCount ?? "?"} children)` : `file (${i.size ?? "?"} bytes)`,
    "item url": i.webUrl,
  };
}

/** What `upload_drive_file` is about to replace: the existing item at the path, or a new file. */
export function uploadDetails(existing: DriveItemRef | undefined): Details {
  if (!existing) return { "target state": "new file (nothing exists at this path)" };
  if (existing.isFolder) return { "target state": "a FOLDER already exists at this path (the upload will fail)" };
  return {
    "target state": `OVERWRITES an existing file (${existing.size ?? "?"} bytes, last modified ${existing.lastModified ?? "?"})`,
    "item url": existing.webUrl,
  };
}
