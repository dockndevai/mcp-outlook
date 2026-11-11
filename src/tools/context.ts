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
