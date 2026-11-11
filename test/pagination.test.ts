import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { GraphClient } from "../src/graph/client.js";
import type { OutlookConnection } from "../src/config.js";

function makeConnection(): OutlookConnection {
  return {
    graphBaseUrl: "https://graph.microsoft.com/v1.0",
    authorityBase: "https://login.microsoftonline.com",
    auth: {
      mode: "token",
      token: "test-token",
      tenantId: "common",
      scopes: [],
      tokenCachePath: "/tmp/unused",
      redirectPort: 0,
    },
    userPath: "/me",
    timeoutMs: 5000,
    maxResults: 50,
  };
}

describe("pagination (@odata.nextLink)", () => {
  const originalFetch = global.fetch;

  beforeEach(() => {
    global.fetch = vi.fn();
  });

  afterEach(() => {
    global.fetch = originalFetch;
  });

  it("round-trips a cursor to fetch the second page of messages", async () => {
    const nextLink = "https://graph.microsoft.com/v1.0/me/messages?$top=50&$skip=50";
    const fetchMock = global.fetch as unknown as ReturnType<typeof vi.fn>;
    fetchMock
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ value: [{ id: "msg-1" }], "@odata.nextLink": nextLink }), { status: 200 }),
      )
      .mockResolvedValueOnce(new Response(JSON.stringify({ value: [{ id: "msg-2" }] }), { status: 200 }));

    const client = new GraphClient(makeConnection());

    const page1 = await client.listMessages({});
    expect(page1.items.map((m) => m.id)).toEqual(["msg-1"]);
    expect(page1.nextLink).toBe(nextLink);

    const page2 = await client.listMessages({ cursor: page1.nextLink });
    expect(page2.items.map((m) => m.id)).toEqual(["msg-2"]);
    expect(page2.nextLink).toBeUndefined();

    // The cursor is requested as-is (the full nextLink URL), not re-derived from other opts.
    expect(fetchMock.mock.calls[1]?.[0]).toBe(nextLink);
  });
});
