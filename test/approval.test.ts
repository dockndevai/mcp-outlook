import { afterEach, describe, expect, it, vi } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { ElicitRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import type { AppConfig } from "../src/config.js";
import { GraphClient, versionToken } from "../src/graph/client.js";
import { buildServer } from "../src/server.js";

const config = (requireHumanApproval = false): AppConfig => ({
  connection: {
    graphBaseUrl: "https://graph.test/v1.0",
    authorityBase: "https://login.test",
    auth: { mode: "token", token: "t", tenantId: "common", scopes: [], tokenCachePath: "/dev/null", redirectPort: 0 },
    userPath: "/me",
    timeoutMs: 1000,
    maxResults: 25,
  } as AppConfig["connection"],
  security: {
    mode: "admin",
    folderAllowlist: [],
    protectedFolders: [],
    allowSend: true,
    allowDelete: true,
    dryRun: false,
    auditLog: false,
    requireHumanApproval,
  },
});

type Answer = { action: "accept" | "decline"; content?: Record<string, unknown> };

/** Connect an in-memory MCP client (optionally able to answer elicitation) to a freshly built server. */
async function connect(opts: { requireHuman?: boolean; answer?: (message: string) => Answer | Promise<Answer>; elicits?: boolean }) {
  const { server } = buildServer(config(opts.requireHuman));
  const prompts: string[] = [];
  const client = new Client({ name: "t", version: "0" }, { capabilities: opts.elicits === false ? {} : { elicitation: {} } });
  if (opts.elicits !== false) {
    client.setRequestHandler(ElicitRequestSchema, async (req) => {
      prompts.push(req.params.message);
      return (await opts.answer?.(req.params.message)) ?? { action: "accept", content: { confirm: true } };
    });
  }
  const [a, b] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(a), client.connect(b)]);
  return { client, prompts };
}

const text = (r: unknown) => ((r as { content: Array<{ text: string }> }).content[0]?.text ?? "");

/** Successive values for a version probe: first read (before the prompt), then second (after the human answered). */
const sequence = (...vs: Array<string | undefined>) => {
  const q = [...vs];
  return async () => q.shift();
};

afterEach(() => vi.restoreAllMocks());

describe("delete_message approval (end to end)", () => {
  it("shows the tool, exact arguments, version and fingerprint, then deletes", async () => {
    vi.spyOn(GraphClient.prototype, "messageContext").mockResolvedValue({ subject: "Q3 numbers", from: "boss@contoso.com", to: ["me@contoso.com"] });
    vi.spyOn(GraphClient.prototype, "messageVersion").mockResolvedValue("v1");
    const del = vi.spyOn(GraphClient.prototype, "deleteMessage").mockResolvedValue();
    const { client, prompts } = await connect({});
    const res = await client.callTool({ name: "delete_message", arguments: { id: "AAMk-1" } });
    expect(text(res)).toContain('"deleted": "AAMk-1"');
    expect(del).toHaveBeenCalledWith("AAMk-1");
    expect(prompts).toHaveLength(1);
    expect(prompts[0]).toContain("tool: delete_message");
    expect(prompts[0]).toContain('{"id":"AAMk-1"}');
    expect(prompts[0]).toContain("message subject: Q3 numbers");
    expect(prompts[0]).toContain("resource version: v1");
    expect(prompts[0]).toMatch(/fingerprint: [0-9a-f]{16}/);
  });

  it("does NOT delete when the message changed while the human was deciding", async () => {
    vi.spyOn(GraphClient.prototype, "messageContext").mockResolvedValue(undefined);
    vi.spyOn(GraphClient.prototype, "messageVersion").mockImplementation(sequence("v1", "v2"));
    const del = vi.spyOn(GraphClient.prototype, "deleteMessage").mockResolvedValue();
    const { client } = await connect({});
    const res = await client.callTool({ name: "delete_message", arguments: { id: "AAMk-1" } });
    expect(text(res)).toContain("resource changed since the prompt");
    expect(del).not.toHaveBeenCalled();
  });

  it("does NOT delete when the message vanished while the human was deciding", async () => {
    vi.spyOn(GraphClient.prototype, "messageContext").mockResolvedValue(undefined);
    vi.spyOn(GraphClient.prototype, "messageVersion").mockImplementation(sequence("v1", undefined));
    const del = vi.spyOn(GraphClient.prototype, "deleteMessage").mockResolvedValue();
    const { client } = await connect({});
    const res = await client.callTool({ name: "delete_message", arguments: { id: "AAMk-1" } });
    expect(text(res)).toContain("resource changed since the prompt");
    expect(del).not.toHaveBeenCalled();
  });

  it("does NOT delete when the human declines", async () => {
    vi.spyOn(GraphClient.prototype, "messageContext").mockResolvedValue(undefined);
    vi.spyOn(GraphClient.prototype, "messageVersion").mockResolvedValue("v1");
    const del = vi.spyOn(GraphClient.prototype, "deleteMessage").mockResolvedValue();
    const { client } = await connect({ answer: () => ({ action: "decline" }) });
    await client.callTool({ name: "delete_message", arguments: { id: "AAMk-1" } });
    expect(del).not.toHaveBeenCalled();
  });
});

describe("send_mail approval (end to end)", () => {
  const args = {
    subject: "Wire details",
    body: "Please pay invoice 42.",
    to: [{ address: "alice@contoso.com" }],
    cc: [{ address: "carol@contoso.com" }],
    bcc: [{ address: "eve@evil.test" }],
  };

  it("shows recipients (to/cc/bcc), subject, body, fingerprint — then sends exactly those arguments", async () => {
    const send = vi.spyOn(GraphClient.prototype, "sendMail").mockResolvedValue();
    const { client, prompts } = await connect({});
    const res = await client.callTool({ name: "send_mail", arguments: args });
    expect(text(res)).toContain("Sent");
    expect(send).toHaveBeenCalledTimes(1);
    const p = prompts[0]!;
    expect(p).toContain("tool: send_mail");
    expect(p).toContain("to: alice@contoso.com");
    expect(p).toContain("cc: carol@contoso.com");
    expect(p).toContain("bcc: eve@evil.test");
    expect(p).toContain("subject: Wire details");
    expect(p).toContain("Please pay invoice 42.");
    expect(p).toMatch(/fingerprint: [0-9a-f]{16}/);
  });

  it("keeps the recipients visible even when the body is far longer than the display limit", async () => {
    vi.spyOn(GraphClient.prototype, "sendMail").mockResolvedValue();
    const { client, prompts } = await connect({});
    await client.callTool({ name: "send_mail", arguments: { ...args, body: "x".repeat(20000) } });
    const p = prompts[0]!;
    expect(p).toContain("not shown; covered by the fingerprint"); // an over-long prose body is summarised (the tool opts in)…
    expect(p).toContain("to: alice@contoso.com"); // …but the recipients are shown outside it
    expect(p).toContain("bcc: eve@evil.test");
    expect(p).toContain("body length: 20000 chars");
  });

  it("different arguments give a different fingerprint", async () => {
    vi.spyOn(GraphClient.prototype, "sendMail").mockResolvedValue();
    const { client, prompts } = await connect({});
    await client.callTool({ name: "send_mail", arguments: args });
    await client.callTool({ name: "send_mail", arguments: { ...args, bcc: [{ address: "someone@else.test" }] } });
    const fp = (s: string) => /fingerprint: ([0-9a-f]{16})/.exec(s)?.[1];
    expect(fp(prompts[0]!)).toBeDefined();
    expect(fp(prompts[0]!)).not.toBe(fp(prompts[1]!));
  });

  it("does NOT send when the human declines", async () => {
    const send = vi.spyOn(GraphClient.prototype, "sendMail").mockResolvedValue();
    const { client } = await connect({ answer: () => ({ action: "decline" }) });
    const res = await client.callTool({ name: "send_mail", arguments: args });
    expect(text(res)).toContain("Send cancelled");
    expect(send).not.toHaveBeenCalled();
  });
});

describe("OUTLOOK_REQUIRE_HUMAN_APPROVAL", () => {
  it("falls back to the flags when the client cannot elicit (default)", async () => {
    vi.spyOn(GraphClient.prototype, "messageContext").mockResolvedValue(undefined);
    vi.spyOn(GraphClient.prototype, "messageVersion").mockResolvedValue("v1");
    const del = vi.spyOn(GraphClient.prototype, "deleteMessage").mockResolvedValue();
    const { client } = await connect({ elicits: false });
    await client.callTool({ name: "delete_message", arguments: { id: "AAMk-1" } });
    expect(del).toHaveBeenCalledWith("AAMk-1");
  });

  it("refuses to delete when the client cannot elicit and human approval is required", async () => {
    vi.spyOn(GraphClient.prototype, "messageContext").mockResolvedValue(undefined);
    vi.spyOn(GraphClient.prototype, "messageVersion").mockResolvedValue("v1");
    const del = vi.spyOn(GraphClient.prototype, "deleteMessage").mockResolvedValue();
    const { client } = await connect({ elicits: false, requireHuman: true });
    const res = await client.callTool({ name: "delete_message", arguments: { id: "AAMk-1" } });
    expect(text(res)).toContain("no human approval channel");
    expect(del).not.toHaveBeenCalled();
  });

  it("refuses to send when the client cannot elicit and human approval is required", async () => {
    const send = vi.spyOn(GraphClient.prototype, "sendMail").mockResolvedValue();
    const { client } = await connect({ elicits: false, requireHuman: true });
    const res = await client.callTool({
      name: "send_mail",
      arguments: { subject: "s", body: "b", to: [{ address: "a@b.test" }] },
    });
    expect(text(res)).toContain("no human approval channel");
    expect(send).not.toHaveBeenCalled();
  });
});

describe("other destructive / outward tools bind too", () => {
  it("cancel_event refuses when the event changed while the human was deciding", async () => {
    vi.spyOn(GraphClient.prototype, "eventContext").mockResolvedValue({ subject: "Board", attendees: ["a@b.test"] });
    vi.spyOn(GraphClient.prototype, "eventVersion").mockImplementation(sequence("e1", "e2"));
    const cancel = vi.spyOn(GraphClient.prototype, "cancelEvent").mockResolvedValue();
    const { client } = await connect({});
    const res = await client.callTool({ name: "cancel_event", arguments: { id: "EV1", comment: "sorry" } });
    expect(text(res)).toContain("resource changed since the prompt");
    expect(cancel).not.toHaveBeenCalled();
  });

  it("cancel_event shows the attendees and the cancellation note, then cancels on a stable version", async () => {
    vi.spyOn(GraphClient.prototype, "eventContext").mockResolvedValue({ subject: "Board", attendees: ["a@b.test"] });
    vi.spyOn(GraphClient.prototype, "eventVersion").mockResolvedValue("e1");
    const cancel = vi.spyOn(GraphClient.prototype, "cancelEvent").mockResolvedValue();
    const { client, prompts } = await connect({});
    await client.callTool({ name: "cancel_event", arguments: { id: "EV1", comment: "sorry" } });
    expect(cancel).toHaveBeenCalledWith("EV1", "sorry");
    expect(prompts[0]).toContain("event attendees: a@b.test");
    expect(prompts[0]).toContain("cancellation note: sorry");
  });

  it("delete_event and update_event refuse on a changed version", async () => {
    vi.spyOn(GraphClient.prototype, "eventContext").mockResolvedValue(undefined);
    vi.spyOn(GraphClient.prototype, "eventVersion").mockImplementation(sequence("e1", "e2", "e1", "e2"));
    const del = vi.spyOn(GraphClient.prototype, "deleteEvent").mockResolvedValue();
    const upd = vi.spyOn(GraphClient.prototype, "updateEvent").mockResolvedValue({});
    const { client } = await connect({});
    expect(text(await client.callTool({ name: "delete_event", arguments: { id: "EV1" } }))).toContain("resource changed");
    expect(text(await client.callTool({ name: "update_event", arguments: { id: "EV1", subject: "new" } }))).toContain("resource changed");
    expect(del).not.toHaveBeenCalled();
    expect(upd).not.toHaveBeenCalled();
  });

  it("delete_drive_item refuses when the item changed while the human was deciding", async () => {
    vi.spyOn(GraphClient.prototype, "driveItemContext").mockResolvedValue({ id: "D1", name: "plan.docx", isFolder: false, size: 10 });
    vi.spyOn(GraphClient.prototype, "driveItemVersion").mockImplementation(sequence("d1", "d2"));
    const del = vi.spyOn(GraphClient.prototype, "deleteDriveItem").mockResolvedValue();
    const { client } = await connect({});
    const res = await client.callTool({ name: "delete_drive_item", arguments: { item_id: "D1" } });
    expect(text(res)).toContain("resource changed since the prompt");
    expect(del).not.toHaveBeenCalled();
  });

  it("reply_mail and forward_mail show who the message came from / goes to, and bind by fingerprint", async () => {
    vi.spyOn(GraphClient.prototype, "messageContext").mockResolvedValue({ subject: "Offer", from: "hr@contoso.com", to: ["me@contoso.com"] });
    const reply = vi.spyOn(GraphClient.prototype, "replyMessage").mockResolvedValue();
    const fwd = vi.spyOn(GraphClient.prototype, "forwardMessage").mockResolvedValue();
    const { client, prompts } = await connect({});
    await client.callTool({ name: "reply_mail", arguments: { id: "M1", comment: "yes" } });
    await client.callTool({ name: "forward_mail", arguments: { id: "M1", to: [{ address: "x@y.test" }], comment: "fyi" } });
    expect(reply).toHaveBeenCalledWith("M1", "yes", false);
    expect(fwd).toHaveBeenCalledTimes(1);
    expect(prompts[0]).toContain("original subject: Offer");
    expect(prompts[0]).toContain("hr@contoso.com");
    expect(prompts[1]).toContain("x@y.test");
    expect(prompts[1]).toMatch(/fingerprint: [0-9a-f]{16}/);
  });

  it("create_event and respond_to_event do not run when declined", async () => {
    vi.spyOn(GraphClient.prototype, "eventContext").mockResolvedValue(undefined);
    const create = vi.spyOn(GraphClient.prototype, "createEvent").mockResolvedValue({});
    const respond = vi.spyOn(GraphClient.prototype, "respondToEvent").mockResolvedValue();
    const { client } = await connect({ answer: () => ({ action: "decline" }) });
    await client.callTool({ name: "create_event", arguments: { subject: "s", start: "2026-01-01T10:00:00", end: "2026-01-01T11:00:00", attendees: ["a@b.test"] } });
    await client.callTool({ name: "respond_to_event", arguments: { id: "EV1", response: "accept" } });
    expect(create).not.toHaveBeenCalled();
    expect(respond).not.toHaveBeenCalled();
  });
});

describe("Graph version probes", () => {
  const client = () => new GraphClient(config().connection);
  const reply = (status: number, body: unknown) => ({
    ok: status >= 200 && status < 300,
    status,
    statusText: String(status),
    text: async () => JSON.stringify(body),
  });

  it("messageVersion keys on id + changeKey, so a changed changeKey is a different version", async () => {
    let key = "CQ-1";
    vi.stubGlobal("fetch", async () => reply(200, { id: "M1", changeKey: key }));
    const a = await client().messageVersion("M1");
    key = "CQ-2";
    const b = await client().messageVersion("M1");
    expect(a).toBe(versionToken("M1", "CQ-1"));
    expect(a).not.toBe(b);
    vi.unstubAllGlobals();
  });

  it("returns undefined for a 404 and throws on transport errors", async () => {
    vi.stubGlobal("fetch", async () => reply(404, { error: { message: "not found" } }));
    expect(await client().messageVersion("gone")).toBeUndefined();
    expect(await client().eventVersion("gone")).toBeUndefined();
    expect(await client().driveItemVersion("gone")).toBeUndefined();
    vi.stubGlobal("fetch", async () => reply(500, { error: { message: "boom" } }));
    await expect(client().messageVersion("M1")).rejects.toThrow("boom");
    vi.unstubAllGlobals();
  });

  it("falls back to @odata.etag when changeKey is absent; drive items use eTag", async () => {
    vi.stubGlobal("fetch", async () => reply(200, { id: "E1", "@odata.etag": 'W/"abc"' }));
    expect(await client().eventVersion("E1")).toBe(versionToken("E1", 'W/"abc"'));
    vi.stubGlobal("fetch", async () => reply(200, { id: "D1", eTag: '"{G},3"', cTag: "ignored" }));
    expect(await client().driveItemVersion("D1")).toBe(versionToken("D1", '"{G},3"'));
    vi.unstubAllGlobals();
  });
});
