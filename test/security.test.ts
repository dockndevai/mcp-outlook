import { describe, expect, it } from "vitest";
import { PolicyError, SecurityPolicy, type SecurityConfig } from "../src/security.js";

function makePolicy(overrides: Partial<SecurityConfig> = {}): SecurityPolicy {
  return new SecurityPolicy({
    mode: "read-only",
    folderAllowlist: [],
    protectedFolders: ["archive"],
    allowSend: false,
    allowDelete: false,
    dryRun: false,
    auditLog: false,
    ...overrides,
  });
}

describe("capability gating", () => {
  it("read-only enables read only", () => {
    const p = makePolicy();
    expect(p.isCapabilityEnabled("read")).toBe(true);
    expect(p.isCapabilityEnabled("write")).toBe(false);
    expect(p.isCapabilityEnabled("admin")).toBe(false);
  });
  it("read-write enables read and write but not admin", () => {
    const p = makePolicy({ mode: "read-write" });
    expect(p.isCapabilityEnabled("write")).toBe(true);
    expect(p.isCapabilityEnabled("admin")).toBe(false);
  });
});

describe("mode vs capability at guard time", () => {
  it("rejects a write in read-only mode", () => {
    const p = makePolicy();
    expect(() => p.guard({ tool: "create_draft", capability: "write" })).toThrow(PolicyError);
  });
  it("rejects admin in read-write mode", () => {
    const p = makePolicy({ mode: "read-write" });
    expect(() => p.guard({ tool: "delete_message", capability: "admin", destructive: true })).toThrow(/admin/);
  });
});

describe("folder allowlist + protection", () => {
  it("blocks writes to folders outside a non-empty allowlist", () => {
    const p = makePolicy({ mode: "read-write", folderAllowlist: ["inbox"] });
    expect(() => p.guard({ tool: "move_message", capability: "write", folder: "junkemail" })).toThrow(/allowlist/);
  });
  it("allows reading a protected folder but not writing to it", () => {
    const p = makePolicy({ mode: "read-write" });
    expect(() => p.guard({ tool: "list_messages", capability: "read", folder: "archive" })).not.toThrow();
    expect(() => p.guard({ tool: "move_message", capability: "write", folder: "archive" })).toThrow(/protected/);
  });
});

describe("send gating", () => {
  it("blocks sending without allowSend even in read-write mode", () => {
    const p = makePolicy({ mode: "read-write" });
    expect(() => p.guard({ tool: "send_mail", capability: "write", requiresSend: true })).toThrow(/ALLOW_SEND/);
  });
  it("permits sending with allowSend", () => {
    const p = makePolicy({ mode: "read-write", allowSend: true });
    expect(() => p.guard({ tool: "send_mail", capability: "write", requiresSend: true })).not.toThrow();
  });
  it("non-sending writes are unaffected by allowSend", () => {
    const p = makePolicy({ mode: "read-write" });
    expect(() => p.guard({ tool: "create_draft", capability: "write" })).not.toThrow();
  });
});

describe("delete gating", () => {
  it("blocks delete without allowDelete even in admin mode", () => {
    const p = makePolicy({ mode: "admin" });
    expect(() => p.guard({ tool: "delete_message", capability: "admin", destructive: true })).toThrow(/ALLOW_DELETE/);
  });
  it("permits delete with allowDelete", () => {
    const p = makePolicy({ mode: "admin", allowDelete: true });
    expect(() => p.guard({ tool: "delete_message", capability: "admin", destructive: true })).not.toThrow();
  });
});

describe("dry run", () => {
  it("flags writes but not reads", () => {
    const p = makePolicy({ mode: "read-write", dryRun: true });
    expect(p.guard({ tool: "get_message", capability: "read" }).dryRun).toBe(false);
    expect(p.guard({ tool: "create_draft", capability: "write" }).dryRun).toBe(true);
  });
});
