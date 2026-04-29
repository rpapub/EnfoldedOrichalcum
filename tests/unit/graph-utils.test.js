import { describe, it, expect, vi, beforeEach } from "vitest";

// Stub browser globals before requiring the module
const mockStorage = (() => {
  let store = {};
  return {
    getItem: (k) => store[k] ?? null,
    setItem: (k, v) => { store[k] = String(v); },
    removeItem: (k) => { delete store[k]; },
    _reset: () => { store = {}; },
  };
})();
vi.stubGlobal("localStorage", mockStorage);
vi.stubGlobal("fetch", vi.fn());

const {
  getCachedToken, cacheToken, clearCachedToken,
  graphError, readGraphExtension, deleteGraphExtension,
} = require("../../docs/shared/graph-utils.js");

beforeEach(() => { mockStorage._reset(); vi.clearAllMocks(); });

describe("getCachedToken", () => {
  it("returns null when cache empty", () => expect(getCachedToken()).toBeNull());

  it("returns null when token expired", () => {
    cacheToken("tok", -100); // already expired
    expect(getCachedToken()).toBeNull();
  });

  it("returns token when valid", () => {
    cacheToken("tok123", 3600);
    expect(getCachedToken()).toBe("tok123");
  });
});

describe("clearCachedToken", () => {
  it("removes a cached token", () => {
    cacheToken("tok", 3600);
    clearCachedToken();
    expect(getCachedToken()).toBeNull();
  });
});

describe("graphError", () => {
  it("attaches .status to the error", () => {
    const e = graphError(401, "Unauthorized");
    expect(e.status).toBe(401);
    expect(e.message).toBe("Unauthorized");
  });
});

describe("readGraphExtension", () => {
  it("returns null on 404", async () => {
    fetch.mockResolvedValue({ status: 404, ok: false });
    const result = await readGraphExtension("tok", "msgId", "ext.name");
    expect(result).toBeNull();
  });

  it("returns parsed JSON on 200", async () => {
    fetch.mockResolvedValue({
      status: 200, ok: true,
      json: async () => ({ id: "ext.name", value: 42 }),
    });
    const result = await readGraphExtension("tok", "msgId", "ext.name");
    expect(result).toEqual({ id: "ext.name", value: 42 });
  });

  it("throws with .status on error", async () => {
    fetch.mockResolvedValue({ status: 500, ok: false });
    await expect(readGraphExtension("tok", "msgId", "ext.name")).rejects.toMatchObject({ status: 500 });
  });
});

describe("deleteGraphExtension", () => {
  it("resolves on 204", async () => {
    fetch.mockResolvedValue({ status: 204, ok: true });
    await expect(deleteGraphExtension("tok", "msgId", "ext.name")).resolves.toBeUndefined();
  });

  it("throws with .status on error", async () => {
    fetch.mockResolvedValue({
      status: 403, ok: false,
      text: async () => "Forbidden",
    });
    await expect(deleteGraphExtension("tok", "msgId", "ext.name")).rejects.toMatchObject({ status: 403 });
  });
});
