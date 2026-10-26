import { existsSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const browser = vi.hoisted(() => ({
  launch: vi.fn(),
  close: vi.fn(),
  evaluate: vi.fn(),
  goto: vi.fn(),
  wait: vi.fn(),
}));

vi.mock("playwright", () => ({
  chromium: { launchPersistentContext: browser.launch },
}));

vi.mock("../src/config.js", () => {
  const dir = join(tmpdir(), "teams-cli-test-auth-refresh");
  return {
    DATA_DIR: dir,
    TOKENS_FILE: join(dir, "tokens.json"),
    BROWSER_PROFILE_DIR: join(dir, "browser-profile"),
    TOKEN_AUDIENCES: {
      ic3: "ic3.teams.office.com",
      search: "outlook.office.com/search",
      presence: "presence.teams.microsoft.com",
      csa: "chatsvcagg.teams.microsoft.com",
    },
    REGION_KEY_PATTERN: "DISCOVER-REGION-GTM",
    TEAMS_URL: "https://teams.cloud.microsoft/",
  };
});

import { getToken, loadTokens, login, saveTokens } from "../src/auth.js";

const testDir = join(tmpdir(), "teams-cli-test-auth-refresh");
let storage: Map<string, string>;

function addToken(client: string, secret: string, expiresOn: number | string) {
  storage.set(
    `msal.2|example|authority|accesstoken|${client}|tenant|https://ic3.teams.office.com/.default|`,
    JSON.stringify({ secret, expiresOn }),
  );
}

function expireSavedToken() {
  saveTokens({
    ic3: { secret: "expired-cache", expires_on: Date.now() / 1000 - 100 },
  });
}

beforeEach(() => {
  vi.useFakeTimers();
  vi.setSystemTime(new Date("2026-01-01T12:00:00Z"));
  storage = new Map();
  vi.stubGlobal("localStorage", {
    get length() {
      return storage.size;
    },
    key: (index: number) => [...storage.keys()][index] ?? null,
    getItem: (key: string) => storage.get(key) ?? null,
  });
  browser.close.mockReset().mockResolvedValue(undefined);
  browser.goto.mockReset().mockResolvedValue(undefined);
  browser.wait.mockReset().mockImplementation(async (milliseconds: number) => {
    vi.setSystemTime(Date.now() + milliseconds);
  });
  browser.evaluate
    .mockReset()
    .mockImplementation(async (fn: (arg: string) => unknown, arg: string) =>
      fn(arg),
    );
  browser.launch.mockReset().mockResolvedValue({
    newPage: async () => ({
      goto: browser.goto,
      evaluate: browser.evaluate,
      waitForTimeout: browser.wait,
    }),
    close: browser.close,
  });
  vi.spyOn(console, "log").mockImplementation(() => {});
  vi.spyOn(console, "error").mockImplementation(() => {});
  mkdirSync(testDir, { recursive: true });
});

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
  rmSync(testDir, { recursive: true, force: true });
});

describe("browser token validity", () => {
  it("preserves the default behavior for missing tokens", async () => {
    expireSavedToken();
    expect(await getToken("csa")).toBeNull();
    expect(browser.launch).not.toHaveBeenCalled();
  });

  it("can explicitly acquire a missing configured resource without changing chat identity", async () => {
    expireSavedToken();
    addToken("chat", "valid-chat", Date.now() / 1000 + 3600);
    storage.set(
      "msal.2|example|authority|accesstoken|csa|tenant|https://chatsvcagg.teams.microsoft.com/.default|",
      JSON.stringify({
        secret: "valid-csa",
        expiresOn: Date.now() / 1000 + 3600,
      }),
    );
    expect(await getToken("csa", true)).toBe("valid-csa");
    expect(await getToken("ic3")).toBe("valid-chat");
    expect(browser.launch).toHaveBeenCalledOnce();
    expect(browser.close).toHaveBeenCalledOnce();
  });

  it("does not attempt to acquire arbitrary missing resources", async () => {
    expireSavedToken();
    expect(await getToken("unknown", true)).toBeNull();
    expect(await getToken("toString", true)).toBeNull();
    expect(browser.launch).not.toHaveBeenCalled();
  });

  it("does not acquire a missing token without a saved sign-in", async () => {
    expect(await getToken("csa", true)).toBeNull();
    expect(browser.launch).not.toHaveBeenCalled();
  });

  it("reports a missing resource token without discarding saved tokens", async () => {
    expireSavedToken();
    addToken("chat", "valid-chat", Date.now() / 1000 + 3600);
    expect(await getToken("csa", true)).toBeNull();
    expect(loadTokens()?.ic3).toMatchObject({ secret: "expired-cache" });
    expect(console.error).toHaveBeenCalledWith(
      expect.stringContaining("requested resource token"),
    );
    expect(browser.close).toHaveBeenCalledOnce();
  });

  it("does not treat an expired browser token as a successful refresh", async () => {
    expireSavedToken();
    addToken("old", "expired-browser", Date.now() / 1000 - 100);
    expect(await getToken("ic3")).toBeNull();
    expect(loadTokens()?.ic3).toMatchObject({ secret: "expired-cache" });
    expect(browser.close).toHaveBeenCalledOnce();
    expect(console.error).toHaveBeenCalledWith(
      expect.stringContaining("unexpired chat token within 30 seconds"),
    );
  });

  it("skips an expired first match and selects the newest valid token", async () => {
    expireSavedToken();
    addToken("old", "expired-browser", Date.now() / 1000 - 100);
    addToken("valid", "older-valid", Date.now() / 1000 + 100);
    addToken("new", "newest-valid", Date.now() / 1000 + 3600);
    expect(await getToken("ic3")).toBe("newest-valid");
    expect(loadTokens()?.ic3).toMatchObject({ secret: "newest-valid" });
    expect(browser.close).toHaveBeenCalledOnce();
  });

  it("waits for a fresh token to appear instead of returning a stale one", async () => {
    expireSavedToken();
    addToken("old", "expired-browser", Date.now() / 1000 - 100);
    browser.wait.mockImplementation(async (milliseconds: number) => {
      vi.setSystemTime(Date.now() + milliseconds);
      addToken("new", "renewed", Date.now() / 1000 + 3600);
    });
    expect(await getToken("ic3")).toBe("renewed");
    expect(browser.wait).toHaveBeenCalledWith(1000);
  });

  it.each([
    "invalid",
    "NaN",
    0,
  ])("rejects malformed browser expiry %s", async (expiry) => {
    expireSavedToken();
    addToken("bad", "invalid-expiry", expiry);
    addToken("new", "valid", Date.now() / 1000 + 3600);
    expect(await getToken("ic3")).toBe("valid");
  });

  it("rejects a token at its exact expiry boundary", async () => {
    expireSavedToken();
    addToken("boundary", "expired-now", Date.now() / 1000);
    expect(await getToken("ic3")).toBeNull();
  });

  it("refreshes a saved token with a malformed expiry instead of returning it", async () => {
    saveTokens({ ic3: { secret: "bad-cache", expires_on: "invalid" } });
    addToken("new", "valid", Date.now() / 1000 + 3600);
    expect(await getToken("ic3")).toBe("valid");
  });

  it("ignores empty browser secrets", async () => {
    expireSavedToken();
    addToken("empty", "", Date.now() / 1000 + 7200);
    addToken("new", "valid", Date.now() / 1000 + 3600);
    expect(await getToken("ic3")).toBe("valid");
  });

  it("closes the persistent context after a navigation failure", async () => {
    expireSavedToken();
    browser.goto.mockRejectedValue(new Error("PRIVATE browser diagnostic"));
    expect(await getToken("ic3")).toBeNull();
    expect(browser.close).toHaveBeenCalledOnce();
    expect(console.error).toHaveBeenCalledWith(
      expect.stringContaining("Teams page navigation"),
    );
    expect(JSON.stringify(vi.mocked(console.error).mock.calls)).not.toContain(
      "PRIVATE",
    );
  });
});

describe("interactive login", () => {
  it("does not report success or delete saved sessions for expired-only cache", async () => {
    expireSavedToken();
    addToken("old", "expired-browser", Date.now() / 1000 - 100);
    const profile = join(testDir, "browser-profile");
    mkdirSync(profile, { recursive: true });
    writeFileSync(join(profile, "keep.txt"), "saved session");
    await expect(login()).rejects.toThrow("without an unexpired chat token");
    expect(loadTokens()?.ic3).toMatchObject({ secret: "expired-cache" });
    expect(existsSync(join(profile, "keep.txt"))).toBe(true);
    expect(browser.close).toHaveBeenCalledOnce();
  });

  it("captures only unexpired matching tokens", async () => {
    addToken("old", "expired-browser", Date.now() / 1000 - 100);
    addToken("new", "valid", Date.now() / 1000 + 3600);
    storage.set(
      "msal.2|example|authority|accesstoken|search|tenant|https://outlook.office.com/search/.default|",
      JSON.stringify({
        secret: "expired-search",
        expiresOn: Date.now() / 1000 - 100,
      }),
    );
    const tokens = await login();
    expect(tokens.ic3).toMatchObject({ secret: "valid" });
    expect(tokens.search).toBeUndefined();
    expect(browser.close).toHaveBeenCalledOnce();
  });
});
