import { existsSync, mkdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// Prevent tryRefresh from launching a real browser
vi.mock("playwright", () => {
  throw new Error("playwright not available in tests");
});

vi.mock("../src/config.js", () => {
  const dir = join(tmpdir(), "teams-cli-test-auth");
  return {
    DATA_DIR: dir,
    TOKENS_FILE: join(dir, "tokens.json"),
    BROWSER_PROFILE_DIR: join(dir, "browser-profile"),
    TOKEN_AUDIENCES: {
      ic3: "ic3.teams.office.com",
      search: "outlook.office.com/search",
      presence: "presence.teams.microsoft.com",
    },
    REGION_KEY_PATTERN: "DISCOVER-REGION-GTM",
    TEAMS_URL: "https://teams.cloud.microsoft/",
  };
});

import {
  getMyMri,
  getRegion,
  getToken,
  loadTokens,
  logout,
  saveTokens,
} from "../src/auth.js";

const testDir = join(tmpdir(), "teams-cli-test-auth");

beforeEach(() => {
  mkdirSync(testDir, { recursive: true });
});

afterEach(() => {
  if (existsSync(testDir)) {
    rmSync(testDir, { recursive: true });
  }
});

describe("saveTokens / loadTokens", () => {
  it("round trips", () => {
    const data = {
      ic3: {
        secret: "tok123",
        expires_on: String(Math.floor(Date.now() / 1000) + 3600),
      },
    };
    saveTokens(data);
    const loaded = loadTokens();
    expect((loaded?.ic3 as { secret: string }).secret).toBe("tok123");
  });
  it("returns null when missing", () => {
    // Remove the test dir so tokens file doesn't exist
    if (existsSync(testDir)) rmSync(testDir, { recursive: true });
    expect(loadTokens()).toBeNull();
  });
});

describe("getToken", () => {
  it("returns secret for valid token", async () => {
    saveTokens({
      ic3: {
        secret: "mytoken",
        expires_on: String(Math.floor(Date.now() / 1000) + 3600),
      },
    });
    expect(await getToken("ic3")).toBe("mytoken");
  });
  it("returns null for expired token", async () => {
    saveTokens({
      ic3: {
        secret: "old",
        expires_on: String(Math.floor(Date.now() / 1000) - 100),
      },
    });
    expect(await getToken("ic3")).toBeNull();
  });
  it("returns null for missing token", async () => {
    saveTokens({});
    expect(await getToken("ic3")).toBeNull();
  });
});

describe("getRegion", () => {
  it("returns region", () => {
    saveTokens({ region: "au" });
    expect(getRegion()).toBe("au");
  });
  it("defaults to amer", () => {
    if (existsSync(testDir)) rmSync(testDir, { recursive: true });
    expect(getRegion()).toBe("amer");
  });
});

describe("getMyMri", () => {
  it("extracts oid from JWT", () => {
    const header = Buffer.from(JSON.stringify({ alg: "RS256" })).toString(
      "base64url",
    );
    const payload = Buffer.from(
      JSON.stringify({ oid: "test-uuid-123" }),
    ).toString("base64url");
    const jwt = `${header}.${payload}.signature`;
    saveTokens({
      ic3: {
        secret: jwt,
        expires_on: String(Math.floor(Date.now() / 1000) + 3600),
      },
    });
    expect(getMyMri()).toBe("8:orgid:test-uuid-123");
  });
  it("returns null when no tokens", () => {
    if (existsSync(testDir)) rmSync(testDir, { recursive: true });
    expect(getMyMri()).toBeNull();
  });

  it("returns null for malformed JWT (not 3 parts)", () => {
    saveTokens({
      ic3: {
        secret: "not-a-jwt",
        expires_on: String(Math.floor(Date.now() / 1000) + 3600),
      },
    });
    expect(getMyMri()).toBeNull();
  });

  it("returns null when JWT has no oid claim", () => {
    const header = Buffer.from(JSON.stringify({ alg: "RS256" })).toString(
      "base64url",
    );
    const payload = Buffer.from(
      JSON.stringify({ sub: "no-oid-here" }),
    ).toString("base64url");
    const jwt = `${header}.${payload}.signature`;
    saveTokens({
      ic3: {
        secret: jwt,
        expires_on: String(Math.floor(Date.now() / 1000) + 3600),
      },
    });
    expect(getMyMri()).toBeNull();
  });

  it("returns null when ic3 entry has no secret", () => {
    saveTokens({
      ic3: {
        secret: "",
        expires_on: String(Math.floor(Date.now() / 1000) + 3600),
      },
    });
    expect(getMyMri()).toBeNull();
  });

  it("returns null when ic3 is a string (not TokenEntry)", () => {
    saveTokens({
      ic3: "just-a-string",
    } as unknown as import("../src/auth.js").TokenStore);
    expect(getMyMri()).toBeNull();
  });
});

describe("logout", () => {
  it("removes tokens file", () => {
    saveTokens({ ic3: { secret: "x", expires_on: "9999999999" } });
    expect(loadTokens()).not.toBeNull();
    logout();
    expect(loadTokens()).toBeNull();
  });

  it("does not throw when nothing to remove", () => {
    if (existsSync(testDir)) rmSync(testDir, { recursive: true });
    expect(() => logout()).not.toThrow();
  });
});
