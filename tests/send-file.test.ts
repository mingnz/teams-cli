import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const hoisted = vi.hoisted(() => ({ playwrightImports: 0 }));

// Any fall-through to the browser ladder must not launch a real browser.
vi.mock("playwright", () => {
  hoisted.playwrightImports++;
  throw new Error("playwright not available in tests");
});

vi.mock("../src/config.js", async (importOriginal) => {
  const orig = await importOriginal<typeof import("../src/config.js")>();
  const { join } = await import("node:path");
  const { tmpdir } = await import("node:os");
  const dir = join(tmpdir(), "teams-cli-test-send-file");
  return {
    ...orig,
    DATA_DIR: dir,
    TOKENS_FILE: join(dir, "tokens.json"),
    LAST_CHATS_FILE: join(dir, "last_chats.json"),
    BROWSER_PROFILE_DIR: join(dir, "browser-profile"),
  };
});

const CHAT_BASE = "https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME";
const chatFetch = vi.fn();
const getChatClient = vi.fn(async () => ({
  baseUrl: CHAT_BASE,
  fetch: chatFetch,
}));

vi.mock("../src/client.js", () => ({
  getChatClient: () => getChatClient(),
  getSearchClient: vi.fn(),
  getStreamClient: vi.fn(),
}));

const { createProgram } = await import("../src/cli.js");
const { getOneDriveToken, peekOneDriveToken, saveTokens } = await import(
  "../src/auth.js"
);

const DATA_DIR = join(tmpdir(), "teams-cli-test-send-file");
const CONV = "19:abc@thread.v2";
const FUTURE = Math.floor(Date.now() / 1000) + 3600;

function fakeJwt(claims: Record<string, unknown>): string {
  const enc = (o: unknown) =>
    Buffer.from(JSON.stringify(o)).toString("base64url");
  return `${enc({ alg: "none" })}.${enc(claims)}.sig`;
}

const ENV_TOKEN = fakeJwt({
  aud: "https://contoso-my.sharepoint.com",
  upn: "a.b@contoso.com",
  exp: FUTURE,
});
const CACHED_TOKEN = fakeJwt({
  aud: "https://fabrikam-my.sharepoint.com",
  upn: "c.d@fabrikam.com",
  exp: FUTURE,
});
const CONTOSO_SITE =
  "https://contoso-my.sharepoint.com/personal/a_b_contoso_com";

function jsonResp(status: number, body: unknown) {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
    text: async () => JSON.stringify(body),
  };
}

function stubOneDrive(failAt?: "link") {
  const spFetch = vi.fn(async (url: string, init?: RequestInit) => {
    if (init?.method === "PUT") {
      return jsonResp(201, { id: "ITEM1", name: "notes.txt" });
    }
    if (url.includes("createLink")) {
      return failAt === "link"
        ? jsonResp(403, { error: "accessDenied" })
        : jsonResp(201, {
            id: "share-1",
            link: { webUrl: "https://contoso-my.sharepoint.com/:t:/g/x" },
          });
    }
    return jsonResp(200, {
      sharepointIds: {
        listItemUniqueId: "11111111-2222-3333-4444-555555555555",
        siteId: "site-1",
        listId: "list-1",
      },
    });
  });
  vi.stubGlobal("fetch", spFetch);
  return spFetch;
}

describe("send-file command", () => {
  let out: string[];
  let err: string[];
  let file: string;

  beforeEach(() => {
    out = [];
    err = [];
    rmSync(DATA_DIR, { recursive: true, force: true });
    mkdirSync(DATA_DIR, { recursive: true });
    file = join(mkdtempSync(join(tmpdir(), "teams-send-file-")), "notes.txt");
    writeFileSync(file, "abc");
    chatFetch.mockReset();
    getChatClient.mockClear();
    delete process.env.TEAMS_SP_TOKEN;
    vi.spyOn(console, "log").mockImplementation((s) => out.push(String(s)));
    vi.spyOn(console, "error").mockImplementation((s) => err.push(String(s)));
    vi.spyOn(process, "exit").mockImplementation(((code?: number) => {
      throw new Error(`exit ${code}`);
    }) as never);
  });

  afterEach(() => {
    delete process.env.TEAMS_SP_TOKEN;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  async function run(...args: string[]) {
    const program = createProgram();
    program.exitOverride();
    await program.parseAsync(["node", "teams", "send-file", ...args]);
  }

  it("--dry-run touches no network and prints a plan", async () => {
    process.env.TEAMS_SP_TOKEN = ENV_TOKEN;
    const netFetch = vi.fn(() => {
      throw new Error("network must not be touched");
    });
    vi.stubGlobal("fetch", netFetch);

    await run(CONV, file, "--caption", "a & b", "--dry-run");

    expect(netFetch).not.toHaveBeenCalled();
    expect(getChatClient).not.toHaveBeenCalled();
    expect(hoisted.playwrightImports).toBe(0);
    const plan = JSON.parse(out.join("\n"));
    expect(plan.conversationId).toBe(CONV);
    expect(plan.site.base).toBe(`${CONTOSO_SITE}/`);
    expect(plan.file).toMatchObject({ name: "notes.txt", size: 3 });
    expect(plan.linkScope).toBe("organization");
    expect(plan.permissionScope).toBe("Organization");
    expect(plan.body.content).toBe("<div><p>a &amp; b</p></div>");
    const [fo] = JSON.parse(plan.body.properties.files);
    expect(fo.id).toBe("<post-upload>");
    expect(JSON.stringify(plan)).not.toContain(ENV_TOKEN);
  });

  it("--dry-run without any offline token still makes no network call", async () => {
    const netFetch = vi.fn();
    vi.stubGlobal("fetch", netFetch);
    await run(CONV, file, "--link-scope", "anon", "--dry-run");
    expect(netFetch).not.toHaveBeenCalled();
    expect(hoisted.playwrightImports).toBe(0);
    const plan = JSON.parse(out.join("\n"));
    expect(plan.site).toMatch(/not derivable offline/);
    expect(plan.permissionScope).toBe("Anyone");
  });

  it("sends with TEAMS_SP_TOKEN taking precedence over a cached token", async () => {
    saveTokens({
      sharepoint: {
        "fabrikam-my.sharepoint.com": {
          secret: CACHED_TOKEN,
          expires_on: FUTURE,
        },
      },
    });
    process.env.TEAMS_SP_TOKEN = ENV_TOKEN;
    const spFetch = stubOneDrive();
    chatFetch.mockResolvedValue(
      jsonResp(201, { OriginalArrivalTime: 1700000000000 }),
    );

    await run(CONV, file);

    for (const [url, init] of spFetch.mock.calls) {
      expect(url.startsWith(`${CONTOSO_SITE}/_api/`)).toBe(true);
      expect(init.headers.Authorization).toBe(`Bearer ${ENV_TOKEN}`);
    }
    expect(chatFetch).toHaveBeenCalledTimes(1);
    expect(out.join("\n")).toContain("Arrival: 1700000000000");
    expect(out.join("\n")).toContain(
      "Share URL: https://contoso-my.sharepoint.com/:t:/g/x",
    );
    const log = err.join("\n");
    for (const s of [
      "[1/4 upload]",
      "[2/4 ids]",
      "[3/4 link]",
      "[4/4 message]",
    ])
      expect(log).toContain(s);
    expect(log).toContain("token: env");
    expect(log).not.toContain(ENV_TOKEN);
  });

  it("reports succeeded stages and exits nonzero when a stage fails", async () => {
    process.env.TEAMS_SP_TOKEN = ENV_TOKEN;
    stubOneDrive("link");
    await expect(run(CONV, file)).rejects.toThrow("exit 1");
    const log = err.join("\n");
    expect(log).toContain("Stage 'link' FAILED: HTTP 403");
    expect(log).toContain("Succeeded: upload, ids");
    expect(log).toContain("upload: notes.txt (item ITEM1)");
    expect(log).toMatch(/created but not posted/);
    expect(chatFetch).not.toHaveBeenCalled();
  });

  it("rejects an unknown --link-scope before any network", async () => {
    const netFetch = vi.fn();
    vi.stubGlobal("fetch", netFetch);
    await expect(run(CONV, file, "--link-scope", "users")).rejects.toThrow(
      "exit 1",
    );
    expect(err.join("\n")).toMatch(/Invalid --link-scope/);
    expect(netFetch).not.toHaveBeenCalled();
  });

  it("fails before uploading when no OneDrive token is available", async () => {
    const netFetch = vi.fn();
    vi.stubGlobal("fetch", netFetch);
    await expect(run(CONV, file)).rejects.toThrow("exit 1");
    expect(err.join("\n")).toMatch(/Nothing was uploaded/);
    expect(netFetch).not.toHaveBeenCalled();
  });
});

describe("OneDrive token ladder", () => {
  beforeEach(() => {
    rmSync(DATA_DIR, { recursive: true, force: true });
    mkdirSync(DATA_DIR, { recursive: true });
    delete process.env.TEAMS_SP_TOKEN;
  });
  afterEach(() => {
    delete process.env.TEAMS_SP_TOKEN;
  });

  function cache(host: string, secret: string, expires_on: number) {
    saveTokens({ sharepoint: { [host]: { secret, expires_on } } });
  }

  it("prefers TEAMS_SP_TOKEN over the cache", async () => {
    cache("fabrikam-my.sharepoint.com", CACHED_TOKEN, FUTURE);
    process.env.TEAMS_SP_TOKEN = ` ${ENV_TOKEN}\n`;
    expect(await getOneDriveToken()).toEqual({
      token: ENV_TOKEN,
      source: "env",
    });
  });

  it("falls back to a cached OneDrive-host token", async () => {
    cache("fabrikam-my.sharepoint.com", CACHED_TOKEN, FUTURE);
    expect(peekOneDriveToken()).toEqual({
      token: CACHED_TOKEN,
      host: "fabrikam-my.sharepoint.com",
      source: "cache",
    });
  });

  it("ignores expired and non-OneDrive cache entries", async () => {
    saveTokens({
      sharepoint: {
        "fabrikam-my.sharepoint.com": { secret: CACHED_TOKEN, expires_on: 1 },
        "fabrikam.sharepoint.com": { secret: CACHED_TOKEN, expires_on: FUTURE },
      },
    });
    expect(peekOneDriveToken()).toBeNull();
    // Browser steps are unavailable in tests, so the ladder yields nothing.
    expect(await getOneDriveToken()).toBeNull();
  });
});
