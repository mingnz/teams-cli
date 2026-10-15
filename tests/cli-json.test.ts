import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const DATA_DIR = mkdtempSync(join(tmpdir(), "teams-cli-json-"));

vi.mock("../src/config.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/config.js")>();
  return {
    ...actual,
    DATA_DIR,
    LAST_CHATS_FILE: join(DATA_DIR, "last_chats.json"),
  };
});

vi.mock("../src/client.js", () => ({
  getChatClient: vi.fn(async () => ({
    baseUrl: "https://test",
    fetch: vi.fn(),
  })),
  getSearchClient: vi.fn(),
  getStreamClient: vi.fn(),
}));

vi.mock("../src/api.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/api.js")>();
  return {
    ...actual,
    listConversations: vi.fn(),
    getMessages: vi.fn(),
  };
});

const { getMessages, listConversations } = await import("../src/api.js");
const { createProgram } = await import("../src/cli.js");

const CONV_ID = "19:aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee@thread.v2";

const conversations = [
  {
    id: CONV_ID,
    threadProperties: { topic: "Test Group", threadType: "chat" },
    properties: { lastimreceivedtime: "2026-01-01T10:00:00.000Z" },
    lastMessage: {
      id: "1767261600000",
      content: "<p>hello</p>",
      imdisplayname: "Test User",
      originalarrivaltime: "2026-01-01T10:00:00.000Z",
    },
  },
  { id: "48:notifications", properties: {}, lastMessage: {} },
];

// Newest first, as the API returns them.
const messages = [
  {
    id: "1767261600001",
    version: "1767261600001",
    messagetype: "RichText/Html",
    content: "<p>second</p>",
    imdisplayname: "Test User",
    originalarrivaltime: "2026-01-01T10:00:01.000Z",
    amsreferences: ["0-wus-d1-00000000000000000000000000000000"],
    properties: {
      files: JSON.stringify([
        {
          fileName: "image.png",
          filePreview: { previewUrl: "https://example.invalid/preview.png" },
        },
      ]),
    },
  },
  {
    id: "1767261600000",
    version: "1767261600000",
    messagetype: "RichText/Html",
    content: "<p>first</p>",
    imdisplayname: "Test User",
    originalarrivaltime: "2026-01-01T10:00:00.000Z",
    properties: {},
  },
];

let stdout: string;

async function run(...args: string[]): Promise<void> {
  const program = createProgram();
  program.exitOverride();
  await program.parseAsync(["node", "teams", ...args]);
}

beforeEach(() => {
  stdout = "";
  vi.spyOn(process.stdout, "write").mockImplementation((chunk) => {
    stdout += String(chunk);
    return true;
  });
  vi.spyOn(console, "log").mockImplementation((...a: unknown[]) => {
    stdout += `${a.map(String).join(" ")}\n`;
  });
  vi.spyOn(console, "error").mockImplementation(() => {});
  vi.mocked(listConversations).mockResolvedValue(conversations);
  vi.mocked(getMessages).mockResolvedValue(messages);
});

afterEach(() => {
  vi.restoreAllMocks();
});

process.on("exit", () => rmSync(DATA_DIR, { recursive: true, force: true }));

describe("chats --json", () => {
  it("prints only a JSON array of raw conversation objects", async () => {
    await run("chats", "--json", "-n", "5");

    const parsed = JSON.parse(stdout);
    expect(Array.isArray(parsed)).toBe(true);
    expect(parsed).toEqual(conversations);
    expect(stdout.trim().startsWith("[")).toBe(true);
    expect(stdout.trim().endsWith("]")).toBe(true);
    expect(stdout).not.toContain("\u001b[");
    expect(listConversations).toHaveBeenCalledWith(expect.anything(), 5);
  });

  it("still saves the chat index", async () => {
    await run("chats", "--json");

    const index = JSON.parse(
      readFileSync(join(DATA_DIR, "last_chats.json"), "utf-8"),
    );
    expect(index).toHaveLength(1);
    expect(index[0].id).toBe(CONV_ID);
  });

  it("keeps stdout clean on API errors", async () => {
    vi.mocked(listConversations).mockRejectedValue(new Error("HTTP 401: nope"));

    await expect(run("chats", "--json")).rejects.toThrow("HTTP 401");
    expect(stdout).toBe("");
  });
});

describe("chats (table)", () => {
  it("renders the table unchanged without --json", async () => {
    await run("chats");

    expect(() => JSON.parse(stdout)).toThrow();
    for (const col of ["ID", "Name", "Type", "Last Message", "Time"]) {
      expect(stdout).toContain(col);
    }
    expect(stdout).toContain("Test Group");
    expect(stdout).not.toContain("48:notifications");
  });
});

describe("messages --json", () => {
  it("prints only a JSON array of raw message objects", async () => {
    await run("messages", CONV_ID, "--json", "-n", "2");

    const parsed = JSON.parse(stdout);
    expect(Array.isArray(parsed)).toBe(true);
    expect(parsed).toEqual(messages);
    expect(stdout).not.toContain("\u001b[");
    expect(getMessages).toHaveBeenCalledWith(expect.anything(), CONV_ID, 2);
  });

  it("preserves properties, amsreferences, and string ids/versions", async () => {
    await run("messages", CONV_ID, "--json");

    const [first] = JSON.parse(stdout);
    expect(typeof first.id).toBe("string");
    expect(typeof first.version).toBe("string");
    expect(first.amsreferences).toEqual(messages[0].amsreferences);
    const files = JSON.parse(first.properties.files);
    expect(files[0].filePreview.previewUrl).toBe(
      "https://example.invalid/preview.png",
    );
  });

  it("prints an empty array when there are no messages", async () => {
    vi.mocked(getMessages).mockResolvedValue([]);

    await run("messages", CONV_ID, "--json");

    expect(JSON.parse(stdout)).toEqual([]);
  });

  it("keeps stdout clean on API errors", async () => {
    vi.mocked(getMessages).mockRejectedValue(new Error("HTTP 404: gone"));

    await expect(run("messages", CONV_ID, "--json")).rejects.toThrow(
      "HTTP 404",
    );
    expect(stdout).toBe("");
  });
});

describe("messages (table)", () => {
  it("renders oldest-first text output unchanged without --json", async () => {
    await run("messages", CONV_ID);

    expect(() => JSON.parse(stdout)).toThrow();
    expect(stdout).toContain("Test User");
    expect(stdout.indexOf("first")).toBeLessThan(stdout.indexOf("second"));
  });
});
