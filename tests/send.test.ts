import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const BASE = "https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME";
const fetchMock = vi.fn();
const getChatClient = vi.fn(async () => ({ baseUrl: BASE, fetch: fetchMock }));

vi.mock("../src/client.js", () => ({
  getChatClient: () => getChatClient(),
  getSearchClient: vi.fn(),
  getStreamClient: vi.fn(),
}));

const { createProgram, resolveSendContent } = await import("../src/cli.js");

const noIo = {
  readFile: () => {
    throw new Error("readFile should not be called");
  },
  readStdin: () => Promise.reject(new Error("readStdin should not be called")),
};

describe("resolveSendContent", () => {
  it("returns plain text for a positional message", async () => {
    expect(await resolveSendContent("hi", {}, noIo)).toEqual({
      content: "hi",
      html: false,
    });
  });

  it("returns --html verbatim", async () => {
    expect(
      await resolveSendContent(undefined, { html: "<p>a</p>" }, noIo),
    ).toEqual({ content: "<p>a</p>", html: true });
  });

  it("reads html from stdin when --html is '-'", async () => {
    const io = { ...noIo, readStdin: async () => "<div><p>x<br>y</p></div>\n" };
    expect(await resolveSendContent(undefined, { html: "-" }, io)).toEqual({
      content: "<div><p>x<br>y</p></div>\n",
      html: true,
    });
  });

  it("reads html from --html-file", async () => {
    const io = { ...noIo, readFile: (p: string) => `<ol><li>${p}</li></ol>` };
    expect(await resolveSendContent(undefined, { htmlFile: "f" }, io)).toEqual({
      content: "<ol><li>f</li></ol>",
      html: true,
    });
  });

  it("errors when both a positional and an html flag are given", async () => {
    const r = await resolveSendContent("hi", { html: "<p>a</p>" }, noIo);
    expect(r).toHaveProperty("error");
    expect((r as { error: string }).error).toMatch(/only one of/);
  });

  it("errors when --html and --html-file are both given", async () => {
    const r = await resolveSendContent(
      undefined,
      { html: "<p>a</p>", htmlFile: "f" },
      noIo,
    );
    expect(r).toHaveProperty("error");
  });

  it("errors when nothing is given", async () => {
    expect(await resolveSendContent(undefined, {}, noIo)).toHaveProperty(
      "error",
    );
  });
});

describe("send command", () => {
  let out: string[];
  let err: string[];

  beforeEach(() => {
    out = [];
    err = [];
    fetchMock.mockReset();
    getChatClient.mockClear();
    vi.spyOn(console, "log").mockImplementation((s) => out.push(String(s)));
    vi.spyOn(console, "error").mockImplementation((s) => err.push(String(s)));
    vi.spyOn(process, "exit").mockImplementation(((code?: number) => {
      throw new Error(`exit ${code}`);
    }) as never);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  async function run(...args: string[]) {
    const program = createProgram();
    program.exitOverride();
    await program.parseAsync(["node", "teams", "send", ...args]);
  }

  it("--dry-run prints the html body and makes no client call", async () => {
    const dir = mkdtempSync(join(tmpdir(), "teams-send-"));
    const file = join(dir, "msg.html");
    const html = "<div><p>line1<br>line2</p></div>";
    writeFileSync(file, html);

    await run("19:abc@thread.v2", "--html-file", file, "--dry-run");

    expect(getChatClient).not.toHaveBeenCalled();
    expect(fetchMock).not.toHaveBeenCalled();
    const printed = JSON.parse(out.join("\n"));
    expect(printed.conversationId).toBe("19:abc@thread.v2");
    expect(printed.body.content).toBe(html);
    expect(printed.body.messagetype).toBe("RichText/Html");
    expect(printed.body.contenttype).toBe("text");
    expect(printed.body.clientmessageid).toMatch(/^\d{19}$/);
  });

  it("--dry-run works in plain mode", async () => {
    await run("19:abc@thread.v2", "hello", "--dry-run");
    expect(getChatClient).not.toHaveBeenCalled();
    const printed = JSON.parse(out.join("\n"));
    expect(printed.body.content).toBe("<p>hello</p>");
  });

  it("errors on positional message plus --html, sending nothing", async () => {
    await expect(
      run("19:abc@thread.v2", "hello", "--html", "<p>x</p>"),
    ).rejects.toThrow("exit 1");
    expect(err.join("\n")).toMatch(/only one of/);
    expect(getChatClient).not.toHaveBeenCalled();
  });

  it("posts --html verbatim and prints the arrival receipt", async () => {
    fetchMock.mockResolvedValue({
      ok: true,
      json: async () => ({ OriginalArrivalTime: 1700000000000 }),
    });
    await run("19:abc@thread.v2", "--html", "<ol><li>a</li></ol>");
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe(`${BASE}/conversations/19:abc@thread.v2/messages`);
    const body = JSON.parse(init.body);
    expect(body.content).toBe("<ol><li>a</li></ol>");
    expect(body.messagetype).toBe("RichText/Html");
    expect(out.join("\n")).toContain("Arrival: 1700000000000");
  });

  it("plain send posts the unchanged <p>-wrapped body", async () => {
    fetchMock.mockResolvedValue({
      ok: true,
      json: async () => ({ OriginalArrivalTime: 1 }),
    });
    await run("19:abc@thread.v2", "hello");
    const body = JSON.parse(fetchMock.mock.calls[0][1].body);
    expect(body).toEqual({
      content: "<p>hello</p>",
      messagetype: "RichText/Html",
      contenttype: "text",
      clientmessageid: expect.stringMatching(/^\d{19}$/),
    });
  });
});
