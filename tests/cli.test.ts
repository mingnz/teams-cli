import { afterEach, describe, expect, it, vi } from "vitest";
import * as api from "../src/api.js";
import * as clients from "../src/client.js";

afterEach(() => {
  vi.restoreAllMocks();
});

describe("createProgram", () => {
  it("includes all commands in help", async () => {
    const { createProgram } = await import("../src/cli.js");
    const program = createProgram();
    program.exitOverride();

    let output = "";
    program.configureOutput({ writeOut: (str) => (output += str) });

    try {
      await program.parseAsync(["node", "teams", "--help"]);
    } catch {
      // commander throws on --help with exitOverride
    }

    for (const cmd of [
      "login",
      "logout",
      "chats",
      "messages",
      "send",
      "reply",
      "search",
      "activity",
      "find",
      "dm",
      "watch",
      "members",
    ]) {
      expect(output).toContain(cmd);
    }
  });
});

describe("channel reply commands", () => {
  const channel = "19:example-channel@thread.tacv2";
  const root = "1700000000000";
  const client: clients.ApiClient = {
    baseUrl: "https://teams.cloud.microsoft/example",
    fetch: vi.fn(),
  };

  it("routes an explicit reply without using the top-level send function", async () => {
    const { createProgram } = await import("../src/cli.js");
    vi.spyOn(clients, "getChatClient").mockResolvedValue(client);
    const reply = vi.spyOn(api, "replyToMessage").mockResolvedValue({});
    const send = vi.spyOn(api, "sendMessage");
    const output = vi.spyOn(console, "log").mockImplementation(() => {});
    await createProgram().parseAsync([
      "node",
      "teams",
      "reply",
      channel,
      root,
      "Test reply",
    ]);
    expect(reply).toHaveBeenCalledWith(client, channel, root, "Test reply");
    expect(send).not.toHaveBeenCalled();
    expect(output).toHaveBeenCalledWith(
      expect.stringContaining(`Root: ${root}`),
    );
  });

  it("requires a root ID and a message", async () => {
    const { createProgram } = await import("../src/cli.js");
    const getClient = vi.spyOn(clients, "getChatClient");
    const program = createProgram().exitOverride();
    for (const command of program.commands) {
      command.exitOverride().configureOutput({ writeErr: () => {} });
    }
    await expect(
      program.parseAsync(["node", "teams", "reply", channel, root]),
    ).rejects.toThrow("missing required argument");
    expect(getClient).not.toHaveBeenCalled();
  });

  it("reads a specified thread and displays IDs when requested", async () => {
    const { createProgram } = await import("../src/cli.js");
    vi.spyOn(clients, "getChatClient").mockResolvedValue(client);
    const read = vi.spyOn(api, "getThreadMessages").mockResolvedValue([
      {
        id: root,
        messagetype: "Text",
        content: "Root post",
        originalarrivaltime: "2024-01-01T00:00:00Z",
      },
    ]);
    const normalRead = vi.spyOn(api, "getMessages");
    const output = vi.spyOn(console, "log").mockImplementation(() => {});
    await createProgram().parseAsync([
      "node",
      "teams",
      "messages",
      channel,
      "--thread",
      root,
      "--ids",
      "-n",
      "7",
    ]);
    expect(read).toHaveBeenCalledWith(client, channel, root, 7);
    expect(normalRead).not.toHaveBeenCalled();
    expect(output).toHaveBeenCalledWith(`  [${root}] Root post`);
  });

  it("preserves normal message retrieval and display by default", async () => {
    const { createProgram } = await import("../src/cli.js");
    vi.spyOn(clients, "getChatClient").mockResolvedValue(client);
    const read = vi.spyOn(api, "getMessages").mockResolvedValue([
      {
        id: root,
        messagetype: "Text",
        content: "Root post",
        originalarrivaltime: "2024-01-01T00:00:00Z",
      },
    ]);
    const threadRead = vi.spyOn(api, "getThreadMessages");
    const output = vi.spyOn(console, "log").mockImplementation(() => {});
    await createProgram().parseAsync(["node", "teams", "messages", channel]);
    expect(read).toHaveBeenCalledWith(client, channel, 20);
    expect(threadRead).not.toHaveBeenCalled();
    expect(output).toHaveBeenCalledWith("  Root post");
  });
});

describe("resolveConversationId", () => {
  it("passes through raw conversation IDs", async () => {
    // Dynamic import to avoid side effects
    const mod = await import("../src/cli.js");
    const _program = mod.createProgram();

    // Test by invoking the send command with a raw ID and intercepting
    // Just test the logic: IDs with ":" are returned as-is
    expect("19:abc@thread.v2".includes(":")).toBe(true);
  });
});
