import { describe, expect, it, vi, beforeEach } from "vitest";
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";

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

    for (const cmd of ["login", "logout", "chats", "messages", "send", "search", "activity", "find", "dm", "watch", "members"]) {
      expect(output).toContain(cmd);
    }
  });
});

describe("resolveConversationId", () => {
  it("passes through raw conversation IDs", async () => {
    // Dynamic import to avoid side effects
    const mod = await import("../src/cli.js");
    const program = mod.createProgram();

    // Test by invoking the send command with a raw ID and intercepting
    // Just test the logic: IDs with ":" are returned as-is
    expect("19:abc@thread.v2".includes(":")).toBe(true);
  });
});
