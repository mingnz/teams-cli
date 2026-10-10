import { describe, expect, it, vi } from "vitest";
import { resolveConversationId } from "../src/cli.js";
import type { FormattedChat } from "../src/formatting.js";

function chat(id: string, short_id: string): FormattedChat {
  return { short_id, name: `Chat ${short_id}`, type: "Chat", preview: "", time: "", id };
}

const OSAMA = "19:63d43783-f462-431d-9a35-02c4bc67344e_b5bb788f-e279-4f3d-acff-0790cbe8ea6f@unq.gbl.spaces";
const OTHER = "19:76559d43-ef26-489f-a90b-06467794080b_11111111-2222-3333-4444-555555555555@unq.gbl.spaces";

describe("resolveConversationId", () => {
  it("passes through full conversation IDs", () => {
    expect(resolveConversationId(OSAMA, [])).toBe(OSAMA);
  });

  it("resolves an exact short id", () => {
    expect(resolveConversationId("ab12", [chat(OSAMA, "ab12"), chat(OTHER, "cd34")])).toBe(OSAMA);
  });

  it("resolves an exact short id when the thread has duplicate index rows", () => {
    const idx = [chat(OSAMA, "ab12"), chat(OSAMA, "zz99"), chat(OTHER, "cd34")];
    expect(resolveConversationId("ab12", idx)).toBe(OSAMA);
  });

  it("resolves a stale short id via unique-suffix fallback", () => {
    // `chats` re-assigned the short id since it was captured; the old
    // value is still a tail of the new assignment.
    const idx = [chat(OSAMA, "01ab12"), chat(OTHER, "cd34")];
    expect(resolveConversationId("ab12", idx)).toBe(OSAMA);
  });

  it("resolves a unique suffix of duplicate rows that share one thread", () => {
    // The conversations API returns duplicate rows for the same thread
    // with different short ids; a shared tail must still resolve.
    const idx = [chat(OSAMA, "c3ea6f"), chat(OSAMA, "d7ea6f"), chat(OTHER, "cd34")];
    expect(resolveConversationId("ea6f", idx)).toBe(OSAMA);
  });

  it("resolves a unique suffix of the conversation id body", () => {
    const idx = [chat(OSAMA, "ab12"), chat(OTHER, "cd34")];
    expect(resolveConversationId("cbe8ea6f", idx)).toBe(OSAMA);
  });

  it("lists candidates when a suffix matches different threads", () => {
    const err = vi.spyOn(console, "error").mockImplementation(() => {});
    const exit = vi.spyOn(process, "exit").mockImplementation(() => undefined as never);
    try {
      const idx = [chat(OSAMA, "c3ea6f"), chat(OTHER, "d7ea6f")];
      resolveConversationId("ea6f", idx);
      expect(exit).toHaveBeenCalledWith(1);
      expect(err).toHaveBeenCalledWith(expect.stringContaining("matches multiple conversations"));
    } finally {
      err.mockRestore();
      exit.mockRestore();
    }
  });

  it("exits when nothing matches", () => {
    const err = vi.spyOn(console, "error").mockImplementation(() => {});
    const exit = vi.spyOn(process, "exit").mockImplementation(() => undefined as never);
    try {
      resolveConversationId("ffff", [chat(OSAMA, "ab12")]);
      expect(exit).toHaveBeenCalledWith(1);
    } finally {
      err.mockRestore();
      exit.mockRestore();
    }
  });

  it("does not suffix-match inputs shorter than 4 characters", () => {
    const exit = vi.spyOn(process, "exit").mockImplementation(() => undefined as never);
    try {
      resolveConversationId("f6f", [chat(OSAMA, "c3f6f")]);
      expect(exit).toHaveBeenCalledWith(1);
    } finally {
      exit.mockRestore();
    }
  });
});
