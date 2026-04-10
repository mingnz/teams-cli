import { describe, expect, it } from "vitest";
import {
  formatChatList,
  formatMember,
  formatMessage,
  formatTimestamp,
  getConversationDisplayName,
  getConversationType,
  makeShortId,
  stripHtml,
} from "../src/formatting.js";

describe("stripHtml", () => {
  it("removes tags", () => {
    expect(stripHtml("<p>Hello <b>world</b></p>")).toBe("Hello world");
  });
  it("handles empty", () => {
    expect(stripHtml("")).toBe("");
  });
  it("handles no tags", () => {
    expect(stripHtml("plain text")).toBe("plain text");
  });
});

describe("formatTimestamp", () => {
  it("formats ISO date", () => {
    expect(formatTimestamp("2025-01-15T10:30:00.000Z")).toBe("2025-01-15 10:30");
  });
  it("returns empty for null", () => {
    expect(formatTimestamp(null)).toBe("");
  });
  it("returns empty for empty string", () => {
    expect(formatTimestamp("")).toBe("");
  });
  it("falls back for invalid date", () => {
    expect(formatTimestamp("not-a-date-but-long-enough")).toBe("not-a-date-but-l");
  });
});

const sampleConversation = {
  id: "19:abc123@thread.v2",
  threadProperties: { topic: "Project Alpha", threadType: "chat" },
  lastMessage: {
    imdisplayname: "Alice",
    content: "<p>Hello world</p>",
    originalarrivaltime: "2025-01-15T10:30:00.000Z",
  },
};

const systemConversation = {
  id: "48:notifications",
  threadProperties: { threadType: "streamofnotifications" },
  lastMessage: {},
};

describe("getConversationDisplayName", () => {
  it("uses topic", () => {
    expect(getConversationDisplayName(sampleConversation)).toBe("Project Alpha");
  });
  it("falls back to sender", () => {
    const conv = { id: "19:abc@thread.v2", threadProperties: {}, lastMessage: { imdisplayname: "Alice" } };
    expect(getConversationDisplayName(conv)).toBe("Chat with Alice");
  });
  it("falls back to id", () => {
    const conv = { id: "19:abc@thread.v2", threadProperties: {}, lastMessage: {} };
    expect(getConversationDisplayName(conv)).toBe("19:abc@thread.v2");
  });
});

describe("getConversationType", () => {
  it("uses product type", () => {
    expect(getConversationType({ threadProperties: { productThreadType: "TeamChannel" } })).toBe("TeamChannel");
  });
  it("maps thread type", () => {
    expect(getConversationType({ threadProperties: { threadType: "meeting" } })).toBe("Meeting");
  });
  it("defaults to Chat", () => {
    expect(getConversationType({ threadProperties: {} })).toBe("Chat");
  });
});

describe("makeShortId", () => {
  it("extracts from thread.v2", () => {
    expect(makeShortId("19:abc123@thread.v2")).toBe("abc123");
  });
  it("extracts from thread.tacv2", () => {
    expect(makeShortId("19:73e9410b1e38453d94aa78274efcf175@thread.tacv2")).toBe(
      "73e9410b1e38453d94aa78274efcf175",
    );
  });
  it("handles no prefix", () => {
    expect(makeShortId("abc@thread.v2")).toBe("abc");
  });
  it("handles plain id", () => {
    expect(makeShortId("someid")).toBe("someid");
  });
});

describe("formatChatList", () => {
  it("skips system conversations", () => {
    const result = formatChatList([sampleConversation, systemConversation]);
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe("Project Alpha");
  });
  it("assigns short ids", () => {
    const result = formatChatList([sampleConversation]);
    expect(result[0].short_id).toBe("c123");
  });
  it("assigns unique short ids", () => {
    const convs = [
      { id: "19:aaa1@thread.v2", threadProperties: {}, lastMessage: {} },
      { id: "19:bbb1@thread.v2", threadProperties: {}, lastMessage: {} },
    ];
    const result = formatChatList(convs);
    const ids = result.map((r) => r.short_id);
    expect(new Set(ids).size).toBe(2);
  });
  it("includes preview", () => {
    const result = formatChatList([sampleConversation]);
    expect(result[0].preview).toContain("Alice");
    expect(result[0].preview).toContain("Hello world");
  });
});

describe("formatMessage", () => {
  it("formats HTML message", () => {
    const msg = {
      id: "1705312200000",
      messagetype: "RichText/Html",
      imdisplayname: "Bob",
      content: "<p>Hey there!</p>",
      originalarrivaltime: "2025-01-15T10:30:00.000Z",
    };
    const result = formatMessage(msg);
    expect(result).not.toBeNull();
    expect(result!.sender).toBe("Bob");
    expect(result!.body).toBe("Hey there!");
  });
  it("skips system message", () => {
    expect(formatMessage({ messagetype: "Event/Call", content: "call started" })).toBeNull();
  });
  it("skips empty content", () => {
    expect(formatMessage({ messagetype: "Text", content: "" })).toBeNull();
  });
});

describe("formatMember", () => {
  it("formats user member", () => {
    const result = formatMember({ id: "8:orgid:user-uuid-123", friendlyName: "Alice Smith", role: "Admin" });
    expect(result.name).toBe("Alice Smith");
    expect(result.type).toBe("User");
    expect(result.role).toBe("Admin");
  });
  it("detects bot", () => {
    const result = formatMember({ id: "28:bot-id", friendlyName: "Helper Bot", role: "User" });
    expect(result.type).toBe("Bot");
  });
  it("falls back to mri for name", () => {
    const result = formatMember({ id: "8:orgid:xyz", role: "User" });
    expect(result.name).toBe("8:orgid:xyz");
  });
});
