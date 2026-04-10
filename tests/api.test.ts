import { describe, expect, it, vi, beforeEach } from "vitest";
import type { ApiClient } from "../src/client.js";
import {
  getActivity,
  getMessages,
  getThreadMembers,
  listConversations,
  sendMessage,
} from "../src/api.js";

const BASE = "https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME";

vi.mock("../src/client.js", () => ({
  getChatsvcBaseUrl: () => BASE,
}));

function mockClient(responseBody: unknown, opts?: { headers?: Headers }): ApiClient {
  return {
    baseUrl: BASE,
    fetch: vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve(responseBody),
      text: () => Promise.resolve(JSON.stringify(responseBody)),
      headers: opts?.headers ?? new Headers(),
    }),
  };
}

describe("listConversations", () => {
  it("returns conversations", async () => {
    const client = mockClient({ conversations: [{ id: "19:abc" }] });
    const result = await listConversations(client, 10);
    expect(result).toHaveLength(1);
    expect(result[0].id).toBe("19:abc");
  });
});

describe("getMessages", () => {
  it("returns messages", async () => {
    const client = mockClient({ messages: [{ id: "1", content: "hi" }] });
    const result = await getMessages(client, "19:abc", 5);
    expect(result).toHaveLength(1);
  });
});

describe("sendMessage", () => {
  it("sends and returns result", async () => {
    const client = mockClient({ OriginalArrivalTime: "123" });
    const result = await sendMessage(client, "19:abc", "hello");
    expect(result.OriginalArrivalTime).toBe("123");
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    const [url, init] = fetchFn.mock.calls[0];
    const body = JSON.parse(init.body);
    expect(body.content).toBe("<p>hello</p>");
    expect(body.messagetype).toBe("RichText/Html");
  });

  it("handles special characters safely", async () => {
    const client = mockClient({ OriginalArrivalTime: "456" });
    await sendMessage(client, "19:abc", 'He said "hello" & {test}');
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    const [, init] = fetchFn.mock.calls[0];
    const body = JSON.parse(init.body);
    expect(body.content).toBe('<p>He said "hello" & {test}</p>');
  });
});

describe("getThreadMembers", () => {
  it("returns members", async () => {
    const client = mockClient({ members: [{ id: "8:orgid:user1" }] });
    const result = await getThreadMembers(client, "19:abc");
    expect(result).toHaveLength(1);
  });
});

describe("getActivity", () => {
  it("returns activity", async () => {
    const client = mockClient({ messages: [{ id: "n1" }] });
    const result = await getActivity(client, "notifications", 5);
    expect(result).toHaveLength(1);
  });
});
