import { describe, expect, it, vi } from "vitest";
import {
  createDmThread,
  getActivity,
  getConversation,
  getMessages,
  getThreadMembers,
  listConversations,
  markAsRead,
  pollConversations,
  pollMessages,
  searchMessages,
  searchPeople,
  sendMessage,
} from "../src/api.js";
import type { ApiClient } from "../src/client.js";

const BASE = "https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME";

vi.mock("../src/client.js", () => ({
  getChatsvcBaseUrl: () => BASE,
}));

function mockClient(
  responseBody: unknown,
  opts?: { headers?: Headers },
): ApiClient {
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
    const [_url, init] = fetchFn.mock.calls[0];
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

describe("HTTP error handling", () => {
  it("throws on non-ok response", async () => {
    const client: ApiClient = {
      baseUrl: BASE,
      fetch: vi.fn().mockResolvedValue({
        ok: false,
        status: 403,
        json: () => Promise.resolve({}),
        text: () => Promise.resolve("Forbidden"),
        headers: new Headers(),
      }),
    };
    await expect(listConversations(client)).rejects.toThrow("HTTP 403");
  });

  it("includes response body in error", async () => {
    const client: ApiClient = {
      baseUrl: BASE,
      fetch: vi.fn().mockResolvedValue({
        ok: false,
        status: 500,
        json: () => Promise.resolve({}),
        text: () => Promise.resolve("Internal Server Error"),
        headers: new Headers(),
      }),
    };
    await expect(getMessages(client, "19:abc")).rejects.toThrow(
      "Internal Server Error",
    );
  });

  it("handles text() rejection gracefully", async () => {
    const client: ApiClient = {
      baseUrl: BASE,
      fetch: vi.fn().mockResolvedValue({
        ok: false,
        status: 502,
        json: () => Promise.reject(new Error("no json")),
        text: () => Promise.reject(new Error("no body")),
        headers: new Headers(),
      }),
    };
    await expect(listConversations(client)).rejects.toThrow("HTTP 502");
  });
});

describe("getConversation", () => {
  it("fetches a single conversation", async () => {
    const client = mockClient({ id: "19:abc", threadProperties: {} });
    const result = await getConversation(client, "19:abc");
    expect(result.id).toBe("19:abc");
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    expect(fetchFn.mock.calls[0][0]).toContain("/conversations/19:abc");
  });
});

describe("createDmThread", () => {
  it("extracts thread ID from location header", async () => {
    const headers = new Headers();
    headers.set(
      "location",
      `${BASE.replace("/users/ME", "")}/threads/19:new-thread@thread.v2`,
    );
    const client: ApiClient = {
      baseUrl: BASE,
      fetch: vi.fn().mockResolvedValue({
        ok: true,
        text: () => Promise.resolve(""),
        headers,
      }),
    };
    const id = await createDmThread(client, "8:orgid:aaa", "8:orgid:bbb");
    expect(id).toBe("19:new-thread@thread.v2");
  });

  it("falls back to deterministic ID when no location header", async () => {
    const client: ApiClient = {
      baseUrl: BASE,
      fetch: vi.fn().mockResolvedValue({
        ok: true,
        text: () => Promise.resolve(""),
        headers: new Headers(),
      }),
    };
    const id = await createDmThread(client, "8:orgid:bbb", "8:orgid:aaa");
    // UUIDs are sorted, so aaa comes first
    expect(id).toBe("19:aaa_bbb@unq.gbl.spaces");
  });

  it("sends correct member payload", async () => {
    const client: ApiClient = {
      baseUrl: BASE,
      fetch: vi.fn().mockResolvedValue({
        ok: true,
        text: () => Promise.resolve(""),
        headers: new Headers(),
      }),
    };
    await createDmThread(client, "8:orgid:me", "8:orgid:them");
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    const [url, init] = fetchFn.mock.calls[0];
    expect(url).toContain("/threads");
    expect(url).not.toContain("/users/ME");
    const body = JSON.parse(init.body);
    expect(body.members).toHaveLength(2);
    expect(body.properties.uniquerosterthread).toBe(true);
  });
});

describe("markAsRead", () => {
  it("sends correct horizon format", async () => {
    const client: ApiClient = {
      baseUrl: BASE,
      fetch: vi.fn().mockResolvedValue({
        ok: true,
        text: () => Promise.resolve(""),
        headers: new Headers(),
      }),
    };
    await markAsRead(client, "19:abc", "msg123", "ver456");
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    const [url, init] = fetchFn.mock.calls[0];
    expect(url).toContain("/conversations/19:abc/properties");
    expect(url).toContain("consumptionhorizon");
    const body = JSON.parse(init.body);
    expect(body.consumptionhorizon).toBe("msg123;ver456;0");
  });
});

describe("pollMessages", () => {
  it("uses syncUrl when provided", async () => {
    const syncUrl = "https://example.com/sync?state=abc";
    const client = mockClient({
      messages: [{ id: "m1" }],
      _metadata: { syncState: "https://example.com/sync?state=def" },
    });
    const result = await pollMessages(client, "19:abc", syncUrl);
    expect(result.messages).toHaveLength(1);
    expect(result.syncUrl).toBe("https://example.com/sync?state=def");
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    expect(fetchFn.mock.calls[0][0]).toBe(syncUrl);
  });

  it("builds initial URL with startTime when no syncUrl", async () => {
    const client = mockClient({
      messages: [],
      _metadata: { syncState: "https://example.com/sync?state=init" },
    });
    const result = await pollMessages(client, "19:abc");
    expect(result.messages).toHaveLength(0);
    expect(result.syncUrl).toBe("https://example.com/sync?state=init");
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    expect(fetchFn.mock.calls[0][0]).toContain("startTime=");
  });
});

describe("pollConversations", () => {
  it("uses syncUrl when provided", async () => {
    const syncUrl = "https://example.com/conv-sync?state=abc";
    const client = mockClient({
      conversations: [{ id: "19:c1" }],
      _metadata: { syncState: "https://example.com/conv-sync?state=def" },
    });
    const result = await pollConversations(client, syncUrl);
    expect(result.conversations).toHaveLength(1);
    expect(result.syncUrl).toBe("https://example.com/conv-sync?state=def");
  });

  it("builds initial URL when no syncUrl", async () => {
    const client = mockClient({
      conversations: [],
      _metadata: { syncState: "" },
    });
    const result = await pollConversations(client);
    expect(result.conversations).toHaveLength(0);
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    expect(fetchFn.mock.calls[0][0]).toContain("startTime=");
  });
});

describe("searchPeople", () => {
  it("returns people suggestions", async () => {
    const client = mockClient({
      Groups: [
        {
          Type: "People",
          Suggestions: [{ DisplayName: "Alice", MRI: "8:orgid:alice" }],
        },
      ],
    });
    const result = await searchPeople(client, "alice");
    expect(result).toHaveLength(1);
    expect(result[0].DisplayName).toBe("Alice");
  });

  it("returns empty when no People group", async () => {
    const client = mockClient({ Groups: [] });
    const result = await searchPeople(client, "nobody");
    expect(result).toHaveLength(0);
  });
});

describe("searchMessages", () => {
  it("returns search results", async () => {
    const client = mockClient({ entitySets: [{ resultSets: [] }] });
    const result = await searchMessages(client, "test query");
    expect(result.entitySets).toBeDefined();
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    const [url, init] = fetchFn.mock.calls[0];
    expect(url).toContain("/searchservice/api/v2/query");
    const body = JSON.parse(init.body);
    expect(body.entityRequests[0].query.queryString).toBe("test query");
  });
});
