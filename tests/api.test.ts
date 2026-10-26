import { describe, expect, it, vi } from "vitest";
import {
  createDmThread,
  downloadTranscriptJson,
  getActivity,
  getConversation,
  getMessages,
  getThreadMembers,
  getThreadMessages,
  getTranscriptMetadata,
  listConversations,
  markAsRead,
  pollConversations,
  pollMessages,
  replyToMessage,
  resolveDriveItem,
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
    expect(_url).toBe(`${BASE}/conversations/19:abc/messages`);
    expect(body).not.toHaveProperty("replyChainId");
    expect(body).not.toHaveProperty("conversationid");
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

describe("channel replies", () => {
  const channel = "19:example-channel@thread.tacv2";
  const root = "1700000000000";
  const thread = encodeURIComponent(`${channel};messageid=${root}`);

  it("posts to the root-scoped conversation with explicit thread linkage", async () => {
    const client = mockClient({ OriginalArrivalTime: "1700000000001" });
    const result = await replyToMessage(client, channel, root, "Test reply");
    expect(result.OriginalArrivalTime).toBe("1700000000001");
    expect(client.fetch).toHaveBeenCalledExactlyOnceWith(
      `${BASE}/conversations/${thread}/messages`,
      expect.objectContaining({ method: "POST" }),
    );
    const [, init] = vi.mocked(client.fetch).mock.calls[0];
    const body = JSON.parse(init?.body as string);
    expect(body).toEqual({
      content: "<p>Test reply</p>",
      messagetype: "RichText/Html",
      contenttype: "Text",
      clientmessageid: expect.stringMatching(/^\d{19}$/),
      conversationid: channel,
      replyChainId: root,
      properties: { formatVariant: "TEAMS" },
    });
  });

  it("reads the same root-scoped conversation and preserves the page size", async () => {
    const client = mockClient({ messages: [{ id: root }] });
    expect(await getThreadMessages(client, channel, root, 7)).toEqual([
      { id: root },
    ]);
    const [url] = vi.mocked(client.fetch).mock.calls[0];
    expect(url).toContain(`/conversations/${thread}/messages?`);
    expect(new URL(url).searchParams.get("pageSize")).toBe("7");
  });

  it.each([
    "",
    "0",
    "-1",
    "1.5",
    "1e12",
    "123;messageid=456",
    "123/456",
  ])("rejects invalid root ID %j without making a request", async (id) => {
    const client = mockClient({});
    await expect(replyToMessage(client, channel, id, "Reply")).rejects.toThrow(
      "Root message ID",
    );
    await expect(getThreadMessages(client, channel, id)).rejects.toThrow(
      "Root message ID",
    );
    expect(client.fetch).not.toHaveBeenCalled();
  });

  it.each([
    "19:example@thread.v2",
    `${channel};messageid=${root}`,
    "19:example/path@thread.tacv2",
    "19:example%2fpath@thread.tacv2",
  ])("rejects non-channel or compound ID %j", async (id) => {
    const client = mockClient({});
    await expect(replyToMessage(client, id, root, "Reply")).rejects.toThrow(
      "channel ID",
    );
    expect(client.fetch).not.toHaveBeenCalled();
  });

  it("rejects empty replies before making a request", async () => {
    const client = mockClient({});
    await expect(replyToMessage(client, channel, root, " \n ")).rejects.toThrow(
      "must not be empty",
    );
    expect(client.fetch).not.toHaveBeenCalled();
  });

  it("surfaces errors without retrying or falling back to a top-level post", async () => {
    const client = mockClient({});
    vi.mocked(client.fetch).mockResolvedValue(
      new Response("Forbidden", { status: 403 }),
    );
    await expect(
      replyToMessage(client, channel, root, "Reply"),
    ).rejects.toThrow("HTTP 403: Forbidden");
    expect(client.fetch).toHaveBeenCalledTimes(1);
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

describe("resolveDriveItem", () => {
  it("encodes the file URL as a share token and extracts ids", async () => {
    const client = mockClient({
      id: "ITEM123",
      parentReference: { driveId: "DRIVE456" },
    });
    const fileUrl = "https://contoso-my.sharepoint.com/personal/u/clip.mp4";
    const result = await resolveDriveItem(client, fileUrl);
    expect(result).toEqual({ driveId: "DRIVE456", itemId: "ITEM123" });

    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    const [url] = fetchFn.mock.calls[0];
    const expectedShare = `u!${Buffer.from(fileUrl).toString("base64url")}`;
    expect(url).toBe(`/_api/v2.0/shares/${expectedShare}/driveItem`);
  });

  it("throws when ids are missing", async () => {
    const client = mockClient({});
    await expect(
      resolveDriveItem(client, "https://x.sharepoint.com/a"),
    ).rejects.toThrow();
  });
});

describe("getTranscriptMetadata", () => {
  it("reads the expanded media.transcripts shape", async () => {
    const client = mockClient({
      media: { transcripts: [{ temporaryDownloadUrl: "https://dl/1" }] },
    });
    const meta = await getTranscriptMetadata(client, "d", "i");
    expect(meta?.temporaryDownloadUrl).toBe("https://dl/1");

    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    const [url] = fetchFn.mock.calls[0];
    expect(url).toContain("/_api/v2.1/drives/d/items/i");
    expect(url).toContain("transcripts");
  });

  it("reads the value collection shape and prefers isDefault", async () => {
    const client = mockClient({
      value: [
        { temporaryDownloadUrl: "https://dl/a" },
        { temporaryDownloadUrl: "https://dl/b", isDefault: true },
      ],
    });
    const meta = await getTranscriptMetadata(client, "d", "i");
    expect(meta?.temporaryDownloadUrl).toBe("https://dl/b");
  });

  it("reads the single-object shape", async () => {
    const client = mockClient({ temporaryDownloadUrl: "https://dl/solo" });
    const meta = await getTranscriptMetadata(client, "d", "i");
    expect(meta?.temporaryDownloadUrl).toBe("https://dl/solo");
  });

  it("returns null when no transcript exists", async () => {
    const client = mockClient({ media: { transcripts: [] } });
    const meta = await getTranscriptMetadata(client, "d", "i");
    expect(meta).toBeNull();
  });
});

describe("downloadTranscriptJson", () => {
  it("appends format=json and returns the body text", async () => {
    const client = mockClient({ entries: [] });
    const text = await downloadTranscriptJson(client, "https://dl/x?token=1");
    expect(JSON.parse(text)).toEqual({ entries: [] });

    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    const [url] = fetchFn.mock.calls[0];
    expect(url).toBe("https://dl/x?token=1&format=json");
  });

  it("uses '?' when the url has no query string", async () => {
    const client = mockClient({ entries: [] });
    await downloadTranscriptJson(client, "https://dl/y");
    const fetchFn = client.fetch as ReturnType<typeof vi.fn>;
    expect(fetchFn.mock.calls[0][0]).toBe("https://dl/y?format=json");
  });
});
