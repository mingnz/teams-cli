import { describe, expect, it } from "vitest";
import {
  convertTranscriptToGrouped,
  convertTranscriptToVtt,
  formatChatList,
  formatMember,
  formatMessage,
  formatPerson,
  formatTimestamp,
  getConversationDisplayName,
  getConversationType,
  makeShortId,
  parseRecordings,
  secondsToVtt,
  stripHtml,
  timeToSeconds,
  transcriptFilename,
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
    expect(formatTimestamp("2025-01-15T10:30:00.000Z")).toBe(
      "2025-01-15 10:30",
    );
  });
  it("returns empty for null", () => {
    expect(formatTimestamp(null)).toBe("");
  });
  it("returns empty for empty string", () => {
    expect(formatTimestamp("")).toBe("");
  });
  it("falls back for invalid date", () => {
    expect(formatTimestamp("not-a-date-but-long-enough")).toBe(
      "not-a-date-but-l",
    );
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
    expect(getConversationDisplayName(sampleConversation)).toBe(
      "Project Alpha",
    );
  });
  it("falls back to sender", () => {
    const conv = {
      id: "19:abc@thread.v2",
      threadProperties: {},
      lastMessage: { imdisplayname: "Alice" },
    };
    expect(getConversationDisplayName(conv)).toBe("Chat with Alice");
  });
  it("falls back to id", () => {
    const conv = {
      id: "19:abc@thread.v2",
      threadProperties: {},
      lastMessage: {},
    };
    expect(getConversationDisplayName(conv)).toBe("19:abc@thread.v2");
  });
});

describe("getConversationType", () => {
  it("uses product type", () => {
    expect(
      getConversationType({
        threadProperties: { productThreadType: "TeamChannel" },
      }),
    ).toBe("TeamChannel");
  });
  it("maps thread type", () => {
    expect(
      getConversationType({ threadProperties: { threadType: "meeting" } }),
    ).toBe("Meeting");
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
    expect(
      makeShortId("19:73e9410b1e38453d94aa78274efcf175@thread.tacv2"),
    ).toBe("73e9410b1e38453d94aa78274efcf175");
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
    expect(result?.sender).toBe("Bob");
    expect(result?.body).toBe("Hey there!");
  });
  it("skips system message", () => {
    expect(
      formatMessage({ messagetype: "Event/Call", content: "call started" }),
    ).toBeNull();
  });
  it("skips empty content", () => {
    expect(formatMessage({ messagetype: "Text", content: "" })).toBeNull();
  });
});

describe("formatMember", () => {
  it("formats user member", () => {
    const result = formatMember({
      id: "8:orgid:user-uuid-123",
      friendlyName: "Alice Smith",
      role: "Admin",
    });
    expect(result.name).toBe("Alice Smith");
    expect(result.type).toBe("User");
    expect(result.role).toBe("Admin");
  });
  it("detects bot", () => {
    const result = formatMember({
      id: "28:bot-id",
      friendlyName: "Helper Bot",
      role: "User",
    });
    expect(result.type).toBe("Bot");
  });
  it("falls back to mri for name", () => {
    const result = formatMember({ id: "8:orgid:xyz", role: "User" });
    expect(result.name).toBe("8:orgid:xyz");
  });
  it("classifies unknown prefix as Other", () => {
    const result = formatMember({
      id: "99:unknown",
      friendlyName: "Mystery",
      role: "User",
    });
    expect(result.type).toBe("Other");
  });
});

describe("formatPerson", () => {
  it("extracts all fields", () => {
    const result = formatPerson({
      DisplayName: "Jane Doe",
      EmailAddresses: ["jane@example.com", "jd@alt.com"],
      MRI: "8:orgid:jane-uuid",
      JobTitle: "Engineer",
      Department: "R&D",
      CompanyName: "Acme",
      PeopleSubtype: "OrganizationUser",
    });
    expect(result.name).toBe("Jane Doe");
    expect(result.email).toBe("jane@example.com");
    expect(result.mri).toBe("8:orgid:jane-uuid");
    expect(result.title).toBe("Engineer");
    expect(result.department).toBe("R&D");
    expect(result.company).toBe("Acme");
    expect(result.type).toBe("OrganizationUser");
  });

  it("handles missing fields", () => {
    const result = formatPerson({});
    expect(result.name).toBe("");
    expect(result.email).toBe("");
    expect(result.mri).toBe("");
  });
});

describe("formatChatList edge cases", () => {
  it("shows preview without sender name", () => {
    const conv = {
      id: "19:xyz@thread.v2",
      threadProperties: {},
      lastMessage: { content: "<p>System notification</p>" },
    };
    const result = formatChatList([conv]);
    expect(result[0].preview).toBe("System notification");
  });

  it("resolves short ID collisions by increasing length", () => {
    // Two conversations whose IDs differ only in early characters
    // but share the same last 4 characters
    const convs = [
      { id: "19:aaaa1234@thread.v2", threadProperties: {}, lastMessage: {} },
      { id: "19:bbbb1234@thread.v2", threadProperties: {}, lastMessage: {} },
    ];
    const result = formatChatList(convs);
    const ids = result.map((r) => r.short_id);
    // Must be unique despite sharing last 4 chars
    expect(new Set(ids).size).toBe(2);
    // IDs should be longer than 4 to resolve collision
    expect(ids[0].length).toBeGreaterThan(4);
  });
});

describe("formatTimestamp edge cases", () => {
  it("truncates short invalid dates", () => {
    expect(formatTimestamp("bad")).toBe("bad");
  });
});

// A finished CallRecording message (URIObject with a SharePoint sharing link).
function recordingMsg(
  href: string,
  opts: { name?: string; title?: string; time?: string } = {},
): Record<string, unknown> {
  const original = opts.name ? `<OriginalName v="${opts.name}" />` : "";
  const title = opts.title ? `<Title>${opts.title}</Title>` : "";
  return {
    messagetype: "RichText/Media_CallRecording",
    originalarrivaltime: opts.time ?? "2024-01-02T03:04:00Z",
    content: `<URIObject type="Video.2/CallRecording.1"><RecordingStatus status="Success" code="200" />${title}<a href="${href}">Play</a>${original}</URIObject>`,
  };
}

describe("parseRecordings", () => {
  it("extracts the sharing link, name and host from a CallRecording message", () => {
    const href = "https://contoso-my.sharepoint.com/:v:/p/user/FAKESHARETOKEN";
    const result = parseRecordings([
      recordingMsg(href, { name: "Team Sync-20260526-Meeting Recording.mp4" }),
    ]);
    expect(result).toHaveLength(1);
    expect(result[0]).toEqual({
      name: "Team Sync-20260526-Meeting Recording.mp4",
      date: "2024-01-02 03:04",
      fileUrl: href,
      host: "contoso-my.sharepoint.com",
    });
  });

  it("falls back to the meeting title when there is no OriginalName", () => {
    const href = "https://contoso-my.sharepoint.com/:v:/p/user/ABC";
    const result = parseRecordings([
      recordingMsg(href, { title: "Weekly Standup" }),
    ]);
    expect(result[0].name).toBe("Weekly Standup");
  });

  it("skips in-progress recordings with an empty href", () => {
    const inProgress = {
      messagetype: "RichText/Media_CallRecording",
      content:
        '<URIObject type="Video.2/CallRecording.1"><RecordingStatus status="Initial" code="0" /><a href="">Play</a></URIObject>',
    };
    expect(parseRecordings([inProgress])).toHaveLength(0);
  });

  it("ignores non-recording messages, including doc-share links", () => {
    const messages = [
      {
        messagetype: "RichText/Html",
        content:
          '<a href="https://contoso.sharepoint.com/sites/x/report.docx">doc</a>',
      },
      { messagetype: "ThreadActivity/AddMember", content: "<addmember/>" },
    ];
    expect(parseRecordings(messages)).toHaveLength(0);
  });

  it("de-duplicates repeated recording messages for the same file", () => {
    const href = "https://contoso-my.sharepoint.com/:v:/p/user/SAME";
    expect(
      parseRecordings([recordingMsg(href), recordingMsg(href)]),
    ).toHaveLength(1);
  });
});

const TRANSCRIPT_JSON = JSON.stringify({
  entries: [
    {
      id: "1",
      startOffset: "00:00:01.000",
      endOffset: "00:00:03.500",
      speakerDisplayName: "Alice",
      text: "Hello there",
    },
    {
      id: "2",
      startOffset: "00:00:03.500",
      endOffset: "00:00:05.000",
      speakerDisplayName: "Alice",
      text: "how are you",
    },
    {
      id: "3",
      startOffset: "00:00:05.000",
      endOffset: "00:00:07.000",
      speakerDisplayName: "Bob",
      text: "Good thanks",
    },
  ],
});

describe("timeToSeconds / secondsToVtt", () => {
  it("parses HH:MM:SS.fff to seconds", () => {
    expect(timeToSeconds("01:02:03.5")).toBe(3723.5);
  });
  it("formats seconds back to a VTT cue timestamp", () => {
    expect(secondsToVtt(3723.5)).toBe("01:02:03.500");
  });
});

describe("convertTranscriptToVtt", () => {
  it("produces a WEBVTT document with voice tags", () => {
    const vtt = convertTranscriptToVtt(TRANSCRIPT_JSON);
    expect(vtt.startsWith("WEBVTT\n\n")).toBe(true);
    expect(vtt).toContain("00:00:01.000 --> 00:00:03.500");
    expect(vtt).toContain("<v Alice>Hello there");
    expect(vtt).toContain("<v Bob>Good thanks");
  });
  it("handles an empty transcript", () => {
    expect(convertTranscriptToVtt('{"entries":[]}')).toBe("WEBVTT\n\n");
  });
});

describe("convertTranscriptToGrouped", () => {
  it("merges consecutive lines from the same speaker", () => {
    const grouped = convertTranscriptToGrouped(TRANSCRIPT_JSON);
    expect(grouped).toBe("Alice: Hello there how are you\n\nBob: Good thanks");
  });
});

describe("transcriptFilename", () => {
  it("builds a safe vtt filename, dropping the source extension", () => {
    expect(transcriptFilename("Team Sync.mp4", "vtt")).toBe(
      "Team_Sync_transcript.vtt",
    );
  });
  it("uses a grouped suffix and .txt extension", () => {
    expect(transcriptFilename("clip", "grouped")).toBe(
      "clip_transcript_grouped.txt",
    );
  });
  it("uses .json for the json format", () => {
    expect(transcriptFilename("clip", "json")).toBe("clip_transcript.json");
  });
  it("falls back when the name is empty", () => {
    expect(transcriptFilename("", "vtt")).toBe("recording_transcript.vtt");
  });
});
