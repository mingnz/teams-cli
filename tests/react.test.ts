import { describe, expect, it } from "vitest";
import {
  getEmotions,
  hasMyReaction,
  isReactionName,
  pickTargetMessage,
  REACTIONS,
  reactionSelectors,
  searchSnippet,
} from "../src/react.js";

const ME = "8:orgid:11111111-1111-1111-1111-111111111111";

function msg(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    type: "Message",
    id: "100",
    content: "<p>hello world</p>",
    ...overrides,
  };
}

describe("reaction names", () => {
  it("accepts the quick-bar reactions", () => {
    for (const name of ["like", "heart", "laugh", "surprised"]) {
      expect(isReactionName(name)).toBe(true);
    }
  });

  it("rejects unknown names", () => {
    expect(isReactionName("party")).toBe(false);
    expect(isReactionName("sad")).toBe(false);
    expect(isReactionName("Like")).toBe(false);
    expect(isReactionName("")).toBe(false);
  });

  it("maps selectors to data-tid anchors", () => {
    expect(reactionSelectors("like").quick).toBe(
      '[data-tid="message-actions-like"]',
    );
    expect(Object.keys(REACTIONS)).toHaveLength(4);
  });
});

describe("pickTargetMessage", () => {
  const list = [
    msg({ id: "3", content: "<p>newest</p>" }),
    { type: "Message", id: "2", content: "" }, // event row, no content
    { type: "Reaction", id: "r1", content: "<p>x</p>" },
    msg({ id: "1", content: "<p>oldest</p>" }),
  ];

  it("picks the newest message with content by default", () => {
    expect(pickTargetMessage(list)?.id).toBe("3");
  });

  it("skips reaction/event rows when scanning", () => {
    const only = [{ type: "Reaction", id: "r1", content: "<p>x</p>" }];
    expect(pickTargetMessage(only)).toBeNull();
  });

  it("finds by explicit id", () => {
    expect(pickTargetMessage(list, "1")?.id).toBe("1");
  });

  it("returns null for an unknown id", () => {
    expect(pickTargetMessage(list, "999")).toBeNull();
  });
});

describe("searchSnippet", () => {
  it("strips html and collapses whitespace", () => {
    expect(searchSnippet(msg({ content: "<p>a&nbsp;&nbsp;b</p>" }))).toBe(
      "a b",
    );
  });

  it("truncates to the max length", () => {
    const long = "x".repeat(100);
    expect(searchSnippet(msg({ content: `<p>${long}</p>` }))).toHaveLength(48);
  });

  it("throws on empty content", () => {
    expect(() => searchSnippet(msg({ content: "<p></p>" }))).toThrow();
  });
});

describe("emotions state", () => {
  it("reads a plain array property", () => {
    const m = msg({
      properties: {
        emotions: [{ key: "like", users: [{ mri: ME }] }],
      },
    });
    expect(hasMyReaction(m, "like", ME)).toBe(true);
    expect(hasMyReaction(m, "heart", ME)).toBe(false);
  });

  it("reads a stringified emotions property", () => {
    const m = msg({
      properties: {
        emotions: JSON.stringify([{ key: "like", users: [{ mri: ME }] }]),
      },
    });
    expect(hasMyReaction(m, "like", ME)).toBe(true);
  });

  it("tolerates missing, empty, and malformed properties", () => {
    expect(getEmotions(msg())).toEqual([]);
    expect(getEmotions(msg({ properties: { emotions: [] } }))).toEqual([]);
    expect(getEmotions(msg({ properties: { emotions: "not-json" } }))).toEqual(
      [],
    );
    expect(hasMyReaction(msg(), "like", ME)).toBe(false);
  });

  it("detects other users only", () => {
    const m = msg({
      properties: {
        emotions: [{ key: "like", users: [{ mri: "8:orgid:other" }] }],
      },
    });
    expect(hasMyReaction(m, "like", ME)).toBe(false);
  });

  it("treats an empty users list as no reaction", () => {
    const m = msg({
      properties: { emotions: [{ key: "like", users: [] }] },
    });
    expect(hasMyReaction(m, "like", ME)).toBe(false);
  });
});
