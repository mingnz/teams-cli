import type { Page } from "playwright";
import { BROWSER_PROFILE_DIR, TEAMS_URL } from "./config.js";
import { stripHtml } from "./formatting.js";

/**
 * Reaction names accepted on the CLI, mapped to the hover quick-bar buttons.
 * The expanded "More reactions" picker is deliberately NOT wired: its flyout
 * can anchor to an adjacent message row, which risks reacting to the wrong
 * message.
 */
export const REACTIONS = {
  like: "like",
  heart: "heart",
  laugh: "laugh",
  surprised: "surprised",
} as const;

export type ReactionName = keyof typeof REACTIONS;

export function isReactionName(value: string): value is ReactionName {
  return Object.hasOwn(REACTIONS, value);
}

/** Selectors for the hover quick bar. */
export function reactionSelectors(name: ReactionName): {
  quick: string;
} {
  const key = REACTIONS[name];
  return {
    quick: `[data-tid="message-actions-${key}"]`,
  };
}

/**
 * Pick the message to react to from a newest-first message list: the one with
 * the given id, or the newest renderable message. Skips non-message rows
 * (read receipts, reaction events, topic updates) which have no `content`.
 */
export function pickTargetMessage(
  messages: Record<string, unknown>[],
  messageId?: string,
): Record<string, unknown> | null {
  if (messageId) {
    return (
      messages.find(
        (m) => m.id === messageId && typeof m.content === "string",
      ) ?? null
    );
  }
  for (const m of messages) {
    if (m.type !== "Message") continue;
    if (typeof m.content !== "string" || m.content.length === 0) continue;
    return m;
  }
  return null;
}

/** A short distinctive plain-text snippet of the message, for search + row location. */
export function searchSnippet(
  message: Record<string, unknown>,
  maxLen = 48,
): string {
  const text = stripHtml((message.content as string) ?? "")
    // Teams content carries HTML entities the DOM renders decoded; match those.
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) throw new Error("Target message has no searchable text content.");
  return text.length > maxLen ? text.slice(0, maxLen) : text;
}

export interface EmotionUser {
  mri?: string;
}

export interface Emotion {
  key?: string;
  users?: EmotionUser[];
}

/** Read the emotions property off a message, tolerating stringified JSON. */
export function getEmotions(message: Record<string, unknown>): Emotion[] {
  const props = (message.properties as Record<string, unknown>) ?? {};
  const raw = props.emotions;
  if (typeof raw === "string") {
    try {
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? (parsed as Emotion[]) : [];
    } catch {
      return [];
    }
  }
  return Array.isArray(raw) ? (raw as Emotion[]) : [];
}

/** True when `myMri` currently has the given reaction on the message. */
export function hasMyReaction(
  message: Record<string, unknown>,
  name: ReactionName,
  myMri: string,
): boolean {
  const key = REACTIONS[name];
  return getEmotions(message).some(
    (e) => e.key === key && (e.users ?? []).some((u) => u.mri === myMri),
  );
}

/** Escape a literal for embedding into a RegExp. */
export function escapeRegExp(literal: string): string {
  return literal.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/* v8 ignore start -- requires a live Playwright browser */

async function waitForApp(page: Page, timeoutMs: number): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if ((await page.locator('[role="combobox"]').count()) > 0) return;
    await page.waitForTimeout(1000);
  }
  throw new Error(
    "Teams app did not finish loading. Try `teams react --headed` to watch it.",
  );
}

/**
 * Drive the Teams web client to add (or toggle off) a reaction on a message.
 *
 * The chatsvc REST API rejects reaction writes outright, and the native
 * protocol rides a WebSocket inside a SharedWorker, so the only scriptable
 * path is the web UI itself: open the chat from the chat list (by its
 * display name), locate the message row by its text, hover it, and click
 * the reaction button on the hover quick bar (falling back to the expanded
 * picker).
 */
export async function driveReaction(opts: {
  name: ReactionName;
  snippet: string;
  chatName: string;
  headed?: boolean;
  timeoutMs: number;
}): Promise<void> {
  const playwright = await import("playwright");
  const context = await playwright.chromium.launchPersistentContext(
    BROWSER_PROFILE_DIR,
    { headless: !opts.headed },
  );
  try {
    const page = await context.newPage();
    await page.goto(TEAMS_URL, { waitUntil: "domcontentloaded" });
    await waitForApp(page, opts.timeoutMs);

    // Open the chat natively from the chat list. Match the row's accessible
    // NAME (the chat title), never subtree text - previews inside other rows
    // can mention the same person and open the wrong chat. (Opening a message
    // through the search page renders the thread in an embedded pane that can
    // fail with an "Oops" error - the native chat view is the reliable path.)
    const rowPattern = new RegExp(escapeRegExp(opts.snippet.slice(0, 20)));
    const namePattern = new RegExp(
      `^(?:Unread message )?(?:Chat )?${escapeRegExp(opts.chatName)}`,
    );
    const candidates = page
      .getByRole("treeitem", { name: namePattern })
      .or(
        page
          .locator('[role="treeitem"], [data-tid="chat-list-item"]')
          .filter({ hasText: opts.chatName }),
      );
    // The chat list renders after app boot; give it a beat before counting.
    await candidates
      .first()
      .waitFor({ state: "visible", timeout: 15_000 })
      .then(
        () => true,
        () => false,
      );
    const total = await candidates.count();
    if (total === 0) {
      throw new Error(
        `Chat '${opts.chatName}' not visible in the chat list. ` +
          "Open it once in the app so it becomes recent, then retry.",
      );
    }
    // The same 1:1 can appear as several rows; try each until the target
    // message heading shows up.
    const chatRow = page.getByRole("heading", { name: rowPattern }).last();
    let found = false;
    for (let i = 0; i < total; i++) {
      await candidates.nth(i).click({ timeout: 10_000 });
      found = await chatRow.waitFor({ state: "visible", timeout: 12_000 }).then(
        () => true,
        () => false,
      );
      if (found) break;
    }
    if (!found) {
      throw new Error(
        `Message "${opts.snippet.slice(0, 24)}..." not found in the opened ` +
          `chat (tried ${total} chat-list row(s)). The message may be older ` +
          "than the rendered window.",
      );
    }

    // Hover the message row to reveal the reaction quick bar. The role=heading
    // match can be a 1x1 accessibility marker - hover the visible GROUP row
    // (or its DOM rect) instead; then a raw mouse move to its center. The
    // hover can silently miss (render timing), so retry re-resolving the row.
    const { quick } = reactionSelectors(opts.name);
    // Bar-up signal: ANY hover action button - the quick reaction buttons
    // and the "More reactions" entry render either/or depending on space.
    const barUp = page
      .locator(
        `${quick}, [data-tid^="message-actions-"], ` +
          '[data-tid="add-reaction-picker-entry-point-button"]',
      )
      .first();
    const groupRow = page.getByRole("group", { name: rowPattern }).last();
    const prefix = opts.snippet.slice(0, 20);
    let hovered = false;
    for (let attempt = 0; attempt < 3 && !hovered; attempt++) {
      const box =
        (await groupRow.boundingBox().catch(() => null)) ??
        (await domRowBox(page, prefix));
      if (box && box.width > 5) {
        await chatRow.scrollIntoViewIfNeeded().catch(() => {});
        // The reaction bar hangs on the row's TOP edge: hovering the center
        // often does not trigger it. Sweep a few offsets until the bar shows.
        const cx = box.x + box.width / 2;
        const ys = [box.y - 4, box.y + 4, box.y + box.height / 2, box.y - 12];
        for (const y of ys) {
          await page.mouse.move(cx, Math.max(5, y));
          hovered = await barUp
            .waitFor({ state: "visible", timeout: 1500 })
            .then(
              () => true,
              () => false,
            );
          if (hovered) break;
        }
      } else {
        hovered = await barUp.waitFor({ state: "visible", timeout: 5000 }).then(
          () => true,
          () => false,
        );
      }
    }
    if (!hovered) {
      throw new Error(
        "The reaction hover bar did not appear on the message row. " +
          "Retry, or run with --headed to watch the drive.",
      );
    }
    await page.locator(quick).first().click({ timeout: 5000 });
  } finally {
    await context.close().catch(() => {});
  }
}

// Find the on-screen rect of the message row containing `prefix`: walk up
// from the deepest visible text node carrying the snippet to its row
// container. Returns null when the text is not rendered.
async function domRowBox(
  page: Page,
  prefix: string,
): Promise<{ x: number; y: number; width: number; height: number } | null> {
  return page.evaluate((p) => {
    const leaves = [...document.querySelectorAll("div,span,p")].filter(
      (el) =>
        el.children.length === 0 &&
        (el.textContent ?? "").includes(p) &&
        el.getBoundingClientRect().width > 20,
    );
    if (leaves.length === 0) return null;
    const leaf = leaves[leaves.length - 1];
    let node = leaf;
    for (let i = 0; i < 8 && node.parentElement; i++) {
      const r = node.getBoundingClientRect();
      if (r.height > 30 && r.width > 200) break;
      node = node.parentElement;
    }
    const r = node.getBoundingClientRect();
    return { x: r.x, y: r.y, width: r.width, height: r.height };
  }, prefix);
}

/* v8 ignore stop */
