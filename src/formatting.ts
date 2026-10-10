export function stripHtml(text: string): string {
  let result = text;
  let prev = "";
  // Repeatedly strip tags until no more remain (handles nested/partial tags)
  while (result !== prev) {
    prev = result;
    result = result.replace(/<[^>]+>/g, "");
  }
  return result.trim();
}

export function formatTimestamp(ts: string | null | undefined): string {
  if (!ts) return "";
  try {
    const dt = new Date(ts);
    if (Number.isNaN(dt.getTime())) throw new Error();
    const y = dt.getUTCFullYear();
    const mo = String(dt.getUTCMonth() + 1).padStart(2, "0");
    const d = String(dt.getUTCDate()).padStart(2, "0");
    const h = String(dt.getUTCHours()).padStart(2, "0");
    const mi = String(dt.getUTCMinutes()).padStart(2, "0");
    return `${y}-${mo}-${d} ${h}:${mi}`;
  } catch {
    return ts.length >= 16 ? ts.slice(0, 16) : ts;
  }
}

export function getConversationDisplayName(
  conv: Record<string, unknown>,
): string {
  const threadProps = (conv.threadProperties as Record<string, string>) ?? {};
  const topic = threadProps.topic ?? "";
  if (topic.trim()) return topic.trim();

  const lastMsg = (conv.lastMessage as Record<string, string>) ?? {};
  const sender = lastMsg.imdisplayname ?? "";
  if (sender) return `Chat with ${sender}`;

  const convId = (conv.id as string) ?? "";
  return convId.slice(0, 50);
}

export function getConversationType(conv: Record<string, unknown>): string {
  const threadProps = (conv.threadProperties as Record<string, string>) ?? {};
  const productType = threadProps.productThreadType ?? "";
  if (productType) return productType;

  const typeMap: Record<string, string> = {
    meeting: "Meeting",
    chat: "Chat",
    space: "Channel",
    streamofnotifications: "System",
    streamofmentions: "System",
  };
  return typeMap[threadProps.threadType ?? ""] ?? "Chat";
}

export function makeShortId(convId: string): string {
  // Strip prefix before first ":" (e.g. "19:" or "8:orgid:")
  let body = convId;
  const colonIdx = body.indexOf(":");
  if (colonIdx !== -1) body = body.slice(colonIdx + 1);
  // Strip domain suffix after "@"
  const atIdx = body.indexOf("@");
  if (atIdx !== -1) body = body.slice(0, atIdx);
  return body;
}

function assignShortIds(entries: Record<string, unknown>[], minLen = 4): void {
  const hashes = entries.map((e) => makeShortId(e.id as string));
  const maxLen = Math.max(...hashes.map((h) => h.length), minLen);
  let length = minLen;
  while (length <= maxLen) {
    const shortIds = hashes.map((h) => h.slice(-length));
    if (new Set(shortIds).size === shortIds.length) break;
    length++;
  }
  for (let i = 0; i < entries.length; i++) {
    entries[i].short_id = hashes[i].slice(-length);
  }
}

export interface FormattedChat {
  short_id: string;
  name: string;
  type: string;
  preview: string;
  time: string;
  id: string;
}

export function formatChatList(
  conversations: Record<string, unknown>[],
): FormattedChat[] {
  const results: Record<string, unknown>[] = [];

  for (const conv of conversations) {
    const convId = (conv.id as string) ?? "";
    if (convId.startsWith("48:")) continue;

    const lastMsg = (conv.lastMessage as Record<string, unknown>) ?? {};
    const lastBody = stripHtml((lastMsg.content as string) ?? "");
    const lastSender = (lastMsg.imdisplayname as string) ?? "";
    const lastTime = formatTimestamp(lastMsg.originalarrivaltime as string);

    let preview = "";
    if (lastSender && lastBody) {
      preview = `${lastSender}: ${lastBody.slice(0, 60)}`;
    } else if (lastBody) {
      preview = lastBody.slice(0, 80);
    }

    results.push({
      short_id: "",
      name: getConversationDisplayName(conv),
      type: getConversationType(conv),
      preview,
      time: lastTime,
      id: convId,
    });
  }

  assignShortIds(results);
  return results as unknown as FormattedChat[];
}

export interface FormattedMessage {
  time: string;
  sender: string;
  body: string;
  id: string;
}

export interface MediaInfo {
  url: string;
  duration: string;
  kind: string;
}

// Voice notes and other card-based media: the playable URL and duration live
// in the message's `properties.cards` JSON payload.
export function extractMedia(msg: Record<string, unknown>): MediaInfo | null {
  try {
    const props = (msg.properties as Record<string, string>) ?? {};
    const cards = props.cards ? JSON.parse(props.cards) : [];
    for (const c of cards as Record<string, unknown>[]) {
      const content = c?.content as Record<string, unknown> | undefined;
      const media = content?.media as Record<string, string>[] | undefined;
      if (Array.isArray(media) && media[0]?.url) {
        const dm = String(content?.duration ?? "").match(/PT(?:(\d+)M)?(\d+)S/);
        const dur = dm ? `${dm[1] ? `${dm[1]}m` : ""}${dm[2] ?? "?"}s` : "?";
        const kind = String(content?.contentType ?? "").includes("audio")
          ? "voice note"
          : "media";
        return { url: media[0].url, duration: dur, kind };
      }
    }
  } catch {
    // malformed cards JSON - not a media message
  }
  return null;
}

// Inline images: the downloadable object URL sits in the raw HTML `<img src>`
// of the message content (e.g. .../v1/objects/<ref>/views/imgo).
export function extractImage(msg: Record<string, unknown>): MediaInfo | null {
  try {
    const html = (msg.content as string) ?? "";
    const m = html.match(/<img[^>]+src="(https:\/\/[^"]+\/objects\/[^"]+)"/i);
    if (!m) return null;
    const url = m[1].replace(/&amp;/g, "&");
    const dim = html.match(/<img[^>]+width="(\d+)"/i);
    return { url, duration: dim ? `${dim[1]}w` : "?", kind: "image" };
  } catch {
    // fall through
  }
  return null;
}

export interface FileInfo {
  fileName: string;
  objectUrl: string;
  url: string;
}

// File attachments: `properties.files` carries the name plus a pre-authed
// tempauth download URL (botFileProperties.url) with the SharePoint objectUrl
// as fallback.
export function extractFile(msg: Record<string, unknown>): FileInfo | null {
  try {
    const props = (msg.properties as Record<string, string>) ?? {};
    const files = props.files ? JSON.parse(props.files) : [];
    for (const f of files as Record<string, unknown>[]) {
      if (f?.fileName) {
        const bot = f?.botFileProperties as Record<string, string> | undefined;
        const objectUrl = (f.objectUrl as string) ?? "";
        return {
          fileName: f.fileName as string,
          objectUrl,
          url: bot?.url ?? objectUrl,
        };
      }
    }
  } catch {
    // malformed files JSON - not a file message
  }
  return null;
}

export function formatMessage(
  msg: Record<string, unknown>,
): FormattedMessage | null {
  const msgType = (msg.messagetype as string) ?? "";
  if (!["RichText/Html", "Text", "RichText"].includes(msgType)) return null;

  const sender = (msg.imdisplayname as string) ?? "Unknown";
  const time = formatTimestamp(msg.originalarrivaltime as string);
  const id = (msg.id as string) ?? "";
  const content = stripHtml((msg.content as string) ?? "");
  if (!content) {
    // Media/file/image-only messages carry no text - render a marker row
    // instead of dropping the message entirely.
    const media = extractMedia(msg);
    if (media) {
      return {
        time,
        sender,
        body: `[${media.kind} ${media.duration} — media id ${id}]`,
        id,
      };
    }
    const file = extractFile(msg);
    if (file) {
      return {
        time,
        sender,
        body: `[file: ${file.fileName} — file id ${id}]`,
        id,
      };
    }
    const image = extractImage(msg);
    if (image) {
      return {
        time,
        sender,
        body: `[${image.kind} ${image.duration} — img id ${id}]`,
        id,
      };
    }
    return null;
  }

  const img = extractImage(msg);
  return {
    time,
    sender,
    body: img
      ? `${content} [${img.kind} ${img.duration} — img id ${id}]`
      : content,
    id,
  };
}

export interface FormattedPerson {
  name: string;
  email: string;
  mri: string;
  title: string;
  department: string;
  company: string;
  type: string;
}

export function formatPerson(person: Record<string, unknown>): FormattedPerson {
  const emails = (person.EmailAddresses as string[]) ?? [];
  return {
    name: (person.DisplayName as string) ?? "",
    email: emails[0] ?? "",
    mri: (person.MRI as string) ?? "",
    title: (person.JobTitle as string) ?? "",
    department: (person.Department as string) ?? "",
    company: (person.CompanyName as string) ?? "",
    type: (person.PeopleSubtype as string) ?? "",
  };
}

export interface RecordingRef {
  name: string;
  date: string;
  fileUrl: string;
  host: string;
}

function matchAttr(xml: string, re: RegExp): string {
  return (xml.match(re)?.[1] ?? "").replace(/&amp;/g, "&").trim();
}

// Derive a human name for a recording from its URIObject, preferring the
// original filename, then the meeting title, then the file in the URL.
function recordingName(xml: string, fileUrl: string): string {
  const original = matchAttr(xml, /<OriginalName\s+v="([^"]*)"/i);
  if (original) return original;
  const title = matchAttr(xml, /<Title>([^<]*)<\/Title>/i);
  if (title) return title;
  try {
    const last = decodeURIComponent(
      new URL(fileUrl).pathname.split("/").pop() ?? "",
    );
    return last || "Recording";
  } catch {
    return "Recording";
  }
}

// Extract meeting recordings from a list of chat messages. Recordings arrive as
// `RichText/Media_CallRecording` messages whose content is a `<URIObject>` with
// a SharePoint sharing link. The recording goes through several states; only the
// finished one carries a usable link, so empty hrefs are skipped. The sharing
// URL is kept verbatim — it is what the shares API resolves to a drive item.
// De-duplicated by URL.
export function parseRecordings(
  messages: Record<string, unknown>[],
): RecordingRef[] {
  const seen = new Set<string>();
  const results: RecordingRef[] = [];

  for (const msg of messages) {
    if (msg.messagetype !== "RichText/Media_CallRecording") continue;
    const xml = (msg.content as string) ?? "";

    const fileUrl = matchAttr(xml, /<a\s+href="([^"]*)"/i);
    if (!fileUrl || !/sharepoint\.com/i.test(fileUrl)) continue;
    if (seen.has(fileUrl)) continue;
    seen.add(fileUrl);

    let host = "";
    try {
      host = new URL(fileUrl).host;
    } catch {
      continue;
    }

    results.push({
      name: recordingName(xml, fileUrl),
      date: formatTimestamp(msg.originalarrivaltime as string),
      fileUrl,
      host,
    });
  }

  return results;
}

interface TranscriptEntry {
  id?: string | number;
  startOffset?: string;
  endOffset?: string;
  speakerDisplayName?: string;
  text?: string;
}

function parseTranscript(jsonText: string): TranscriptEntry[] {
  const data = JSON.parse(jsonText) as { entries?: TranscriptEntry[] };
  return data.entries ?? [];
}

// "HH:MM:SS(.fff)" -> seconds (rounded to ms).
export function timeToSeconds(t: string): number {
  const [h, m, s] = t.split(":");
  return (
    Math.round(
      (Number(h) * 3600 + Number(m) * 60 + Number.parseFloat(s)) * 1000,
    ) / 1000
  );
}

// seconds -> "HH:MM:SS.fff" (WebVTT cue timestamp).
export function secondsToVtt(seconds: number): string {
  const h = String(Math.floor(seconds / 3600)).padStart(2, "0");
  const m = String(Math.floor((seconds % 3600) / 60)).padStart(2, "0");
  const s = (seconds % 60).toFixed(3).padStart(6, "0");
  return `${h}:${m}:${s}`;
}

// Convert the MS Stream transcript JSON to WebVTT with <v Speaker> voice tags.
export function convertTranscriptToVtt(jsonText: string): string {
  const entries = parseTranscript(jsonText);
  let vtt = "WEBVTT\n\n";
  entries.forEach((entry, index) => {
    const start = secondsToVtt(timeToSeconds(entry.startOffset ?? "00:00:00"));
    const end = secondsToVtt(timeToSeconds(entry.endOffset ?? "00:00:00"));
    const speaker = entry.speakerDisplayName || "Unknown";
    const text = entry.text ?? "";
    vtt += `${entry.id ?? index + 1}\n`;
    vtt += `${start} --> ${end}\n`;
    vtt += `<v ${speaker}>${text}\n\n`;
  });
  return vtt;
}

// Convert the transcript JSON to readable text, merging consecutive lines from
// the same speaker into one paragraph.
export function convertTranscriptToGrouped(jsonText: string): string {
  const entries = parseTranscript(jsonText);
  const grouped: string[] = [];
  let currentSpeaker: string | null = null;
  let buffer = "";

  for (const entry of entries) {
    const speaker = entry.speakerDisplayName || "Unknown";
    const text = entry.text ?? "";
    if (speaker !== currentSpeaker) {
      if (buffer) grouped.push(`${currentSpeaker}: ${buffer.trim()}`);
      currentSpeaker = speaker;
      buffer = text;
    } else {
      buffer += ` ${text}`;
    }
  }
  if (buffer && currentSpeaker)
    grouped.push(`${currentSpeaker}: ${buffer.trim()}`);

  return grouped.join("\n\n");
}

const TRANSCRIPT_EXTENSIONS: Record<string, string> = {
  json: ".json",
  vtt: ".vtt",
  grouped: ".txt",
};

// Build a safe default output filename for a downloaded transcript.
export function transcriptFilename(name: string, format: string): string {
  const base = (name || "recording")
    .replace(/\.[^.]+$/, "") // drop extension (e.g. .mp4)
    .replace(/[^\w.-]+/g, "_")
    .replace(/^_+|_+$/g, "");
  const suffix = format === "grouped" ? "_transcript_grouped" : "_transcript";
  const ext = TRANSCRIPT_EXTENSIONS[format] ?? ".txt";
  return `${base || "recording"}${suffix}${ext}`;
}

export interface FormattedMember {
  name: string;
  role: string;
  type: string;
  mri: string;
}

export function formatMember(member: Record<string, unknown>): FormattedMember {
  const mri = (member.id as string) ?? "";
  const name = (member.friendlyName as string) ?? "";
  const role = (member.role as string) ?? "User";

  let memberType = "Other";
  if (mri.startsWith("28:")) memberType = "Bot";
  else if (mri.startsWith("8:orgid:")) memberType = "User";

  return { name: name || mri, role, type: memberType, mri };
}
