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
  let body = convId.includes(":")
    ? convId.split(":").slice(1).join(":")
    : convId;
  body = body.includes("@") ? body.split("@")[0] : body;
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

export function formatMessage(
  msg: Record<string, unknown>,
): FormattedMessage | null {
  const msgType = (msg.messagetype as string) ?? "";
  if (!["RichText/Html", "Text", "RichText"].includes(msgType)) return null;

  const sender = (msg.imdisplayname as string) ?? "Unknown";
  const content = stripHtml((msg.content as string) ?? "");
  if (!content) return null;

  return {
    time: formatTimestamp(msg.originalarrivaltime as string),
    sender,
    body: content,
    id: (msg.id as string) ?? "",
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
