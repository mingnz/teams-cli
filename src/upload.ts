import { basename } from "node:path";
import { buildMessageBody, type MessageBody, postMessage } from "./api.js";
import type { ApiClient } from "./client.js";

// Sending a file = upload it to the sender's OneDrive (the folder Teams itself
// uses), mint a share link, then post a chat message whose `properties.files`
// carries a file object the Teams client renders as a file card. The file
// object shape below is field-proven against the real client — keep it exact.

export const CHAT_FILES_FOLDER = "Microsoft Teams Chat Files";

// OneDrive simple upload (single PUT) is capped at 250 MB; larger files need an
// upload session, which this command does not implement.
export const MAX_SIMPLE_UPLOAD_BYTES = 250 * 1024 * 1024;

export const POST_UPLOAD = "<post-upload>";

export type LinkScope = "organization" | "anonymous";

const LINK_SCOPE_ALIASES: Record<string, LinkScope> = {
  organization: "organization",
  org: "organization",
  anonymous: "anonymous",
  anon: "anonymous",
};

export function parseLinkScope(value: string): LinkScope | null {
  return LINK_SCOPE_ALIASES[value.toLowerCase()] ?? null;
}

// The file object's `permissionScope`. "Anyone" is what the client sent for an
// anonymous link in the proven payload.
export const PERMISSION_SCOPE: Record<LinkScope, string> = {
  anonymous: "Anyone",
  organization: "Organization",
};

// ── Naming / encoding ───────────────────────────────────────────────────────

/** The upload name for a local path: its basename. Throws when empty. */
export function uploadFileName(path: string): string {
  const name = basename(path);
  if (!name || name === "." || name === "..") {
    throw new Error(`Invalid file name: '${path}'.`);
  }
  return name;
}

/** Lowercased extension without the dot ("" when there is none). */
export function fileExtension(name: string): string {
  const dot = name.lastIndexOf(".");
  return dot > 0 ? name.slice(dot + 1).toLowerCase() : "";
}

/** URL-encode each path segment on its own (survives Unicode / Arabic names). */
export function encodePathSegments(...segments: string[]): string {
  return segments.map((s) => encodeURIComponent(s)).join("/");
}

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Message content for an optional caption: escaped into <div><p>…</p></div>. */
export function captionHtml(caption?: string): string {
  return caption ? `<div><p>${escapeHtml(caption)}</p></div>` : "";
}

// ── Site derivation ─────────────────────────────────────────────────────────

export interface OneDriveSite {
  host: string; // "<tenant>-my.sharepoint.com"
  siteUrl: string; // "https://<host>/personal/<slug>"   (API prefix)
  base: string; // siteUrl + "/"                         (file object baseUrl)
}

export function decodeJwtClaims(token: string): Record<string, unknown> {
  try {
    const payload = token.split(".")[1];
    const padded = payload + "=".repeat((4 - (payload.length % 4)) % 4);
    return JSON.parse(Buffer.from(padded, "base64url").toString());
  } catch {
    throw new Error("SharePoint token is not a decodable JWT.");
  }
}

/** Personal-site slug for a UPN: `a.b@tenant.com` → `a_b_tenant_com`. */
export function personalSiteSlug(upn: string): string {
  return upn.toLowerCase().replace(/[^a-z0-9_-]/g, "_");
}

const ONEDRIVE_HOST = /^[a-z0-9-]+-my\.sharepoint\.com$/;

/**
 * Derive the user's OneDrive personal site from the SharePoint token: the
 * `aud` claim gives the host (falling back to `fallbackHost`, the host the
 * token was captured on, when `aud` is the SharePoint app id rather than a
 * URL), the `upn`/`unique_name` claim gives the personal-site slug.
 */
export function deriveOneDriveSite(
  token: string,
  fallbackHost?: string,
): OneDriveSite {
  const claims = decodeJwtClaims(token);
  let host = "";
  const aud = typeof claims.aud === "string" ? claims.aud : "";
  try {
    host = new URL(aud).host.toLowerCase();
  } catch {
    host = (fallbackHost ?? "").toLowerCase();
  }
  if (!ONEDRIVE_HOST.test(host)) {
    throw new Error(
      `SharePoint token audience '${aud || "(none)"}' is not a OneDrive ` +
        "host (<tenant>-my.sharepoint.com).",
    );
  }
  const upn =
    (typeof claims.upn === "string" && claims.upn) ||
    (typeof claims.unique_name === "string" && claims.unique_name) ||
    "";
  if (!upn.includes("@")) {
    throw new Error("SharePoint token carries no upn/unique_name claim.");
  }
  const siteUrl = `https://${host}/personal/${personalSiteSlug(upn)}`;
  return { host, siteUrl, base: `${siteUrl}/` };
}

// ── File object + message body ──────────────────────────────────────────────

export interface FileObjectInput {
  fileName: string; // server-returned (final) name
  base: string;
  listItemUniqueId: string;
  listId: string;
  siteId: string;
  shareUrl: string;
  shareId: string;
  scope: LinkScope;
}

export function objectUrlFor(base: string, fileName: string): string {
  return `${base}Documents/${encodePathSegments(CHAT_FILES_FOLDER, fileName)}`;
}

/** The `properties.files[]` entry, in the exact field order the client sends. */
export function buildFileObject(f: FileObjectInput): Record<string, unknown> {
  const objectUrl = objectUrlFor(f.base, f.fileName);
  const ext = fileExtension(f.fileName);
  return {
    itemid: f.listItemUniqueId,
    fileName: f.fileName,
    fileType: ext,
    fileInfo: {
      itemId: null,
      fileUrl: objectUrl,
      siteUrl: f.base,
      serverRelativeUrl: "",
      shareUrl: f.shareUrl,
      shareId: f.shareId,
    },
    fileChicletState: { serviceName: "p2p", state: "active" },
    "@type": "http://schema.skype.com/File",
    version: 2,
    id: f.listItemUniqueId,
    baseUrl: f.base,
    objectUrl,
    openUrl: null,
    type: ext,
    title: f.fileName,
    state: "active",
    chicletBreadcrumbs: null,
    providerData: "",
    botFileProperties: {},
    isUploadError: null,
    progressComplete: null,
    permissionScope: PERMISSION_SCOPE[f.scope],
    filePreview: { previewUrl: "", previewHeight: 0, previewWidth: 0 },
    sharepointIds: {
      listId: f.listId,
      listItemUniqueId: f.listItemUniqueId,
      siteId: f.siteId,
      siteUrl: null,
      webId: null,
      driveId: null,
    },
    publication: null,
    site: null,
  };
}

/** Chat message carrying file objects; `files` is a JSON string, as stored. */
export function buildFileMessageBody(
  files: Record<string, unknown>[],
  caption?: string,
): MessageBody {
  return buildMessageBody(captionHtml(caption), {
    html: true,
    amsreferences: [],
    properties: {
      mentions: "[]",
      cards: "[]",
      links: "[]",
      files: JSON.stringify(files),
      formatVariant: "TEAMS",
    },
  });
}

// ── OneDrive calls (plain fetch, explicit SharePoint headers) ───────────────
// Never the chat client: its fetch attaches ic3 auth to a different host.

async function spFetch(
  token: string,
  url: string,
  init: RequestInit & { timeoutMs?: number } = {},
): Promise<unknown> {
  const { timeoutMs = 30_000, headers, ...rest } = init;
  const resp = await fetch(url, {
    ...rest,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      ...(headers as Record<string, string>),
    },
    signal: AbortSignal.timeout(timeoutMs),
  });
  if (!resp.ok) {
    const body = await resp.text().catch(() => "");
    throw new Error(`HTTP ${resp.status}: ${body.slice(0, 200)}`);
  }
  return resp.json();
}

export async function uploadToChatFiles(
  site: OneDriveSite,
  token: string,
  fileName: string,
  data: Uint8Array<ArrayBuffer>,
): Promise<{ id: string; name: string }> {
  const path = encodePathSegments(CHAT_FILES_FOLDER, fileName);
  const url =
    `${site.siteUrl}/_api/v2.1/drive/root:/${path}:/content` +
    "?@name.conflictBehavior=rename";
  let res: Record<string, unknown>;
  try {
    res = (await spFetch(token, url, {
      method: "PUT",
      headers: { "Content-Type": "application/octet-stream" },
      body: data,
      timeoutMs: 300_000,
    })) as Record<string, unknown>;
  } catch (e) {
    // First drive call: a wrong derived site shows up here.
    throw new Error(
      `${(e as Error).message} (derived OneDrive site: ${site.siteUrl})`,
    );
  }
  const id = typeof res.id === "string" ? res.id : "";
  if (!id) throw new Error("Upload response carried no item id.");
  // With conflictBehavior=rename the server may suffix the name — use its name.
  const name = typeof res.name === "string" && res.name ? res.name : fileName;
  return { id, name };
}

export interface SharepointIds {
  listItemUniqueId: string;
  siteId: string;
  listId: string;
}

export async function getItemSharepointIds(
  site: OneDriveSite,
  token: string,
  itemId: string,
): Promise<SharepointIds> {
  const res = (await spFetch(
    token,
    `${site.siteUrl}/_api/v2.1/drive/items/${itemId}?select=id,sharepointIds`,
  )) as Record<string, unknown>;
  const ids = (res.sharepointIds as Record<string, unknown>) ?? {};
  const out = {
    listItemUniqueId: String(ids.listItemUniqueId ?? ""),
    siteId: String(ids.siteId ?? ""),
    listId: String(ids.listId ?? ""),
  };
  if (!out.listItemUniqueId || !out.siteId || !out.listId) {
    throw new Error("Item response is missing sharepointIds.");
  }
  return out;
}

export async function createShareLink(
  site: OneDriveSite,
  token: string,
  itemId: string,
  scope: LinkScope,
): Promise<{ shareUrl: string; shareId: string }> {
  const res = (await spFetch(
    token,
    `${site.siteUrl}/_api/v2.1/drive/items/${itemId}/microsoft.graph.createLink`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ type: "view", scope }),
    },
  )) as Record<string, unknown>;
  const link = (res.link as Record<string, unknown>) ?? {};
  const shareUrl = typeof link.webUrl === "string" ? link.webUrl : "";
  const shareId = typeof res.id === "string" ? res.id : "";
  if (!shareUrl || !shareId) {
    throw new Error("createLink response is missing link.webUrl or id.");
  }
  return { shareUrl, shareId };
}

// ── Staged send ─────────────────────────────────────────────────────────────

export type Stage = "upload" | "ids" | "link" | "message";
export const STAGES: Stage[] = ["upload", "ids", "link", "message"];

export interface StageReceipt {
  stage: Stage;
  receipt: string;
}

/** A stage failed; `completed` lists the stages that already succeeded. */
export class StageError extends Error {
  constructor(
    readonly stage: Stage,
    readonly completed: StageReceipt[],
    cause: Error,
  ) {
    super(cause.message);
  }
}

export interface SendFileResult {
  fileName: string;
  itemId: string;
  shareUrl: string;
  arrival: string;
  receipts: StageReceipt[];
}

export async function sendFile(opts: {
  chat: ApiClient;
  conversationId: string;
  site: OneDriveSite;
  token: string;
  fileName: string;
  data: Uint8Array<ArrayBuffer>;
  scope: LinkScope;
  caption?: string;
  onStage?: (r: StageReceipt) => void;
}): Promise<SendFileResult> {
  const receipts: StageReceipt[] = [];
  async function stage<T>(
    name: Stage,
    run: () => Promise<T>,
    receipt: (v: T) => string,
  ): Promise<T> {
    let value: T;
    try {
      value = await run();
    } catch (e) {
      throw new StageError(name, [...receipts], e as Error);
    }
    const r = { stage: name, receipt: receipt(value) };
    receipts.push(r);
    opts.onStage?.(r);
    return value;
  }

  const { site, token } = opts;
  const uploaded = await stage(
    "upload",
    () => uploadToChatFiles(site, token, opts.fileName, opts.data),
    (u) => `${u.name} (item ${u.id}) at ${objectUrlFor(site.base, u.name)}`,
  );
  const ids = await stage(
    "ids",
    () => getItemSharepointIds(site, token, uploaded.id),
    (i) => `listItemUniqueId ${i.listItemUniqueId}`,
  );
  const link = await stage(
    "link",
    () => createShareLink(site, token, uploaded.id, opts.scope),
    (l) => `${l.shareUrl} (shareId ${l.shareId}, scope ${opts.scope})`,
  );
  const body = buildFileMessageBody(
    [
      buildFileObject({
        fileName: uploaded.name,
        base: site.base,
        ...ids,
        ...link,
        scope: opts.scope,
      }),
    ],
    opts.caption,
  );
  const posted = await stage(
    "message",
    () => postMessage(opts.chat, opts.conversationId, body),
    (p) => `arrival ${p.OriginalArrivalTime ?? "ok"}`,
  );
  return {
    fileName: uploaded.name,
    itemId: uploaded.id,
    shareUrl: link.shareUrl,
    arrival: String(posted.OriginalArrivalTime ?? "ok"),
    receipts,
  };
}

/** Offline plan for --dry-run: no network, ids left as placeholders. */
export function buildSendFilePlan(opts: {
  conversationId: string;
  path: string;
  fileName: string;
  size: number;
  scope: LinkScope;
  caption?: string;
  site: OneDriveSite | null;
}): Record<string, unknown> {
  const base = opts.site?.base ?? POST_UPLOAD;
  const fileObject = buildFileObject({
    fileName: opts.fileName,
    base,
    listItemUniqueId: POST_UPLOAD,
    listId: POST_UPLOAD,
    siteId: POST_UPLOAD,
    shareUrl: POST_UPLOAD,
    shareId: POST_UPLOAD,
    scope: opts.scope,
  });
  return {
    conversationId: opts.conversationId,
    site: opts.site
      ? { host: opts.site.host, base: opts.site.base }
      : "not derivable offline (no TEAMS_SP_TOKEN or cached OneDrive token)",
    file: {
      path: opts.path,
      name: opts.fileName,
      size: opts.size,
      fileType: fileExtension(opts.fileName),
      note: "the server may rename on conflict; the returned name is used",
    },
    upload: {
      method: "PUT",
      path:
        `/_api/v2.1/drive/root:/${encodePathSegments(CHAT_FILES_FOLDER, opts.fileName)}` +
        ":/content?@name.conflictBehavior=rename",
    },
    linkScope: opts.scope,
    permissionScope: PERMISSION_SCOPE[opts.scope],
    body: buildFileMessageBody([fileObject], opts.caption),
  };
}
