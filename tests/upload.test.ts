import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  buildFileMessageBody,
  buildFileObject,
  buildSendFilePlan,
  captionHtml,
  deriveOneDriveSite,
  encodePathSegments,
  fileExtension,
  type OneDriveSite,
  objectUrlFor,
  PERMISSION_SCOPE,
  parseLinkScope,
  personalSiteSlug,
  StageError,
  sendFile,
  uploadFileName,
} from "../src/upload.js";

function fakeJwt(claims: Record<string, unknown>): string {
  const enc = (o: unknown) =>
    Buffer.from(JSON.stringify(o)).toString("base64url");
  return `${enc({ alg: "none" })}.${enc(claims)}.sig`;
}

const SITE: OneDriveSite = {
  host: "contoso-my.sharepoint.com",
  siteUrl: "https://contoso-my.sharepoint.com/personal/a_b_contoso_com",
  base: "https://contoso-my.sharepoint.com/personal/a_b_contoso_com/",
};
const GUID = "11111111-2222-3333-4444-555555555555";
const LIST_ID = "aaaaaaaa-0000-0000-0000-000000000001";
const SITE_ID = "bbbbbbbb-0000-0000-0000-000000000002";
const SHARE_URL = "https://contoso-my.sharepoint.com/:t:/g/personal/x/AbCdEf";
const SHARE_ID = "share-123";

// Golden fixture: the field-proven file object, dynamic values parameterized.
function goldenFileObject(p: {
  fileName: string;
  ext: string;
  objectUrl: string;
  permissionScope: string;
}) {
  return {
    itemid: GUID,
    fileName: p.fileName,
    fileType: p.ext,
    fileInfo: {
      itemId: null,
      fileUrl: p.objectUrl,
      siteUrl: SITE.base,
      serverRelativeUrl: "",
      shareUrl: SHARE_URL,
      shareId: SHARE_ID,
    },
    fileChicletState: { serviceName: "p2p", state: "active" },
    "@type": "http://schema.skype.com/File",
    version: 2,
    id: GUID,
    baseUrl: SITE.base,
    objectUrl: p.objectUrl,
    openUrl: null,
    type: p.ext,
    title: p.fileName,
    state: "active",
    chicletBreadcrumbs: null,
    providerData: "",
    botFileProperties: {},
    isUploadError: null,
    progressComplete: null,
    permissionScope: p.permissionScope,
    filePreview: { previewUrl: "", previewHeight: 0, previewWidth: 0 },
    sharepointIds: {
      listId: LIST_ID,
      listItemUniqueId: GUID,
      siteId: SITE_ID,
      siteUrl: null,
      webId: null,
      driveId: null,
    },
    publication: null,
    site: null,
  };
}

function fileInput(fileName: string, scope: "anonymous" | "organization") {
  return {
    fileName,
    base: SITE.base,
    listItemUniqueId: GUID,
    listId: LIST_ID,
    siteId: SITE_ID,
    shareUrl: SHARE_URL,
    shareId: SHARE_ID,
    scope,
  };
}

describe("buildFileObject (golden fixture)", () => {
  it("matches the proven shape field-for-field, in order", () => {
    const built = buildFileObject(fileInput("notes.txt", "anonymous"));
    const golden = goldenFileObject({
      fileName: "notes.txt",
      ext: "txt",
      objectUrl: `${SITE.base}Documents/Microsoft%20Teams%20Chat%20Files/notes.txt`,
      permissionScope: "Anyone",
    });
    expect(built).toStrictEqual(golden);
    // Key order too (the payload is serialized as-is).
    expect(JSON.stringify(built)).toBe(JSON.stringify(golden));
  });

  it("maps link scope to permissionScope", () => {
    expect(PERMISSION_SCOPE).toEqual({
      anonymous: "Anyone",
      organization: "Organization",
    });
    expect(
      buildFileObject(fileInput("a.txt", "organization")).permissionScope,
    ).toBe("Organization");
    expect(
      buildFileObject(fileInput("a.txt", "anonymous")).permissionScope,
    ).toBe("Anyone");
  });
});

describe("buildFileMessageBody", () => {
  it("builds the proven message body with files as a JSON string", () => {
    const fo = buildFileObject(fileInput("notes.txt", "anonymous"));
    const body = buildFileMessageBody([fo]);
    expect(body).toStrictEqual({
      content: "",
      messagetype: "RichText/Html",
      contenttype: "text",
      amsreferences: [],
      clientmessageid: expect.stringMatching(/^\d{19}$/),
      properties: {
        mentions: "[]",
        cards: "[]",
        links: "[]",
        files: JSON.stringify([fo]),
        formatVariant: "TEAMS",
      },
    });
    expect(typeof body.properties?.files).toBe("string");
  });

  it("escapes the caption into <div><p>…</p></div>", () => {
    expect(captionHtml('a & b <i>"x"</i>')).toBe(
      "<div><p>a &amp; b &lt;i&gt;&quot;x&quot;&lt;/i&gt;</p></div>",
    );
    expect(captionHtml(undefined)).toBe("");
    expect(captionHtml("")).toBe("");
    expect(buildFileMessageBody([], "<b>").content).toBe(
      "<div><p>&lt;b&gt;</p></div>",
    );
  });
});

describe("names and encoding", () => {
  it("takes the basename and rejects empty names", () => {
    expect(uploadFileName("/tmp/dir/report.pdf")).toBe("report.pdf");
    expect(uploadFileName("report.pdf")).toBe("report.pdf");
    expect(() => uploadFileName("")).toThrow(/Invalid file name/);
    expect(() => uploadFileName("/")).toThrow(/Invalid file name/);
    expect(() => uploadFileName("..")).toThrow(/Invalid file name/);
  });

  it("lowercases the extension", () => {
    expect(fileExtension("Report.PDF")).toBe("pdf");
    expect(fileExtension("archive.tar.GZ")).toBe("gz");
    expect(fileExtension("README")).toBe("");
    expect(fileExtension(".env")).toBe("");
  });

  it("encodes Unicode / Arabic names per segment", () => {
    const name = "تقرير #1 ?.pdf";
    const enc = encodePathSegments("Microsoft Teams Chat Files", name);
    expect(enc).toBe(
      "Microsoft%20Teams%20Chat%20Files/" +
        "%D8%AA%D9%82%D8%B1%D9%8A%D8%B1%20%231%20%3F.pdf",
    );
    expect(decodeURIComponent(enc.split("/")[1])).toBe(name);
    expect(objectUrlFor(SITE.base, name)).toBe(`${SITE.base}Documents/${enc}`);
  });

  it("parses link scopes and aliases", () => {
    expect(parseLinkScope("org")).toBe("organization");
    expect(parseLinkScope("Organization")).toBe("organization");
    expect(parseLinkScope("anon")).toBe("anonymous");
    expect(parseLinkScope("anonymous")).toBe("anonymous");
    expect(parseLinkScope("users")).toBeNull();
  });
});

describe("deriveOneDriveSite", () => {
  it("derives host from aud and slug from upn", () => {
    const token = fakeJwt({
      aud: "https://contoso-my.sharepoint.com",
      upn: "A.B@contoso.com",
    });
    expect(deriveOneDriveSite(token)).toEqual(SITE);
  });

  it("falls back to unique_name and the capture host for a GUID aud", () => {
    const token = fakeJwt({
      aud: "00000003-0000-0ff1-ce00-000000000000",
      unique_name: "a.b@contoso.com",
    });
    expect(deriveOneDriveSite(token, "contoso-my.sharepoint.com")).toEqual(
      SITE,
    );
  });

  it("rejects non-OneDrive audiences and missing upn", () => {
    expect(() =>
      deriveOneDriveSite(
        fakeJwt({ aud: "https://contoso.sharepoint.com", upn: "a@b.com" }),
      ),
    ).toThrow(/not a OneDrive host/);
    expect(() =>
      deriveOneDriveSite(fakeJwt({ aud: "https://contoso-my.sharepoint.com" })),
    ).toThrow(/no upn/);
    expect(() => deriveOneDriveSite("garbage")).toThrow(/not a decodable JWT/);
  });

  it("slugs a UPN like SharePoint", () => {
    expect(personalSiteSlug("a.b@tenant.com")).toBe("a_b_tenant_com");
  });
});

// ── Staged flow with mocked network ──────────────────────────────────────

const CHAT_BASE = "https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME";

function jsonResp(status: number, body: unknown) {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
    text: async () => JSON.stringify(body),
  };
}

function setupNetwork(opts: { failAt?: string; serverName?: string } = {}): {
  spFetch: ReturnType<typeof vi.fn>;
  chatFetch: ReturnType<typeof vi.fn>;
} {
  const fail = (stage: string) => opts.failAt === stage;
  const spFetch = vi.fn(async (url: string, init?: RequestInit) => {
    if (init?.method === "PUT") {
      return fail("upload")
        ? jsonResp(404, { error: "notFound" })
        : jsonResp(201, { id: "ITEM1", name: opts.serverName ?? "notes.txt" });
    }
    if (url.includes("createLink")) {
      return fail("link")
        ? jsonResp(403, { error: "accessDenied" })
        : jsonResp(201, { id: SHARE_ID, link: { webUrl: SHARE_URL } });
    }
    return fail("ids")
      ? jsonResp(500, { error: "boom" })
      : jsonResp(200, {
          id: "ITEM1",
          sharepointIds: {
            listItemUniqueId: GUID,
            siteId: SITE_ID,
            listId: LIST_ID,
          },
        });
  });
  const chatFetch = vi.fn(async () =>
    fail("message")
      ? jsonResp(400, { message: "bad" })
      : jsonResp(201, { OriginalArrivalTime: 1700000000000 }),
  );
  vi.stubGlobal("fetch", spFetch);
  return { spFetch, chatFetch };
}

function run(chatFetch: ReturnType<typeof vi.fn>, fileName = "notes.txt") {
  return sendFile({
    chat: { baseUrl: CHAT_BASE, fetch: chatFetch },
    conversationId: "19:abc@thread.v2",
    site: SITE,
    token: "SP-TOKEN",
    fileName,
    data: new Uint8Array([1, 2, 3]),
    scope: "organization",
    caption: "see <this>",
  });
}

describe("sendFile", () => {
  beforeEach(() => vi.unstubAllGlobals());
  afterEach(() => vi.unstubAllGlobals());

  it("runs upload → ids → link → message with explicit SP headers", async () => {
    const { spFetch, chatFetch } = setupNetwork();
    const stages: string[] = [];
    const result = await sendFile({
      chat: { baseUrl: CHAT_BASE, fetch: chatFetch },
      conversationId: "19:abc@thread.v2",
      site: SITE,
      token: "SP-TOKEN",
      fileName: "notes.txt",
      data: new Uint8Array([1, 2, 3]),
      scope: "organization",
      onStage: (r) => stages.push(r.stage),
    });
    expect(stages).toEqual(["upload", "ids", "link", "message"]);
    expect(result.shareUrl).toBe(SHARE_URL);
    expect(result.arrival).toBe("1700000000000");

    const [putUrl, putInit] = spFetch.mock.calls[0];
    expect(putUrl).toBe(
      `${SITE.siteUrl}/_api/v2.1/drive/root:/Microsoft%20Teams%20Chat%20Files/notes.txt:/content?@name.conflictBehavior=rename`,
    );
    expect(putInit.method).toBe("PUT");
    expect(putInit.headers).toMatchObject({
      Authorization: "Bearer SP-TOKEN",
      "Content-Type": "application/octet-stream",
    });
    expect(spFetch.mock.calls[1][0]).toBe(
      `${SITE.siteUrl}/_api/v2.1/drive/items/ITEM1?select=id,sharepointIds`,
    );
    const [linkUrl, linkInit] = spFetch.mock.calls[2];
    expect(linkUrl).toBe(
      `${SITE.siteUrl}/_api/v2.1/drive/items/ITEM1/microsoft.graph.createLink`,
    );
    expect(JSON.parse(linkInit.body)).toEqual({
      type: "view",
      scope: "organization",
    });

    // Message goes through the chat client, never the SP fetch.
    const [msgUrl, msgInit] = chatFetch.mock.calls[0];
    expect(msgUrl).toBe(`${CHAT_BASE}/conversations/19:abc@thread.v2/messages`);
    const files = JSON.parse(JSON.parse(msgInit.body).properties.files);
    expect(files).toStrictEqual([
      goldenFileObject({
        fileName: "notes.txt",
        ext: "txt",
        objectUrl: objectUrlFor(SITE.base, "notes.txt"),
        permissionScope: "Organization",
      }),
    ]);
  });

  it("uses the server-returned (renamed) name downstream", async () => {
    const { spFetch, chatFetch } = setupNetwork({
      serverName: "Report 1.PDF",
    });
    const result = await run(chatFetch, "Report.PDF");
    expect(spFetch.mock.calls[0][0]).toContain("/Report.PDF:/content");
    expect(result.fileName).toBe("Report 1.PDF");
    const body = JSON.parse(chatFetch.mock.calls[0][1].body);
    const [fo] = JSON.parse(body.properties.files);
    expect(fo.fileName).toBe("Report 1.PDF");
    expect(fo.title).toBe("Report 1.PDF");
    expect(fo.fileType).toBe("pdf");
    expect(fo.type).toBe("pdf");
    expect(fo.objectUrl).toBe(
      `${SITE.base}Documents/Microsoft%20Teams%20Chat%20Files/Report%201.PDF`,
    );
    expect(fo.fileInfo.fileUrl).toBe(fo.objectUrl);
    expect(body.content).toBe("<div><p>see &lt;this&gt;</p></div>");
  });

  it("encodes a Unicode name in the upload URL", async () => {
    const name = "تقرير.txt";
    const { spFetch, chatFetch } = setupNetwork({ serverName: name });
    await run(chatFetch, name);
    expect(spFetch.mock.calls[0][0]).toContain(
      `/Microsoft%20Teams%20Chat%20Files/${encodeURIComponent(name)}:/content`,
    );
  });

  it.each([
    ["upload", []],
    ["ids", ["upload"]],
    ["link", ["upload", "ids"]],
    ["message", ["upload", "ids", "link"]],
  ])("a failure at %s reports the stages that succeeded", async (failAt, done) => {
    const { spFetch, chatFetch } = setupNetwork({ failAt });
    const err = await run(chatFetch).catch((e) => e);
    expect(err).toBeInstanceOf(StageError);
    expect(err.stage).toBe(failAt);
    expect(
      (err as StageError).completed.map((r: { stage: string }) => r.stage),
    ).toEqual(done);
    // Nothing runs past the failed stage.
    expect(chatFetch).toHaveBeenCalledTimes(failAt === "message" ? 1 : 0);
    expect(spFetch).toHaveBeenCalledTimes(Math.min(done.length + 1, 3));
  });

  it("names the derived site when the first drive call fails", async () => {
    const { chatFetch } = setupNetwork({ failAt: "upload" });
    const err = await run(chatFetch).catch((e) => e);
    expect(err.message).toContain("HTTP 404");
    expect(err.message).toContain(SITE.siteUrl);
  });
});

describe("buildSendFilePlan", () => {
  it("uses placeholders for post-upload ids", () => {
    const plan = buildSendFilePlan({
      conversationId: "19:abc@thread.v2",
      path: "./notes.txt",
      fileName: "notes.txt",
      size: 3,
      scope: "anonymous",
      site: SITE,
    });
    expect(plan.site).toEqual({ host: SITE.host, base: SITE.base });
    expect(plan.permissionScope).toBe("Anyone");
    const body = plan.body as { properties: { files: string } };
    const [fo] = JSON.parse(body.properties.files);
    expect(fo.id).toBe("<post-upload>");
    expect(fo.fileInfo.shareUrl).toBe("<post-upload>");
    expect(fo.sharepointIds.siteId).toBe("<post-upload>");
    expect(fo.baseUrl).toBe(SITE.base);
  });
});
