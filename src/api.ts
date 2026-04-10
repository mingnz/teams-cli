import { randomUUID } from "node:crypto";
import type { ApiClient } from "./client.js";
import { getChatsvcBaseUrl } from "./client.js";

function randomClientMessageId(): string {
  const min = 10n ** 18n;
  const max = 10n ** 19n - 1n;
  const range = max - min;
  const rand = BigInt(Math.floor(Math.random() * Number(range)));
  return (min + rand).toString();
}

async function jsonOrThrow(resp: Response): Promise<unknown> {
  if (!resp.ok) {
    const body = await resp.text().catch(() => "");
    throw new Error(`HTTP ${resp.status}: ${body.slice(0, 200)}`);
  }
  return resp.json();
}

export async function listConversations(
  client: ApiClient,
  pageSize = 30,
): Promise<Record<string, unknown>[]> {
  const params = new URLSearchParams({ view: "msnp24Equivalent", pageSize: String(pageSize) });
  const resp = await client.fetch(`${client.baseUrl}/conversations?${params}`);
  const data = (await jsonOrThrow(resp)) as Record<string, unknown>;
  return (data.conversations as Record<string, unknown>[]) ?? [];
}

export async function getConversation(
  client: ApiClient,
  conversationId: string,
): Promise<Record<string, unknown>> {
  const params = new URLSearchParams({ view: "msnp24Equivalent" });
  const resp = await client.fetch(`${client.baseUrl}/conversations/${conversationId}?${params}`);
  return (await jsonOrThrow(resp)) as Record<string, unknown>;
}

export async function getMessages(
  client: ApiClient,
  conversationId: string,
  pageSize = 20,
): Promise<Record<string, unknown>[]> {
  const params = new URLSearchParams({
    view: "msnp24Equivalent|supportsMessageProperties",
    pageSize: String(pageSize),
  });
  const resp = await client.fetch(
    `${client.baseUrl}/conversations/${conversationId}/messages?${params}`,
  );
  const data = (await jsonOrThrow(resp)) as Record<string, unknown>;
  return (data.messages as Record<string, unknown>[]) ?? [];
}

export async function sendMessage(
  client: ApiClient,
  conversationId: string,
  content: string,
): Promise<Record<string, unknown>> {
  const clientMsgId = randomClientMessageId();
  const resp = await client.fetch(
    `${client.baseUrl}/conversations/${conversationId}/messages`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        content: `<p>${content}</p>`,
        messagetype: "RichText/Html",
        contenttype: "text",
        clientmessageid: clientMsgId,
      }),
    },
  );
  return (await jsonOrThrow(resp)) as Record<string, unknown>;
}

export async function getThreadMembers(
  client: ApiClient,
  conversationId: string,
): Promise<Record<string, unknown>[]> {
  const threadsBase = client.baseUrl.replace("/users/ME", "");
  const resp = await client.fetch(`${threadsBase}/threads/${conversationId}`);
  const data = (await jsonOrThrow(resp)) as Record<string, unknown>;
  return (data.members as Record<string, unknown>[]) ?? [];
}

export async function getActivity(
  client: ApiClient,
  feed = "notifications",
  pageSize = 20,
): Promise<Record<string, unknown>[]> {
  return getMessages(client, `48:${feed}`, pageSize);
}

export async function markAsRead(
  client: ApiClient,
  conversationId: string,
  messageId: string,
  messageVersion: string,
  clientMessageId = "0",
): Promise<void> {
  const horizon = `${messageId};${messageVersion};${clientMessageId}`;
  const params = new URLSearchParams({ name: "consumptionhorizon" });
  const resp = await client.fetch(
    `${client.baseUrl}/conversations/${conversationId}/properties?${params}`,
    {
      method: "PUT",
      headers: { "Content-Type": "text/json" },
      body: JSON.stringify({ consumptionhorizon: horizon }),
    },
  );
  if (!resp.ok) {
    const body = await resp.text().catch(() => "");
    throw new Error(`HTTP ${resp.status}: ${body.slice(0, 200)}`);
  }
}

export async function searchPeople(
  client: ApiClient,
  query: string,
  size = 10,
): Promise<Record<string, unknown>[]> {
  const cvid = randomUUID();
  const logicalId = randomUUID();
  const body = {
    EntityRequests: [
      {
        Query: { QueryString: query, DisplayQueryString: query },
        EntityType: "People",
        Size: size,
        Fields: [
          "Id",
          "MRI",
          "DisplayName",
          "EmailAddresses",
          "PeopleType",
          "PeopleSubtype",
          "UserPrincipalName",
          "GivenName",
          "Surname",
          "ExternalDirectoryObjectId",
          "CompanyName",
          "JobTitle",
          "Department",
        ],
        Filter: {
          And: [
            { Or: [{ Term: { PeopleType: "Person" } }, { Term: { PeopleType: "Other" } }] },
            {
              Or: [
                { Term: { PeopleSubtype: "OrganizationUser" } },
                { Term: { PeopleSubtype: "MTOUser" } },
                { Term: { PeopleSubtype: "Guest" } },
              ],
            },
            { Or: [{ Term: { Flags: "NonHidden" } }] },
          ],
        },
        Provenances: ["Mailbox", "Directory"],
        From: 0,
        ServeNoEmailContacts: true,
      },
    ],
    Scenario: { Name: "peoplepicker.newChat" },
    Cvid: cvid,
    AppName: "Microsoft Teams",
    LogicalId: logicalId,
  };

  const resp = await client.fetch(
    "/search/api/v1/suggestions?scenario=peoplepicker.newChat",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-client-flights": "EnableSelfSuggestion,TeamsHiddenPeopleDirectorySearch",
      },
      body: JSON.stringify(body),
    },
  );
  const data = (await jsonOrThrow(resp)) as Record<string, unknown>;
  const groups = (data.Groups as Record<string, unknown>[]) ?? [];
  for (const group of groups) {
    if (group.Type === "People") {
      return (group.Suggestions as Record<string, unknown>[]) ?? [];
    }
  }
  return [];
}

export async function createDmThread(
  client: ApiClient,
  myMri: string,
  theirMri: string,
): Promise<string> {
  const threadsBase = client.baseUrl.replace("/users/ME", "");
  const resp = await client.fetch(`${threadsBase}/threads`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      members: [
        { id: myMri, role: "Admin" },
        { id: theirMri, role: "Admin" },
      ],
      properties: {
        threadType: "chat",
        fixedRoster: true,
        uniquerosterthread: true,
      },
    }),
  });
  if (!resp.ok) {
    const body = await resp.text().catch(() => "");
    throw new Error(`HTTP ${resp.status}: ${body.slice(0, 200)}`);
  }

  const location = resp.headers.get("location") ?? "";
  if (location.includes("/threads/")) {
    return location.split("/threads/").pop()!;
  }
  // Fallback: construct deterministic ID from sorted UUIDs
  const uuids = [myMri.replace("8:orgid:", ""), theirMri.replace("8:orgid:", "")].sort();
  return `19:${uuids[0]}_${uuids[1]}@unq.gbl.spaces`;
}

export async function pollMessages(
  client: ApiClient,
  conversationId: string,
  syncUrl?: string,
  pageSize = 50,
): Promise<{ messages: Record<string, unknown>[]; syncUrl: string }> {
  let resp: Response;
  if (syncUrl) {
    resp = await client.fetch(syncUrl);
  } else {
    const params = new URLSearchParams({
      view: "msnp24Equivalent|supportsMessageProperties",
      pageSize: String(pageSize),
      startTime: String(Date.now()),
    });
    resp = await client.fetch(
      `${client.baseUrl}/conversations/${conversationId}/messages?${params}`,
    );
  }
  const data = (await jsonOrThrow(resp)) as Record<string, unknown>;
  const messages = (data.messages as Record<string, unknown>[]) ?? [];
  const metadata = (data._metadata as Record<string, string>) ?? {};
  return { messages, syncUrl: metadata.syncState ?? "" };
}

export async function pollConversations(
  client: ApiClient,
  syncUrl?: string,
  pageSize = 30,
): Promise<{ conversations: Record<string, unknown>[]; syncUrl: string }> {
  let resp: Response;
  if (syncUrl) {
    resp = await client.fetch(syncUrl);
  } else {
    const params = new URLSearchParams({
      view: "msnp24Equivalent",
      pageSize: String(pageSize),
      startTime: String(Date.now()),
    });
    resp = await client.fetch(`${client.baseUrl}/conversations?${params}`);
  }
  const data = (await jsonOrThrow(resp)) as Record<string, unknown>;
  const conversations = (data.conversations as Record<string, unknown>[]) ?? [];
  const metadata = (data._metadata as Record<string, string>) ?? {};
  return { conversations, syncUrl: metadata.syncState ?? "" };
}

export async function searchMessages(
  client: ApiClient,
  query: string,
  size = 25,
  offset = 0,
): Promise<Record<string, unknown>> {
  const cvid = randomUUID();
  const logicalId = randomUUID();
  const body = {
    entityRequests: [
      {
        entityType: "Message",
        contentSources: ["Teams"],
        fields: [
          "Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String",
          "Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String",
          "Extension_SkypeSpaces_ConversationPost_Extension_SkypeGroupId_String",
        ],
        propertySet: "Optimized",
        query: { queryString: query, displayQueryString: query },
        from: offset,
        size,
        topResultsCount: 5,
      },
    ],
    QueryAlterationOptions: {
      EnableAlteration: true,
      EnableSuggestion: true,
      SupportedRecourseDisplayTypes: ["Suggestion"],
    },
    cvid,
    logicalId,
    scenario: {
      Dimensions: [
        { DimensionName: "QueryType", DimensionValue: "Messages" },
        { DimensionName: "FormFactor", DimensionValue: "general.web.reactSearch" },
      ],
      Name: "powerbar",
    },
    timezone: "UTC",
  };

  const resp = await client.fetch("/searchservice/api/v2/query", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return (await jsonOrThrow(resp)) as Record<string, unknown>;
}
