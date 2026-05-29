import { getRegion, getSharepointToken, getToken } from "./auth.js";
import { CLIENT_INFO } from "./config.js";

export interface ApiClient {
  fetch(url: string, init?: RequestInit): Promise<Response>;
  baseUrl: string;
}

async function requireToken(name: string, label: string): Promise<string> {
  const token = await getToken(name);
  if (!token) {
    console.error(
      `No valid ${label} token found. Run \`teams login\` to authenticate.`,
    );
    process.exit(1);
  }
  return token;
}

function createClient(
  baseUrl: string,
  headers: Record<string, string>,
): ApiClient {
  const controller = new AbortController();
  const timeout = 30_000;

  return {
    baseUrl,
    async fetch(url: string, init?: RequestInit): Promise<Response> {
      const fullUrl = url.startsWith("http") ? url : `${baseUrl}${url}`;
      const timer = setTimeout(() => controller.abort(), timeout);
      try {
        return await fetch(fullUrl, {
          ...init,
          headers: { ...headers, ...init?.headers },
          signal: AbortSignal.timeout(timeout),
        });
      } finally {
        clearTimeout(timer);
      }
    },
  };
}

export function getChatsvcBaseUrl(): string {
  const region = getRegion();
  return `https://teams.cloud.microsoft/api/chatsvc/${region}/v1/users/ME`;
}

export async function getChatClient(): Promise<ApiClient> {
  const token = await requireToken("ic3", "Teams chat");
  return createClient(getChatsvcBaseUrl(), {
    Authorization: `Bearer ${token}`,
    "x-ms-test-user": "False",
    "x-ms-migration": "True",
    behavioroverride: "redirectAs404",
    clientinfo: CLIENT_INFO,
  });
}

export async function getSearchClient(): Promise<ApiClient> {
  const token = await requireToken("search", "search");
  return createClient("https://substrate.office.com", {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
    "x-client-version": "T2.1",
  });
}

// Client for the SharePoint/Stream API that hosts meeting recordings and
// transcripts. The bearer token is audience-scoped per host, so the host must
// be supplied (derived from the recording's file URL). `recordingUrl` is the
// recording's share link, opened headlessly to intercept a token if we lack one.
export async function getStreamClient(
  host: string,
  recordingUrl: string,
): Promise<ApiClient> {
  const token = await getSharepointToken(host, recordingUrl);
  if (!token) {
    console.error(
      `Could not obtain a SharePoint token for ${host}.\n` +
        "Make sure you can open the recording in Teams, then try again (or re-run `teams login`).",
    );
    process.exit(1);
  }
  return createClient(`https://${host}`, {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
  });
}

export async function getPresenceClient(): Promise<ApiClient> {
  const token = await requireToken("presence", "presence");
  const region = getRegion();
  return createClient(`https://teams.cloud.microsoft/ups/${region}/v1`, {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  });
}
