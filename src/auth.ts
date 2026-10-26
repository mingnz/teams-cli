import {
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import {
  BROWSER_PROFILE_DIR,
  DATA_DIR,
  REGION_KEY_PATTERN,
  TEAMS_URL,
  TOKEN_AUDIENCES,
  TOKENS_FILE,
} from "./config.js";

export interface TokenEntry {
  secret: string;
  expires_on: number | string;
}

export interface TokenStore {
  [key: string]: TokenEntry | string | Record<string, TokenEntry> | undefined;
  region?: string;
  // SharePoint/Stream tokens, keyed by host (e.g. "contoso-my.sharepoint.com").
  sharepoint?: Record<string, TokenEntry>;
}

export function logout(): void {
  if (existsSync(TOKENS_FILE)) rmSync(TOKENS_FILE);
  if (existsSync(BROWSER_PROFILE_DIR))
    rmSync(BROWSER_PROFILE_DIR, { recursive: true });
}

export function saveTokens(tokens: TokenStore): void {
  mkdirSync(DATA_DIR, { recursive: true });
  writeFileSync(TOKENS_FILE, JSON.stringify(tokens, null, 2));
}

export function loadTokens(): TokenStore | null {
  if (!existsSync(TOKENS_FILE)) return null;
  return JSON.parse(readFileSync(TOKENS_FILE, "utf-8"));
}

function isExpired(entry: TokenEntry): boolean {
  const expiresOn = Number(entry.expires_on);
  return !Number.isFinite(expiresOn) || expiresOn <= Date.now() / 1000;
}

// A token store value is a single token (vs. a string like `region` or the
// `sharepoint` host map) when it has a `secret` field.
function isTokenEntry(
  value: TokenEntry | string | Record<string, TokenEntry> | undefined,
): value is TokenEntry {
  return (
    typeof value === "object" &&
    value !== null &&
    "secret" in value &&
    typeof value.secret === "string" &&
    value.secret.length > 0
  );
}

// The IC3 audience used for detecting login completion
const IC3_AUDIENCE = TOKEN_AUDIENCES.ic3;

// Browser-side function to find an MSAL v2 access token by audience.
// MSAL v2 localStorage keys use "|" as delimiter:
//   msal.2|{oid}|{authority}|accesstoken|{clientId}|{tid}|{scopes}|
// Scopes are space-separated URLs. We match the audience as a URL prefix
// to avoid CodeQL URL-substring warnings.
/* v8 ignore start -- runs inside browser via page.evaluate, not testable in Node */
function findMsalToken(
  aud: string,
): { secret: string; expires_on: string } | null {
  const prefix = `https://${aud.toLowerCase()}/`;
  const now = Date.now() / 1000;
  let newest: { secret: string; expires_on: string } | null = null;
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key === null) continue;
    const parts = key.split("|");
    if (parts[3]?.toLowerCase() !== "accesstoken") continue;
    const scopes = (parts[6] ?? "").split(" ");
    if (!scopes.some((s) => s.toLowerCase().startsWith(prefix))) continue;
    const stored = localStorage.getItem(key);
    if (stored === null) continue;
    const data = JSON.parse(stored);
    const expiresOn = Number(data.expiresOn);
    if (
      typeof data.secret !== "string" ||
      !data.secret ||
      !Number.isFinite(expiresOn) ||
      expiresOn <= now
    )
      continue;
    if (newest === null || expiresOn > Number(newest.expires_on)) {
      newest = { secret: data.secret, expires_on: String(expiresOn) };
    }
  }
  return newest;
}

/* v8 ignore stop */

// Poll until an unexpired access token for the requested audience is available.
/* v8 ignore start -- requires live Playwright page */
async function waitForToken(
  page: {
    evaluate: <T>(fn: (aud: string) => T, arg: string) => Promise<T>;
    waitForTimeout: (ms: number) => Promise<void>;
  },
  audience: string,
  timeoutMs: number,
  pollMs: number,
): Promise<boolean> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const found = await page.evaluate(findMsalToken, audience);
      if (found) return true;
    } catch {
      // page may not be ready
    }
    await page.waitForTimeout(pollMs);
  }
  return false;
}

/* v8 ignore stop */

// Extract all configured tokens from the page's localStorage.
/* v8 ignore start -- requires live Playwright page */
async function extractTokens(
  page: { evaluate: <T>(fn: (aud: string) => T, arg: string) => Promise<T> },
  audiences: Record<string, string>,
  existing?: TokenStore,
): Promise<TokenStore> {
  const tokens: TokenStore = existing ?? {};
  for (const [name, audience] of Object.entries(audiences)) {
    const result = await page.evaluate(findMsalToken, audience);
    if (result) tokens[name] = result;
  }
  return tokens;
}

/* v8 ignore stop */

// Read the `exp` claim (epoch seconds) from a JWT, for the token store's
// `expires_on`. Falls back to ~1h from now if the token can't be decoded.
function jwtExpiry(secret: string): number {
  try {
    const payload = secret.split(".")[1];
    const padded = payload + "=".repeat((4 - (payload.length % 4)) % 4);
    const claims = JSON.parse(Buffer.from(padded, "base64url").toString());
    if (typeof claims.exp === "number") return claims.exp;
  } catch {
    // not a decodable JWT
  }
  return Math.floor(Date.now() / 1000) + 3600;
}

function saveSharepointToken(host: string, secret: string): void {
  const tokens = loadTokens() ?? {};
  const map: Record<string, TokenEntry> = { ...(tokens.sharepoint ?? {}) };
  map[host.toLowerCase()] = { secret, expires_on: jwtExpiry(secret) };
  tokens.sharepoint = map;
  saveTokens(tokens);
}

/* v8 ignore start -- requires live Playwright browser */
async function tryRefresh(requiredAudience?: string): Promise<boolean> {
  let playwright: typeof import("playwright");
  try {
    playwright = await import("playwright");
  } catch {
    console.error(
      "Automatic Teams token refresh requires Playwright. Run `teams login` to reconnect.",
    );
    return false;
  }

  let context: import("playwright").BrowserContext | undefined;
  let stage = "browser launch";
  try {
    mkdirSync(BROWSER_PROFILE_DIR, { recursive: true });
    context = await playwright.chromium.launchPersistentContext(
      BROWSER_PROFILE_DIR,
      { headless: true },
    );
    stage = "Teams page navigation";
    const page = await context.newPage();
    await page.goto(TEAMS_URL);

    if (!(await waitForToken(page, IC3_AUDIENCE, 30_000, 1000))) {
      console.error(
        "Automatic Teams refresh did not produce an unexpired chat token within 30 seconds. Run `teams login` to complete sign-in.",
      );
      return false;
    }
    if (
      requiredAudience &&
      !(await waitForToken(page, requiredAudience, 30_000, 1000))
    ) {
      console.error(
        "Automatic Teams refresh did not produce the requested resource token within 30 seconds. Run `teams login` to reconnect.",
      );
      return false;
    }

    stage = "token extraction";
    await page.waitForTimeout(2000);
    const tokens = await extractTokens(
      page,
      TOKEN_AUDIENCES,
      loadTokens() ?? {},
    );
    const chat = tokens.ic3;
    if (!isTokenEntry(chat) || isExpired(chat)) {
      console.error(
        "Automatic Teams refresh did not capture a valid chat token. Run `teams login` to reconnect.",
      );
      return false;
    }
    stage = "token persistence";
    saveTokens(tokens);
    return true;
  } catch {
    console.error(
      `Automatic Teams token refresh failed during ${stage}. Run \`teams login\` to reconnect.`,
    );
    return false;
  } finally {
    await context?.close();
  }
}
/* v8 ignore stop */

// SharePoint never stores its access token in localStorage — the SPA holds it in
// memory and only sends it on the wire. So to get a token for a host we don't
// have, drive the persistent (already-signed-in) browser profile to that host and
// intercept the Bearer token off an `/_api/` request. Visiting the host root
// first bootstraps the SharePoint session via SSO (the login redirect that the
// Teams-only login never triggers); the OneDrive/site SPA then makes an
// authenticated `/_api/` call. The recording link is used as a fallback warmup.
/* v8 ignore start -- requires live Playwright browser */
async function acquireSharepointToken(
  host: string,
  warmupUrl: string,
): Promise<boolean> {
  let playwright: typeof import("playwright");
  try {
    playwright = await import("playwright");
  } catch {
    return false;
  }

  let context: import("playwright").BrowserContext | undefined;
  try {
    mkdirSync(BROWSER_PROFILE_DIR, { recursive: true });
    context = await playwright.chromium.launchPersistentContext(
      BROWSER_PROFILE_DIR,
      { headless: true },
    );
    const page = await context.newPage();

    const wantHost = host.toLowerCase();
    const captured = new Promise<string | null>((resolve) => {
      const timer = setTimeout(() => resolve(null), 45_000);
      page.on("request", (req) => {
        const url = req.url();
        if (!/\/_api\//i.test(url)) return;
        try {
          if (new URL(url).host.toLowerCase() !== wantHost) return;
        } catch {
          return;
        }
        const token = req
          .headers()
          .authorization?.match(/^Bearer\s+(.+)$/i)?.[1];
        // Ignore non-JWT placeholders (some calls send a tiny dummy header).
        if (token && token.length > 100) {
          clearTimeout(timer);
          resolve(token);
        }
      });
    });

    // Bootstrap the SharePoint session at the host root (triggers the SSO login
    // redirect); if no token shows up shortly, fall back to the recording page.
    await page
      .goto(`https://${host}/`, { waitUntil: "domcontentloaded" })
      .catch(() => {});
    const fallback = setTimeout(() => {
      page.goto(warmupUrl, { waitUntil: "domcontentloaded" }).catch(() => {});
    }, 15_000);
    const secret = await captured;
    clearTimeout(fallback);
    await context.close();

    if (!secret) return false;
    saveSharepointToken(host, secret);
    return true;
  } catch {
    await context?.close().catch(() => {});
    return false;
  }
}
/* v8 ignore stop */

export async function getToken(
  name: string,
  refreshMissing = false,
): Promise<string | null> {
  const tokens = loadTokens();
  if (!tokens) return null;
  const entry = tokens[name];
  if (
    !isTokenEntry(entry) &&
    !(refreshMissing && Object.hasOwn(TOKEN_AUDIENCES, name))
  )
    return null;
  if (!isTokenEntry(entry) || isExpired(entry)) {
    if (await tryRefresh(refreshMissing ? TOKEN_AUDIENCES[name] : undefined)) {
      const refreshedEntry = loadTokens()?.[name];
      if (isTokenEntry(refreshedEntry) && !isExpired(refreshedEntry)) {
        return refreshedEntry.secret;
      }
    }
    return null;
  }
  return entry.secret;
}

// Look up a SharePoint token for a specific host, refreshing if expired.
function readSharepointEntry(host: string): TokenEntry | null {
  const tokens = loadTokens();
  const map = tokens?.sharepoint;
  if (!map) return null;
  const entry = map[host.toLowerCase()];
  return entry ?? null;
}

export async function getSharepointToken(
  host: string,
  warmupUrl: string,
): Promise<string | null> {
  let entry = readSharepointEntry(host);
  if (entry && !isExpired(entry)) return entry.secret;

  // Missing or expired: open the recording to intercept a fresh token, re-read.
  if (await acquireSharepointToken(host, warmupUrl)) {
    entry = readSharepointEntry(host);
    if (entry && !isExpired(entry)) return entry.secret;
  }
  return null;
}

export function getMyMri(): string | null {
  const tokens = loadTokens();
  if (!tokens) return null;
  const entry = tokens.ic3;
  if (!isTokenEntry(entry)) return null;
  const secret = entry.secret;
  if (!secret) return null;
  try {
    const payload = secret.split(".")[1];
    const padded = payload + "=".repeat((4 - (payload.length % 4)) % 4);
    const claims = JSON.parse(Buffer.from(padded, "base64url").toString());
    const oid = claims.oid;
    if (oid) return `8:orgid:${oid}`;
  } catch {
    // invalid JWT
  }
  return null;
}

export function getRegion(): string {
  const tokens = loadTokens();
  return (tokens?.region as string) ?? "amer";
}

/* v8 ignore start -- interactive browser login, not unit-testable */
export async function login(): Promise<TokenStore> {
  let playwright: typeof import("playwright");
  try {
    playwright = await import("playwright");
  } catch {
    throw new Error(
      "Playwright is required for login. Run:\n  npx playwright install chromium",
    );
  }

  const tokens: TokenStore = {};
  mkdirSync(BROWSER_PROFILE_DIR, { recursive: true });

  const context = await playwright.chromium.launchPersistentContext(
    BROWSER_PROFILE_DIR,
    {
      headless: false,
    },
  );
  try {
    const page = await context.newPage();

    console.log("Opening Teams login page...");
    await page.goto(TEAMS_URL);
    console.log("Please sign in to Microsoft Teams in the browser window.");
    console.log('When prompted "Stay signed in?", click Yes.');
    console.log("Waiting for login to complete (timeout: 5 minutes)...");

    if (!(await waitForToken(page, IC3_AUDIENCE, 300_000, 2000))) {
      throw new Error(
        "Login timed out after 5 minutes without an unexpired chat token. Existing session data was preserved.",
      );
    }

    // Wait for MSAL to acquire all tokens (search/presence tokens are lazy)
    await page.waitForTimeout(8000);
    console.log("Login detected! Extracting tokens...");

    await extractTokens(page, TOKEN_AUDIENCES, tokens);

    const regionData = await page.evaluate((pattern: string) => {
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key !== null && key.indexOf(pattern) !== -1) {
          const data = JSON.parse(localStorage.getItem(key)!);
          const gtm =
            typeof data.item === "string" ? JSON.parse(data.item) : data.item;
          const ams: string = gtm?.ams ?? "";
          const match = ams.match(/https:\/\/([a-z]+)-prod/);
          return match ? match[1] : null;
        }
      }
      return null;
    }, REGION_KEY_PATTERN);
    tokens.region = regionData ?? "amer";

    const chat = tokens.ic3;
    if (!isTokenEntry(chat) || isExpired(chat)) {
      throw new Error(
        "No unexpired chat token was captured. Existing session data was preserved; please run `teams login` again.",
      );
    }

    saveTokens(tokens);
    return tokens;
  } finally {
    await context.close();
  }
}
/* v8 ignore stop */
