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
  [key: string]: TokenEntry | string | undefined;
  region?: string;
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
  const expiresOn = entry.expires_on ?? 0;
  return Date.now() / 1000 > Number(expiresOn);
}

// The IC3 audience used for detecting login completion
const IC3_AUDIENCE = TOKEN_AUDIENCES.ic3;

async function tryRefresh(): Promise<boolean> {
  let playwright: typeof import("playwright");
  try {
    playwright = await import("playwright");
  } catch {
    return false;
  }

  try {
    mkdirSync(BROWSER_PROFILE_DIR, { recursive: true });
    const context = await playwright.chromium.launchPersistentContext(
      BROWSER_PROFILE_DIR,
      {
        headless: true,
      },
    );
    const page = await context.newPage();
    await page.goto(TEAMS_URL);

    // Wait for MSAL to silently acquire tokens (up to 30s)
    // MSAL v2 localStorage keys use "|" as delimiter:
    //   msal.2|{oid}|{authority}|accesstoken|{clientId}|{tid}|{scopes}|
    // Scopes are space-separated URLs. We parse each URL's hostname to
    // match the audience exactly, avoiding CodeQL URL-substring warnings.
    const deadline = Date.now() + 30_000;
    let found = false;
    while (Date.now() < deadline) {
      try {
        found = await page.evaluate((aud: string) => {
          const target = aud.toLowerCase();
          for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key !== null) {
              const parts = key.split("|");
              if (parts[3]?.toLowerCase() === "accesstoken") {
                const scopes = (parts[6] ?? "").split(" ");
                const prefix = `https://${target}/`;
                if (scopes.some((s) => s.toLowerCase().startsWith(prefix))) {
                  return true;
                }
              }
            }
          }
          return false;
        }, IC3_AUDIENCE);
        if (found) break;
      } catch {
        // page may not be ready
      }
      await page.waitForTimeout(1000);
    }

    if (!found) {
      await context.close();
      return false;
    }

    await page.waitForTimeout(2000);

    const tokens: TokenStore = loadTokens() ?? {};
    for (const [name, audience] of Object.entries(TOKEN_AUDIENCES)) {
      const result = await page.evaluate((aud: string) => {
        const target = aud.toLowerCase();
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key !== null) {
            const parts = key.split("|");
            if (parts[3]?.toLowerCase() === "accesstoken") {
              const scopes = (parts[6] ?? "").split(" ");
              const prefix = `https://${target}/`;
              if (scopes.some((s) => s.toLowerCase().startsWith(prefix))) {
                const data = JSON.parse(localStorage.getItem(key)!);
                return { secret: data.secret, expires_on: data.expiresOn };
              }
            }
          }
        }
        return null;
      }, audience);
      if (result) tokens[name] = result;
    }

    await context.close();
    saveTokens(tokens);
    return true;
  } catch {
    return false;
  }
}

export async function getToken(name: string): Promise<string | null> {
  const tokens = loadTokens();
  if (!tokens) return null;
  const entry = tokens[name];
  if (!entry || typeof entry === "string") return null;
  if (isExpired(entry)) {
    if (await tryRefresh()) {
      const refreshed = loadTokens();
      const refreshedEntry = refreshed?.[name];
      if (
        refreshedEntry &&
        typeof refreshedEntry !== "string" &&
        !isExpired(refreshedEntry)
      ) {
        return refreshedEntry.secret;
      }
    }
    return null;
  }
  return entry.secret;
}

export function getMyMri(): string | null {
  const tokens = loadTokens();
  if (!tokens) return null;
  const entry = tokens.ic3;
  if (!entry || typeof entry === "string") return null;
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
  const page = await context.newPage();

  console.log("Opening Teams login page...");
  await page.goto(TEAMS_URL);
  console.log("Please sign in to Microsoft Teams in the browser window.");
  console.log('When prompted "Stay signed in?", click Yes.');
  console.log("Waiting for login to complete (timeout: 5 minutes)...");

  // MSAL v2 localStorage keys use "|" as delimiter:
  //   msal.2|{oid}|{authority}|accesstoken|{clientId}|{tid}|{scopes}|
  // Scopes are space-separated URLs. We parse each URL's hostname to
  // match the audience exactly, avoiding CodeQL URL-substring warnings.
  const deadline = Date.now() + 300_000;
  let found = false;
  while (Date.now() < deadline) {
    try {
      found = await page.evaluate((aud: string) => {
        const target = aud.toLowerCase();
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key !== null) {
            const parts = key.split("|");
            if (parts[3]?.toLowerCase() === "accesstoken") {
              const scopes = (parts[6] ?? "").split(" ");
              const prefix = `https://${target}/`;
              if (scopes.some((s) => s.toLowerCase().startsWith(prefix))) {
                return true;
              }
            }
          }
        }
        return false;
      }, IC3_AUDIENCE);
      if (found) break;
    } catch {
      // page may not be ready
    }
    await page.waitForTimeout(2000);
  }

  if (!found) {
    await context.close();
    throw new Error("Login timed out after 5 minutes.");
  }

  // Wait for MSAL to acquire all tokens (search/presence tokens are lazy)
  await page.waitForTimeout(8000);
  console.log("Login detected! Extracting tokens...");

  for (const [name, audience] of Object.entries(TOKEN_AUDIENCES)) {
    const result = await page.evaluate((aud: string) => {
      const target = aud.toLowerCase();
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key !== null) {
          const parts = key.split("|");
          if (parts[3]?.toLowerCase() === "accesstoken") {
            const scopes = (parts[6] ?? "").split(" ");
            const prefix = `https://${target}/`;
            if (scopes.some((s) => s.toLowerCase().startsWith(prefix))) {
              const data = JSON.parse(localStorage.getItem(key)!);
              return { secret: data.secret, expires_on: data.expiresOn };
            }
          }
        }
      }
      return null;
    }, audience);
    if (result) tokens[name] = result;
  }

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

  await context.close();

  const validTokens = Object.entries(tokens).filter(
    ([k, v]) =>
      k !== "region" &&
      typeof v === "object" &&
      v !== null &&
      "secret" in v &&
      v.secret,
  );
  if (validTokens.length === 0) {
    logout();
    throw new Error(
      'No valid tokens captured. Did you click "Yes" on "Stay signed in?"? Please run `teams login` again and select Yes.',
    );
  }

  saveTokens(tokens);
  return tokens;
}
