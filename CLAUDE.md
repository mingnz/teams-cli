# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```sh
npm install                    # Install dependencies
npm run build                  # Build with tsup
npm run dev -- --help          # Run CLI in dev mode (no build)
npm test                       # Run all tests
npx vitest run tests/formatting.test.ts  # Run a single test file
```

## Architecture

This is a TypeScript CLI for Microsoft Teams that talks directly to the internal Teams Chat Service API (chatsvc) — the same HTTP API the Teams web client uses. It does **not** use Microsoft Graph.

### Module layers (dependencies flow downward only)

- **cli.ts** — Commander.js commands. Each command creates an API client, calls async API functions, formats results, and renders with cli-table3/chalk.
- **formatting.ts** — Pure functions transforming API response dicts into display-ready objects. No I/O.
- **api.ts** — Async functions making HTTP calls via the ApiClient wrapper. All take an `ApiClient` as first arg (for testability with mocked fetch).
- **client.ts** — Factory functions creating pre-configured `ApiClient` instances (thin wrappers around native `fetch`) for three API surfaces (chat, search, presence), each with its own token.
- **auth.ts** — Token persistence (`~/.teams-cli/tokens.json`) and Playwright-based browser login flow that extracts tokens from localStorage.
- **config.ts** — Constants only, no imports from other project modules.

### API surfaces

| API | Token name | Base URL pattern |
|-----|-----------|-----------------|
| Chat Service (chatsvc) | `ic3` | `teams.cloud.microsoft/api/chatsvc/{region}/v1/users/ME` |
| Substrate Search | `search` | `substrate.office.com` |
| Presence | `presence` | `teams.cloud.microsoft/ups/{region}/v1` |
| SharePoint Stream (transcripts) | `sharepoint[host]` | `https://{host}.sharepoint.com` |

### Key design details

- `sendMessage()` wraps content in `<p>` tags and generates a random 19-digit `clientmessageid` using BigInt. Both plain and HTML sends use `messagetype: "RichText/Html"`; `send --html/--html-file` skip the `<p>` wrap and post the caller's HTML untouched (no escaping — caller's job). Body construction (`buildMessageBody`) is split from the POST (`postMessage`) so `--dry-run` can print the exact body
- **`send-file`** (`upload.ts`) = upload to the sender's OneDrive `Microsoft Teams Chat Files/` → `sharepointIds` → `createLink` → chat POST via `postMessage()` with `buildMessageBody(caption, {html, amsreferences: [], properties})`. Gotchas:
  - The file object in `properties.files` is **field-proven** against the Teams client (field set + order; golden fixture in `tests/upload.test.ts`) — silent changes break the card render. `properties.files` is a JSON **string**
  - Site is derived from the SP token (`aud` → `<tenant>-my.sharepoint.com`, `upn`/`unique_name` → personal-site slug); never hardcode tenant/user
  - `conflictBehavior=rename` (never `replace`, which would overwrite a file already shared elsewhere); use the server-returned name downstream
  - OneDrive calls use plain `fetch` with explicit SP headers — never the chat client (it attaches ic3 auth)
  - Token ladder (`getOneDriveToken`): `TEAMS_SP_TOKEN` env → cached `tokens.sharepoint[*-my host]` → MSAL localStorage in the CLI profile → `getSharepointToken()` interception. Token is never logged or put in argv
  - Stages `upload → ids → link → message` each log a receipt to stderr; a failure exits 1 listing which stages succeeded (orphans). Both tokens are resolved before the upload. `--dry-run` makes no network call and opens no browser
- `createDmThread()` uses `POST /threads` (without `/users/ME`) with `uniquerosterthread: true` for idempotent 1:1 chat creation
- `searchPeople()` uses Substrate Suggestions API (`/search/api/v1/suggestions?scenario=peoplepicker.newChat`), same endpoint as the Teams web client people picker
- `getMyMri()` extracts the user's MRI from the ic3 JWT token's `oid` claim — no extra API call needed
- `pollMessages()` and `pollConversations()` use **delta sync** via `startTime` + `syncState` — first call anchors at current time (returns nothing), subsequent calls with the returned `syncState` URL return only new items
- The `watch` command creates a new `ApiClient` per poll iteration (tokens may refresh between polls)
- Members are at `/threads/{id}`, not `/conversations/{id}/members`
- Activity feed reuses `getMessages()` with system conversation IDs (`48:notifications`, `48:mentions`, `48:calllogs`)
- System conversations (ID prefix `48:`) are filtered out of chat listings
- Login uses a **persistent browser profile** (`~/.teams-cli/browser-profile/`) so session cookies survive for silent token refresh
- Expired tokens are automatically refreshed via headless Playwright before each command — no user interaction unless the session itself has expired
- Playwright is dynamically imported at runtime (`await import("playwright")`) so it's not required at install time
- Native `fetch` is used for all HTTP calls (no external HTTP library needed with Node 18+)
- **Meeting transcripts** (`recordings`/`transcript` commands) use the SharePoint Stream API, not Graph:
  - `parseRecordings()` (formatting.ts) scans a meeting chat's messages for embedded SharePoint URLs (`.mp4` or `stream.aspx?id=`), normalizing Stream playback links to the underlying file path and deriving the host from the URL
  - `resolveDriveItem()` turns a file URL into `{driveId, itemId}` via `GET /_api/v2.0/shares/u!{base64url(url)}/driveItem`
  - `getTranscriptMetadata()` calls `GET /_api/v2.1/drives/{driveId}/items/{itemId}?$expand=media/transcripts` and tolerates three response shapes (`media.transcripts[]`, `value[]`, single object), returning the `temporaryDownloadUrl`
  - `downloadTranscriptJson()` fetches that URL with `?format=json`; `convertTranscriptToVtt()`/`convertTranscriptToGrouped()` turn the JSON `entries[]` into WebVTT or speaker-grouped text
- SharePoint tokens are **audience-scoped per host** and, unlike ic3/search/presence, are **not stored in localStorage** — the SPA holds them in memory and only sends them on the wire. So `acquireSharepointToken(host, recordingUrl)` drives the persistent (signed-in) browser profile headlessly and **intercepts the `Bearer` token off an `/_api/` request**. It navigates to the **host root first** (`https://{host}/`) — the Teams-only login never visits SharePoint, so this is what triggers the SSO login redirect that establishes a SharePoint session; the OneDrive/site SPA then makes an authenticated `/_api/` call. The recording link is only a fallback warmup (and some links 404). The token is cached in `tokens.sharepoint[host]` (with `expires_on` from the JWT `exp`) and reused until expiry. Tokens shorter than ~100 chars are ignored (some requests carry a tiny placeholder auth header)

## Conventions

- Use [Conventional Commits](https://www.conventionalcommits.org/) for all commit messages (e.g. `feat:`, `fix:`, `refactor:`, `docs:`, `test:`).
- PR titles must also follow the Conventional Commits format.
- **Never commit PII** (names, emails, user IDs, conversation content) to the repo — including in test fixtures, PR descriptions, commit messages, and comments.

## Maintenance

When making changes to features or architecture, update the relevant docs in the same PR:
- **CLAUDE.md** — key guidance for maintainers (including AI agents): design decisions, gotchas, commands, conventions
- **docs/ARCHITECTURE.md** — module responsibilities, data flow, testing strategy
- **README.md** — user-facing usage examples and install instructions
