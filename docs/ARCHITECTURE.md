# Architecture

## Overview

teams-cli is a TypeScript CLI that talks directly to the internal Microsoft Teams Chat Service API — the same HTTP API that the Teams web client uses. It bypasses the Microsoft Graph SDK entirely, which avoids the need for Azure app registrations or admin-granted API permissions.

Authentication is handled by launching a real browser via Playwright, letting the user sign in normally (including MFA), then extracting tokens from the browser's `localStorage`.

## Layered design

The codebase follows a strict layered architecture where each module has a single responsibility and dependencies only flow downward:

```
cli.ts          Commands (user-facing)
  |
  +-- formatting.ts   Pure display logic (no I/O)
  +-- api.ts          HTTP request/response logic
        |
        +-- client.ts     HTTP client construction
              |
              +-- auth.ts      Token storage and browser login
              +-- config.ts    Constants (no dependencies)
```

### config.ts

All constants live here: file paths, token audience strings, API headers, default page sizes, and the Teams login URL. No imports from other project modules — this is the leaf of the dependency tree.

Key values:
- `DATA_DIR` — `~/.teams-cli/`, where tokens and caches are stored
- `TOKEN_AUDIENCES` — maps token names (`ic3`, `search`, `presence`) to the localStorage key patterns used to find them
- `CLIENT_INFO` — the `clientinfo` header that the chatsvc API requires

### auth.ts

Handles token persistence, automatic refresh, and the Playwright login flow.

- `saveTokens()` / `loadTokens()` — JSON round-trip to `~/.teams-cli/tokens.json`
- `getToken()` — check token expiry, attempt silent refresh if expired, and retrieve by name
- `getRegion()` — returns the user's region code (e.g. `au`, `amer`) extracted during login
- `getMyMri()` — extracts the current user's MRI (`8:orgid:{oid}`) by decoding the `oid` claim from the ic3 JWT token using `Buffer.from(payload, 'base64url')`. Used by the `dm` command to construct thread creation requests.
- `tryRefresh()` — silently refreshes expired tokens by launching a headless Chromium with the persistent browser profile (`~/.teams-cli/browser-profile/`). MSAL re-acquires tokens using the stored session cookies — no user interaction needed. Returns `true` on success, `false` if the session itself has expired.
- `login()` — the interactive login flow:
  1. Launches Chromium (headless=false) with a persistent browser profile so session cookies are preserved
  2. Navigates to `teams.cloud.microsoft`
  3. Polls `localStorage` every 2 seconds for up to 5 minutes until the IC3 token appears
  4. Extracts all three tokens (ic3, search, presence) and the region from `localStorage`
  5. Saves everything to disk
- `logout()` — removes stored tokens and the browser profile directory

Playwright is dynamically imported (`await import("playwright")`) so the CLI doesn't fail at startup if Playwright isn't installed — it only fails when login or token refresh is actually needed.

### client.ts

Factory functions that create `ApiClient` instances — thin wrappers around native `fetch` with preset headers and base URLs. Each client targets a different API surface:

| Factory | API | Token | Base URL |
|---------|-----|-------|----------|
| `getChatClient()` | Chat Service (chatsvc) | `ic3` | (per-request, includes region) |
| `getSearchClient()` | Substrate Search | `search` | `substrate.office.com` |
| `getPresenceClient()` | Presence | `presence` | `teams.cloud.microsoft/ups/{region}/v1` |

All clients use `AbortSignal.timeout(30000)` for request timeouts. If the required token is missing or expired, the client factory prints an error and exits.

### api.ts

Async functions that make HTTP calls via the `ApiClient` wrapper. Each function takes an `ApiClient` as its first argument, making them easy to test with mocked `fetch` functions. No formatting or display logic lives here.

Key design decisions:
- `sendMessage()` wraps content in `<p>` tags and generates a random 19-digit `clientmessageid` using BigInt arithmetic, matching the Teams web client's behaviour
- `getThreadMembers()` uses the `/threads/{id}` endpoint rather than `/conversations/{id}/members` (which returns 404)
- `searchMessages()` builds the full Substrate Search API request body including correlation IDs and dimension metadata
- `searchPeople()` uses the Substrate Suggestions API (`/search/api/v1/suggestions?scenario=peoplepicker.newChat`) — the same people picker endpoint the Teams web client uses when composing a new chat
- `createDmThread()` creates (or retrieves) a 1:1 chat thread via `POST /threads` with `uniquerosterthread: true`, which makes it idempotent — it returns the existing thread if one already exists
- `pollMessages()` and `pollConversations()` implement delta sync polling. The first call (with no `syncUrl`) sends `startTime=now` to establish a sync anchor and returns no messages. Subsequent calls pass the `syncState` URL from the previous response and get only new items. Returns `{ messages/conversations, syncUrl }`.
- `getActivity()` is a thin wrapper around `getMessages()` that targets the system conversations (`48:notifications`, `48:mentions`, `48:calllogs`)

### formatting.ts

Pure functions that transform API response dicts into display-ready objects with TypeScript interfaces (`FormattedChat`, `FormattedMessage`, `FormattedPerson`, `FormattedMember`). No I/O, no HTTP calls, no side effects. This makes them trivially testable.

- `stripHtml()` — regex-based HTML tag removal
- `formatTimestamp()` — ISO 8601 to `YYYY-MM-DD HH:MM` (UTC)
- `getConversationDisplayName()` — extracts a name from a conversation object, falling back through topic, sender name, and truncated ID
- `getConversationType()` — maps thread type strings to human labels
- `formatChatList()` — filters out system conversations (ID prefix `48:`), assigns deterministic short IDs, and builds preview strings
- `formatMessage()` — returns `null` for non-displayable messages (system events, call notifications), keeping the filtering logic out of `cli.ts`
- `formatPerson()` — formats a people search result with name, email, MRI, title, department, and company
- `formatMember()` — determines member type from MRI prefix (`8:orgid:` = User, `28:` = Bot)

### cli.ts

Commander.js command definitions. Each command follows the same pattern:

1. Create the API client (triggers token refresh if needed)
2. Resolve input (convert short ID to conversation ID if needed)
3. Call async API functions
4. Format the results
5. Render with cli-table3 (tables for lists) or chalk (coloured text for messages)

The chat index system works by having `teams chats` cache its formatted output to `~/.teams-cli/last_chats.json`. Other commands read this file to resolve short IDs to conversation IDs, so `teams messages abc1` works without the user needing to copy-paste long IDs.

The `watch` command uses an async poll loop, creating a fresh `ApiClient` per iteration because tokens may be refreshed between polls.

## API surfaces

The CLI talks to three separate Microsoft APIs, each requiring its own auth token:

### Chat Service API (chatsvc)

The primary API for conversations, messages, and members.

- Base URL: `https://teams.cloud.microsoft/api/chatsvc/{region}/v1/users/ME`
- Token audience: `ic3.teams.office.com`
- Used by: `chats`, `messages`, `send`, `dm` (thread creation + send), `watch`, `activity`, `members`

### Substrate Search API

Microsoft's unified search service, used here for message search.

- Base URL: `https://substrate.office.com`
- Token audience: `outlook.office.com/search`
- Used by: `search`, `find`, `dm` (people search step)

### Presence API

User presence/status information.

- Base URL: `https://teams.cloud.microsoft/ups/{region}/v1`
- Token audience: `presence.teams.microsoft.com`
- Not yet exposed as a CLI command

## Testing strategy

Tests are organized to match the module structure:

- `formatting.test.ts` — unit tests for all pure formatting functions (27 tests). No mocking needed.
- `auth.test.ts` — tests token persistence, expiry, and JWT parsing using `vi.mock` to redirect file I/O to temp directories (9 tests). Does not test the Playwright login/refresh flows directly.
- `api.test.ts` — tests API functions using mocked `ApiClient.fetch` (6 tests). The chatsvc base URL is patched to a fixed value.
- `cli.test.ts` — tests command registration and structure (2 tests).

The Playwright login flow is not unit-tested since it requires a real browser and user interaction. It's verified manually.

## Data flow example

Here's what happens when a user runs `teams messages a1b2`:

```
cli.ts: messages("a1b2")
  -> resolveConversationId("a1b2")
     -> reads ~/.teams-cli/last_chats.json
     -> returns "19:abc123@thread.v2"
  -> getChatClient()
     -> auth.getToken("ic3") -> reads ~/.teams-cli/tokens.json
        -> if expired: tryRefresh() launches headless browser, re-extracts tokens
     -> returns ApiClient with Bearer token headers
  -> api.getMessages(client, "19:abc123@thread.v2")
     -> GET /conversations/19:abc123@thread.v2/messages
     -> returns list of raw message dicts
  -> formatting.formatMessage() for each message
     -> filters out system events, strips HTML, formats timestamps
  -> chalk + console.log output
```
