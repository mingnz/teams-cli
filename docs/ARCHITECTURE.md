# Architecture

## Overview

teams-cli is a Python CLI that talks directly to the internal Microsoft Teams Chat Service API — the same HTTP API that the Teams web client uses. It bypasses the Microsoft Graph SDK entirely, which avoids the need for Azure app registrations or admin-granted API permissions.

Authentication is handled by launching a real browser via Playwright, letting the user sign in normally (including MFA), then extracting tokens from the browser's `localStorage`.

## Layered design

The codebase follows a strict layered architecture where each module has a single responsibility and dependencies only flow downward:

```
cli.py          Commands (user-facing)
  |
  +-- formatting.py   Pure display logic (no I/O)
  +-- api.py          HTTP request/response logic
        |
        +-- client.py     HTTP client construction
              |
              +-- auth.py      Token storage and browser login
              +-- config.py    Constants (no dependencies)
```

### config.py

All constants live here: file paths, token audience strings, API headers, default page sizes, and the Teams login URL. No imports from other project modules — this is the leaf of the dependency tree.

Key values:
- `DATA_DIR` — `~/.teams-cli/`, where tokens and caches are stored
- `TOKEN_AUDIENCES` — maps token names (`ic3`, `search`, `presence`) to the localStorage key patterns used to find them
- `CLIENT_INFO` — the `clientinfo` header that the chatsvc API requires

### auth.py

Handles token persistence, automatic refresh, and the Playwright login flow.

- `save_tokens()` / `load_tokens()` — JSON round-trip to `~/.teams-cli/tokens.json`
- `is_expired()` / `get_token()` — check token expiry, attempt silent refresh if expired, and retrieve by name
- `get_region()` — returns the user's region code (e.g. `au`, `amer`) extracted during login
- `get_my_mri()` — extracts the current user's MRI (`8:orgid:{oid}`) by decoding the `oid` claim from the ic3 JWT token. Used by the `dm` command to construct thread creation requests.
- `_try_refresh()` — silently refreshes expired tokens by launching a headless Chromium with the persistent browser profile (`~/.teams-cli/browser-profile/`). MSAL re-acquires tokens using the stored session cookies — no user interaction needed. Returns `True` on success, `False` if the session itself has expired.
- `login()` — the interactive login flow:
  1. Launches Chromium (headless=False) with a persistent browser profile so session cookies are preserved
  2. Navigates to `teams.cloud.microsoft`
  3. Polls `localStorage` every 2 seconds for up to 5 minutes until the IC3 token appears
  4. Extracts all three tokens (ic3, search, presence) and the region from `localStorage`
  5. Saves everything to disk
- `logout()` — removes stored tokens and the browser profile directory

### client.py

Factory functions that create pre-configured `httpx.AsyncClient` instances. Each client targets a different API surface:

| Factory | API | Token | Base URL |
|---------|-----|-------|----------|
| `get_chat_client()` | Chat Service (chatsvc) | `ic3` | (per-request, includes region) |
| `get_search_client()` | Substrate Search | `search` | `substrate.office.com` |
| `get_presence_client()` | Presence | `presence` | `teams.cloud.microsoft/ups/{region}/v1` |

All clients set a 30-second timeout. If the required token is missing or expired, the client factory prints an error and exits.

### api.py

Async functions that make HTTP calls. Each function takes an `httpx.AsyncClient` as its first argument, making them easy to test with `pytest-httpx`. No formatting or display logic lives here.

Key design decisions:
- `send_message()` wraps content in `<p>` tags and generates a random 19-digit `clientmessageid`, matching the Teams web client's behaviour
- `get_thread_members()` uses the `/threads/{id}` endpoint rather than `/conversations/{id}/members` (which returns 404)
- `search_messages()` builds the full Substrate Search API request body including correlation IDs and dimension metadata
- `search_people()` uses the Substrate Suggestions API (`/search/api/v1/suggestions?scenario=peoplepicker.newChat`) — the same people picker endpoint the Teams web client uses when composing a new chat
- `create_dm_thread()` creates (or retrieves) a 1:1 chat thread via `POST /threads` with `uniquerosterthread: true`, which makes it idempotent — it returns the existing thread if one already exists
- `poll_messages()` and `poll_conversations()` implement delta sync polling. The first call (with `sync_url=None`) sends `startTime=now` to establish a sync anchor and returns no messages. Subsequent calls pass the `syncState` URL from the previous response and get only new items. Returns `(items, next_sync_url)`.
- `get_activity()` is a thin wrapper around `get_messages()` that targets the system conversations (`48:notifications`, `48:mentions`, `48:calllogs`)

### formatting.py

Pure functions that transform API response dicts into display-ready dicts. No I/O, no HTTP calls, no side effects. This makes them trivially testable.

- `strip_html()` — regex-based HTML tag removal
- `format_timestamp()` — ISO 8601 to `YYYY-MM-DD HH:MM`
- `get_conversation_display_name()` — extracts a name from a conversation object, falling back through topic, sender name, and truncated ID
- `get_conversation_type()` — maps thread type strings to human labels
- `format_chat_list()` — filters out system conversations (ID prefix `48:`), assigns sequential index numbers, and builds preview strings
- `format_message()` — returns `None` for non-displayable messages (system events, call notifications), keeping the filtering logic out of `cli.py`
- `format_person()` — formats a people search result into a display dict with name, email, MRI, title, department, and company
- `format_member()` — determines member type from MRI prefix (`8:orgid:` = User, `28:` = Bot)

### cli.py

Typer command definitions. Each command follows the same pattern:

1. Create the HTTP client (triggers token refresh if needed — must happen before `asyncio.run()` since Playwright's sync API can't run inside an existing event loop)
2. Resolve input (convert chat index to conversation ID if needed)
3. Run an async API call via `asyncio.run()`
4. Format the results
5. Render with Rich (tables for lists, plain text for messages)

The chat index system works by having `teams chats` cache its formatted output to `~/.teams-cli/last_chats.json`. Other commands read this file to resolve short IDs to conversation IDs, so `teams messages abc1` works without the user needing to copy-paste long IDs.

The `watch` command is the exception — it uses a poll loop instead of a single `asyncio.run()`. It creates a fresh `httpx.AsyncClient` per poll iteration because tokens may be refreshed between polls. The loop runs synchronously with `time.sleep()` between polls, calling `asyncio.run()` for each HTTP request.

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

- `test_formatting.py` — unit tests for all pure formatting functions (25 tests). No mocking needed.
- `test_auth.py` — tests token persistence, expiry, and refresh logic using `monkeypatch` to redirect file I/O to `tmp_path` and stub `_try_refresh` (8 tests). Does not test the Playwright login/refresh flows directly.
- `test_api.py` — tests API functions using `pytest-httpx` to mock HTTP responses (5 tests). The chatsvc base URL is patched to a fixed value.
- `test_cli.py` — tests command registration and the index resolution helper using Typer's `CliRunner` (4 tests).

The Playwright login flow is not unit-tested since it requires a real browser and user interaction. It's verified manually.

## Data flow example

Here's what happens when a user runs `teams messages 3`:

```
cli.py: messages("3")
  -> _resolve_conversation_id("3")
     -> reads ~/.teams-cli/last_chats.json
     -> returns "19:abc123@thread.v2"
  -> get_chat_client()  [before asyncio.run()]
     -> auth.get_token("ic3") -> reads ~/.teams-cli/tokens.json
        -> if expired: _try_refresh() launches headless browser, re-extracts tokens
     -> returns httpx.AsyncClient with Bearer token
  -> api.get_messages(client, "19:abc123@thread.v2")
     -> GET /conversations/19:abc123@thread.v2/messages
     -> returns list of raw message dicts
  -> formatting.format_message() for each message
     -> filters out system events, strips HTML, formats timestamps
  -> Rich console output
```
