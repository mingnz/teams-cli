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

### Three API surfaces

| API | Token name | Base URL pattern |
|-----|-----------|-----------------|
| Chat Service (chatsvc) | `ic3` | `teams.cloud.microsoft/api/chatsvc/{region}/v1/users/ME` |
| Substrate Search | `search` | `substrate.office.com` |
| Presence | `presence` | `teams.cloud.microsoft/ups/{region}/v1` |

### Key design details

- `sendMessage()` wraps content in `<p>` tags and generates a random 19-digit `clientmessageid` using BigInt
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

## Conventions

- Use [Conventional Commits](https://www.conventionalcommits.org/) for all commit messages (e.g. `feat:`, `fix:`, `refactor:`, `docs:`, `test:`).
- PR titles must also follow the Conventional Commits format.
- **Never commit PII** (names, emails, user IDs, conversation content) to the repo — including in test fixtures, PR descriptions, commit messages, and comments.

## Maintenance

When making changes to features or architecture, update the relevant docs in the same PR:
- **CLAUDE.md** — key guidance for maintainers (including AI agents): design decisions, gotchas, commands, conventions
- **docs/ARCHITECTURE.md** — module responsibilities, data flow, testing strategy
- **README.md** — user-facing usage examples and install instructions
