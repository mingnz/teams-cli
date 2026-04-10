# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```sh
uv sync                        # Install dependencies
uv run pytest                  # Run all tests
uv run pytest -v               # Run tests with verbose output
uv run pytest tests/test_formatting.py::TestStripHtml::test_removes_tags  # Run a single test
uv run teams --help            # Show CLI help
uv run teams <command> --help  # Show help for a specific command
```

## Architecture

This is a Python CLI for Microsoft Teams that talks directly to the internal Teams Chat Service API (chatsvc) — the same HTTP API the Teams web client uses. It does **not** use Microsoft Graph.

### Module layers (dependencies flow downward only)

- **cli.py** — Typer commands. Each command runs an async API call via `asyncio.run()`, formats results, renders with Rich. The `chats` command caches output to `~/.teams-cli/last_chats.json` so other commands can accept numeric chat indices instead of conversation IDs.
- **formatting.py** — Pure functions transforming API response dicts into display-ready dicts. No I/O.
- **api.py** — Async functions making HTTP calls. All take `httpx.AsyncClient` as first arg (for testability with pytest-httpx).
- **client.py** — Factory functions creating pre-configured `httpx.AsyncClient` instances for three API surfaces (chat, search, presence), each with its own token.
- **auth.py** — Token persistence (`~/.teams-cli/tokens.json`) and Playwright-based browser login flow that extracts tokens from localStorage.
- **config.py** — Constants only, no imports from other project modules.

### Three API surfaces

| API | Token name | Base URL pattern |
|-----|-----------|-----------------|
| Chat Service (chatsvc) | `ic3` | `teams.cloud.microsoft/api/chatsvc/{region}/v1/users/ME` |
| Substrate Search | `search` | `substrate.office.com` |
| Presence | `presence` | `teams.cloud.microsoft/ups/{region}/v1` |

### Key design details

- `send_message()` wraps content in `<p>` tags and generates a random 19-digit `clientmessageid`
- `create_dm_thread()` uses `POST /threads` (without `/users/ME`) with `uniquerosterthread: true` for idempotent 1:1 chat creation
- `search_people()` uses Substrate Suggestions API (`/search/api/v1/suggestions?scenario=peoplepicker.newChat`), same endpoint as the Teams web client people picker
- `get_my_mri()` extracts the user's MRI from the ic3 JWT token's `oid` claim — no extra API call needed
- `poll_messages()` and `poll_conversations()` use **delta sync** via `startTime` + `syncState` — first call anchors at current time (returns nothing), subsequent calls with the returned `syncState` URL return only new items
- The `watch` command creates a new `httpx.AsyncClient` per poll iteration (tokens may refresh between polls)
- Members are at `/threads/{id}`, not `/conversations/{id}/members`
- Activity feed reuses `get_messages()` with system conversation IDs (`48:notifications`, `48:mentions`, `48:calllogs`)
- System conversations (ID prefix `48:`) are filtered out of chat listings
- Login uses a **persistent browser profile** (`~/.teams-cli/browser-profile/`) so session cookies survive for silent token refresh
- Expired tokens are automatically refreshed via headless Playwright before each command — no user interaction unless the session itself has expired
- HTTP clients are created **outside** `asyncio.run()` because Playwright's sync API conflicts with an existing event loop

## Conventions

- Use [Conventional Commits](https://www.conventionalcommits.org/) for all commit messages (e.g. `feat:`, `fix:`, `refactor:`, `docs:`, `test:`).
- PR titles must also follow the Conventional Commits format.
- Run `uv run ruff check --fix . && uv run ruff format .` before opening PRs.

## Maintenance

When making changes to features or architecture, update the relevant docs in the same PR:
- **CLAUDE.md** — key guidance for maintainers (including AI agents): design decisions, gotchas, commands, conventions
- **docs/ARCHITECTURE.md** — module responsibilities, data flow, testing strategy
- **README.md** — user-facing usage examples and install instructions
