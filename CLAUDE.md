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

When making major changes (new modules, new API surfaces, new commands, or architectural shifts), update this file to keep it accurate.
