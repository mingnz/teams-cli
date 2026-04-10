# teams-cli

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
[![uv](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/uv/main/assets/badge/v0.json)](https://github.com/astral-sh/uv)

A command-line interface for Microsoft Teams. List chats, read and send messages, search, view activity, and more — all from your terminal.

Uses the internal Teams Chat Service API (the same API the Teams web client uses), authenticated via a Playwright-based browser login flow. No Entra ID app registration or tenant-level admin setup required — just sign in with your browser.

### AI agent integration

This CLI can be used as a tool for AI agents (e.g. Claude Code, Codex, GitHub Copilot) to access Microsoft Teams on your behalf — without needing a direct Microsoft integration. Most Teams integrations require Entra ID app registrations and tenant admin approval. This doesn't.

## Disclaimer

This project is not affiliated with, endorsed by, or associated with Microsoft. It uses undocumented internal APIs that Microsoft can change or restrict at any time without notice. Use at your own risk — this tool may break unexpectedly.

## Install

Requires Python 3.10+ and [uv](https://docs.astral.sh/uv/).

```sh
# Clone and install
git clone https://github.com/mingnz/teams-cli.git
cd teams-cli
uv sync
uv run playwright install chromium

# Or install globally
uv tool install -e .
```

## Authentication

Log in by launching a browser session. Sign in as you normally would (including MFA), and tokens are captured automatically from the browser:

```sh
teams login
```

Tokens are stored in `~/.teams-cli/tokens.json` and expire after the session lifetime set by your organisation (typically 1 hour). Re-run `teams login` when they expire.

## Usage

```sh
# List recent chats
teams chats

# Read messages from chat #3 (index from `teams chats`)
teams messages 3

# Read messages by conversation ID
teams messages "19:abc123@thread.v2"

# Send a message
teams send 3 "Hello from the CLI"

# Search across all conversations
teams search "quarterly report"

# View activity feed
teams activity
teams activity --feed mentions
teams activity --feed calllogs

# List members of a chat
teams members 3
```

### Options

Most commands accept `--limit` / `-n` to control how many results to fetch:

```sh
teams chats --limit 50
teams messages 3 --limit 40
teams search "budget" --limit 10
```

Run `teams --help` or `teams <command> --help` for full details.

## Development

```sh
# Install dev dependencies
uv sync

# Run tests
uv run pytest

# Run tests with verbose output
uv run pytest -v
```

## How it works

1. `teams login` opens Chromium via Playwright, navigates to Teams, and waits for you to complete sign-in
2. Auth tokens are extracted from the browser's `localStorage` (three tokens: chat, search, presence) along with your region
3. CLI commands use these tokens to call the Teams Chat Service API (`teams.cloud.microsoft/api/chatsvc/`) and the Substrate Search API (`substrate.office.com`) directly via `httpx`
4. The `teams chats` command caches the conversation list locally so you can reference chats by index number in subsequent commands

## Project structure

```
src/teams_cli/
  config.py       # Constants and paths
  auth.py         # Token management and Playwright login
  client.py       # HTTP client factories
  api.py          # Async API functions
  formatting.py   # Output formatting
  cli.py          # Typer command definitions
tests/
  test_auth.py
  test_api.py
  test_cli.py
  test_formatting.py
```
