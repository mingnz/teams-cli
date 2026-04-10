---
name: teams-cli
description: >
  Interact with Microsoft Teams on behalf of the user — read chats, send messages, search conversations,
  find people, send DMs, watch for new messages, and view activity. Use this skill whenever the user
  mentions Teams, wants to check or read messages, send a message to someone, look up a coworker,
  search chat history, monitor conversations, or anything involving Microsoft Teams communication.
  Also trigger when the user asks to "message someone", "ping someone", "check what X said",
  "find someone at work", or "reply to the chat about Y" — even if they don't explicitly say "Teams".
---

# Microsoft Teams CLI

You have access to a `teams` CLI tool that talks to Microsoft Teams.

## Setup

Before using any commands, check if the CLI is installed and the user is authenticated:

1. **Check prerequisites:** Run `uv --version`. If `uv` is not found, ask the user if they'd like to install it. If yes, install via `curl -LsSf https://astral.sh/uv/install.sh | sh` (macOS/Linux) or `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"` (Windows). See https://docs.astral.sh/uv/getting-started/installation for other options.
2. **Check if installed:** Run `teams --help` (or `uv run teams --help`). If the command is not found, install it:
   ```bash
   git clone https://github.com/mingnz/teams-cli.git ~/.teams-cli-install
   cd ~/.teams-cli-install
   uv sync
   uv run playwright install chromium
   uv tool install -e .
   ```
3. **Check if authenticated:** Try a command like `teams chats`. If it fails with a token error, run `teams login` — this launches a browser for the user to sign in (including MFA). Tell the user a browser window has opened and to complete the sign-in. Wait for the command to finish. Only if `teams login` itself fails should you ask the user to run it manually.

Tokens refresh automatically via a headless browser between commands. If the session has fully expired, run `teams login` again.

## Commands

### List chats
```bash
teams chats                # List 30 most recent chats
teams chats -n 50          # List 50 chats
```
Returns a table with short IDs (like `a1b2`), chat name, type, last message preview, and timestamp. The short IDs are used by other commands to reference chats — always run `teams chats` first if you need to find a conversation.

### Read messages
```bash
teams messages <chat_id>        # Read 20 most recent messages
teams messages <chat_id> -n 40  # Read 40 messages
```
`<chat_id>` is either a short ID from `teams chats` or a full conversation ID (e.g. `19:abc@thread.v2`). Messages are shown oldest-first.

### Send a message
```bash
teams send <chat_id> "Your message here"
```
Sends a message to an existing conversation. Always confirm with the user before sending.

### Search messages
```bash
teams search "quarterly report"       # Search across all chats
teams search "budget review" -n 10    # Limit results
```
Searches message content across all conversations.

### View activity
```bash
teams activity                        # Notifications feed
teams activity --feed mentions        # Messages where user was @mentioned
teams activity --feed calllogs        # Call history
```

### Find people
```bash
teams find "Jane Smith"       # Search by name
teams find "jane@company.com" # Search by email
```
Returns a table with name, email, title, and MRI (internal user ID).

### Send a direct message
```bash
teams dm "Jane Smith" "Hey, quick question about the project"
teams dm "jane@company.com" "Hello!"
teams dm "8:orgid:00000000-0000-..." "Hello via MRI"
```
Searches for the user by name or email, creates a 1:1 chat if one doesn't exist, and sends the message. If multiple matches are found, uses the first result.

### List members
```bash
teams members <chat_id>
```
Shows the members of a conversation with their names, roles, and type (User vs Bot).

### Watch for new messages
```bash
teams watch <chat_id>              # Watch a specific chat
teams watch                        # Watch all chats
teams watch <chat_id> -i 5         # Poll every 5 seconds
```
Polls for new messages in real-time. Runs until interrupted with Ctrl+C. Because this blocks the terminal, use it with a timeout or run it in the background when you need to monitor briefly.

## Typical workflows

### "Check my Teams messages"
1. Run `teams chats` to list recent conversations
2. Look for chats with recent activity
3. Run `teams messages <short_id>` on the relevant ones
4. Summarize what you found for the user

### "Reply to X about Y"
1. Run `teams chats` to find the conversation
2. Run `teams messages <short_id>` to read recent context
3. Draft a reply and confirm with the user before sending
4. Run `teams send <short_id> "message"`

### "Message someone I haven't chatted with"
1. Run `teams find "their name"` to look them up
2. Run `teams dm "their name" "message"` to create a 1:1 chat and send

### "What did someone say about X?"
1. Run `teams search "X"` to find relevant messages
2. Summarize the results

## Important guidelines

- **Always confirm before sending.** Never send a message without the user's explicit approval. Show them the draft and the recipient first.
- **Run `teams chats` first** when you need to reference a conversation — the short IDs are only valid after a fresh `teams chats` call.
- **Quote messages faithfully.** When relaying what someone said, use their exact words rather than paraphrasing.
- **Handle auth errors proactively.** If a command fails with a token error, run `teams login` yourself and tell the user a browser window has opened for them to sign in. Only ask them to run it manually if `teams login` itself errors out. If a short ID isn't found, run `teams chats` to refresh the index.
- **Watch is blocking.** The `watch` command runs until Ctrl+C. When using it programmatically, set a reasonable timeout so it doesn't hang forever.
