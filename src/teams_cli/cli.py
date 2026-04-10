"""Typer CLI application for Microsoft Teams."""

import asyncio
import json
import time

import typer
from rich.console import Console
from rich.table import Table

from .api import (
    create_dm_thread,
    get_activity,
    get_messages,
    get_thread_members,
    list_conversations,
    poll_conversations,
    poll_messages,
    search_messages,
    search_people,
    send_message,
)
from .auth import get_my_mri, login, logout
from .client import get_chat_client, get_search_client
from .config import LAST_CHATS_FILE, DATA_DIR
from .formatting import (
    format_chat_list,
    format_member,
    format_message,
    format_person,
    get_conversation_display_name,
    strip_html,
)

app = typer.Typer(
    name="teams",
    help="CLI for Microsoft Teams – list chats, read and send messages, search, and more.",
    no_args_is_help=True,
)
console = Console()


def _save_chat_index(chats: list[dict]) -> None:
    """Cache formatted chat list so other commands can resolve by index."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    LAST_CHATS_FILE.write_text(json.dumps(chats, indent=2))


def _load_chat_index() -> list[dict]:
    if not LAST_CHATS_FILE.exists():
        return []
    return json.loads(LAST_CHATS_FILE.read_text())


def _resolve_conversation_id(id_or_short: str) -> str:
    """Resolve a short ID (from `teams chats`) or raw conversation ID."""
    # Raw conversation IDs contain ':' (e.g. 19:abc@thread.v2)
    if ":" in id_or_short:
        return id_or_short
    chats = _load_chat_index()
    for c in chats:
        if c.get("short_id") == id_or_short:
            return c["id"]
    typer.echo(
        f"Short ID '{id_or_short}' not found. Run `teams chats` first.", err=True
    )
    raise typer.Exit(1)


# ── Commands ─────────────────────────────────────────────────────────


@app.command("login")
def login_cmd() -> None:
    """Launch a browser to sign in to Microsoft Teams and store auth tokens."""
    try:
        tokens = login()
    except RuntimeError as e:
        typer.echo(str(e), err=True)
        raise typer.Exit(1)
    except TimeoutError:
        typer.echo("Login timed out. Please try again.", err=True)
        raise typer.Exit(1)

    names = [k for k in tokens if k != "region"]
    console.print(f"[green]Logged in.[/green] Tokens saved: {', '.join(names)}")


@app.command("logout")
def logout_cmd() -> None:
    """Remove stored tokens and browser session data."""
    logout()
    console.print("Logged out. Tokens and session data removed.")


@app.command()
def chats(
    limit: int = typer.Option(30, "--limit", "-n", help="Number of chats to list."),
) -> None:
    """List recent conversations."""
    client = get_chat_client()

    async def _run() -> list[dict]:
        async with client:
            return await list_conversations(client, page_size=limit)

    convos = asyncio.run(_run())
    formatted = format_chat_list(convos)
    _save_chat_index(formatted)

    table = Table(title="Recent Chats")
    table.add_column("ID", style="cyan")
    table.add_column("Name", style="bold")
    table.add_column("Type", style="dim")
    table.add_column("Last Message")
    table.add_column("Time", style="green")

    for c in formatted:
        table.add_row(
            c["short_id"],
            c["name"],
            c["type"],
            c["preview"][:60],
            c["time"],
        )
    console.print(table)


@app.command()
def messages(
    chat: str = typer.Argument(
        help="Chat index (from `teams chats`) or conversation ID."
    ),
    limit: int = typer.Option(20, "--limit", "-n", help="Number of messages to fetch."),
) -> None:
    """Read messages from a conversation."""
    conv_id = _resolve_conversation_id(chat)
    client = get_chat_client()

    async def _run() -> list[dict]:
        async with client:
            return await get_messages(client, conv_id, page_size=limit)

    raw = asyncio.run(_run())
    msgs = [m for m in (format_message(r) for r in raw) if m]

    if not msgs:
        typer.echo("No displayable messages.")
        raise typer.Exit()

    # Show oldest first
    for m in reversed(msgs):
        console.print(f"[green]{m['time']}[/green] [bold]{m['sender']}[/bold]")
        console.print(f"  {m['body']}")
        console.print()


@app.command()
def send(
    chat: str = typer.Argument(help="Chat index or conversation ID."),
    message: str = typer.Argument(help="Message text to send."),
) -> None:
    """Send a message to a conversation."""
    conv_id = _resolve_conversation_id(chat)
    client = get_chat_client()

    async def _run() -> dict:
        async with client:
            return await send_message(client, conv_id, message)

    result = asyncio.run(_run())
    console.print(
        f"[green]Sent.[/green] Arrival: {result.get('OriginalArrivalTime', 'ok')}"
    )


@app.command()
def search(
    query: str = typer.Argument(help="Search query."),
    limit: int = typer.Option(25, "--limit", "-n", help="Max results."),
) -> None:
    """Search messages across all conversations."""
    client = get_search_client()

    async def _run() -> dict:
        async with client:
            return await search_messages(client, query, size=limit)

    data = asyncio.run(_run())

    results = []
    for entity_set in data.get("entitySets", []):
        for result in entity_set.get("resultSets", []):
            for hit in result.get("results", []):
                body = strip_html(hit.get("preview", ""))
                sender = hit.get("extensions", {}).get(
                    "Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String",
                    "",
                )
                results.append({"sender": sender, "preview": body})

    if not results:
        typer.echo("No results found.")
        raise typer.Exit()

    for i, r in enumerate(results, 1):
        sender_label = r["sender"] or "Unknown"
        console.print(f"[cyan]{i}.[/cyan] [bold]{sender_label}[/bold]")
        console.print(f"   {r['preview'][:120]}")
        console.print()


@app.command()
def activity(
    feed: str = typer.Option(
        "notifications",
        "--feed",
        "-f",
        help="Feed: notifications, mentions, or calllogs.",
    ),
    limit: int = typer.Option(20, "--limit", "-n", help="Number of items."),
) -> None:
    """Show the activity feed (notifications, mentions, or call logs)."""
    client = get_chat_client()

    async def _run() -> list[dict]:
        async with client:
            return await get_activity(client, feed=feed, page_size=limit)

    raw = asyncio.run(_run())
    msgs = [m for m in (format_message(r) for r in raw) if m]

    if not msgs:
        typer.echo("No activity items.")
        raise typer.Exit()

    for m in reversed(msgs):
        console.print(f"[green]{m['time']}[/green] [bold]{m['sender']}[/bold]")
        console.print(f"  {m['body']}")
        console.print()


@app.command()
def find(
    query: str = typer.Argument(help="Name or email to search for."),
    limit: int = typer.Option(10, "--limit", "-n", help="Max results."),
) -> None:
    """Search for people by name or email."""
    client = get_search_client()

    async def _run() -> list[dict]:
        async with client:
            return await search_people(client, query, size=limit)

    results = asyncio.run(_run())
    if not results:
        typer.echo("No people found.")
        raise typer.Exit()

    table = Table(title="People")
    table.add_column("#", style="cyan")
    table.add_column("Name", style="bold")
    table.add_column("Email")
    table.add_column("Title", style="dim")
    table.add_column("MRI", style="dim")

    for i, person in enumerate(results, 1):
        p = format_person(person)
        table.add_row(str(i), p["name"], p["email"], p["title"], p["mri"])

    console.print(table)


@app.command()
def dm(
    user: str = typer.Argument(
        help="User name to search, email, or MRI (8:orgid:...)."
    ),
    message: str = typer.Argument(help="Message text to send."),
) -> None:
    """Send a direct message to a user by name, email, or MRI."""
    my_mri = get_my_mri()
    if not my_mri:
        typer.echo("Could not determine your user ID. Run `teams login`.", err=True)
        raise typer.Exit(1)

    # If it looks like an MRI, use it directly
    if user.startswith("8:orgid:"):
        their_mri = user
    else:
        # Search for the user
        search_client = get_search_client()

        async def _search() -> list[dict]:
            async with search_client:
                return await search_people(search_client, user, size=5)

        results = asyncio.run(_search())
        if not results:
            typer.echo(f"No user found matching '{user}'.", err=True)
            raise typer.Exit(1)

        if len(results) == 1:
            their_mri = results[0]["MRI"]
            name = results[0].get("DisplayName", user)
        else:
            # Show matches and pick the first
            console.print(f"[yellow]Multiple matches for '{user}':[/yellow]")
            for i, r in enumerate(results, 1):
                emails = r.get("EmailAddresses", [])
                email = emails[0] if emails else ""
                console.print(f"  {i}. {r.get('DisplayName', '?')} ({email})")
            console.print("[green]Using first match.[/green]")
            their_mri = results[0]["MRI"]
            name = results[0].get("DisplayName", user)

        console.print(f"Sending DM to [bold]{name}[/bold]...")

    # Create or get the 1:1 thread, then send the message
    chat_client = get_chat_client()

    async def _send() -> dict:
        async with chat_client:
            conv_id = await create_dm_thread(chat_client, my_mri, their_mri)
            return await send_message(chat_client, conv_id, message)

    result = asyncio.run(_send())
    console.print(
        f"[green]Sent.[/green] Arrival: {result.get('OriginalArrivalTime', 'ok')}"
    )


@app.command()
def watch(
    chat: str = typer.Argument(
        None, help="Chat short ID or conversation ID. Omit to watch all chats."
    ),
    interval: int = typer.Option(
        3, "--interval", "-i", help="Poll interval in seconds."
    ),
) -> None:
    """Watch a chat (or all chats) for new messages in real-time.

    Press Ctrl+C to stop.
    """
    client = get_chat_client()

    if chat:
        conv_id = _resolve_conversation_id(chat)
        _watch_chat(client, conv_id, interval)
    else:
        _watch_all(client, interval)


def _watch_chat(client, conv_id: str, interval: int) -> None:
    """Poll a single conversation for new messages."""

    async def _init() -> str:
        async with client:
            _, sync_url = await poll_messages(client, conv_id)
            return sync_url

    sync_url = asyncio.run(_init())
    console.print(
        f"[dim]Watching for new messages (poll every {interval}s). Ctrl+C to stop.[/dim]"
    )

    try:
        while True:
            time.sleep(interval)
            poll_client = get_chat_client()

            async def _poll() -> tuple[list[dict], str]:
                async with poll_client:
                    return await poll_messages(poll_client, conv_id, sync_url=sync_url)

            raw, sync_url = asyncio.run(_poll())
            msgs = [m for m in (format_message(r) for r in raw) if m]
            for m in msgs:
                console.print(f"[green]{m['time']}[/green] [bold]{m['sender']}[/bold]")
                console.print(f"  {m['body']}")
                console.print()
    except KeyboardInterrupt:
        console.print("\n[dim]Stopped watching.[/dim]")


def _watch_all(client, interval: int) -> None:
    """Poll all conversations for new activity."""
    chats = _load_chat_index()
    chat_names = {c["id"]: c["name"] for c in chats}

    async def _init() -> str:
        async with client:
            _, sync_url = await poll_conversations(client)
            return sync_url

    sync_url = asyncio.run(_init())
    console.print(
        f"[dim]Watching all chats (poll every {interval}s). Ctrl+C to stop.[/dim]"
    )

    try:
        while True:
            time.sleep(interval)
            poll_client = get_chat_client()

            async def _poll() -> tuple[list[dict], str]:
                async with poll_client:
                    return await poll_conversations(poll_client, sync_url=sync_url)

            convos, sync_url = asyncio.run(_poll())
            for conv in convos:
                conv_id = conv.get("id", "")
                if conv_id.startswith("48:"):
                    continue
                last_msg = conv.get("lastMessage", {})
                msg = format_message(last_msg)
                if not msg:
                    continue
                name = chat_names.get(conv_id, get_conversation_display_name(conv))
                console.print(
                    f"[cyan]{name}[/cyan] "
                    f"[green]{msg['time']}[/green] [bold]{msg['sender']}[/bold]"
                )
                console.print(f"  {msg['body']}")
                console.print()
    except KeyboardInterrupt:
        console.print("\n[dim]Stopped watching.[/dim]")


@app.command()
def members(
    chat: str = typer.Argument(help="Chat index or conversation ID."),
) -> None:
    """List members of a conversation."""
    conv_id = _resolve_conversation_id(chat)
    client = get_chat_client()

    async def _run() -> list[dict]:
        async with client:
            return await get_thread_members(client, conv_id)

    raw = asyncio.run(_run())
    formatted = [format_member(m) for m in raw]

    table = Table(title="Members")
    table.add_column("Name", style="bold")
    table.add_column("Role")
    table.add_column("Type", style="dim")

    for m in formatted:
        table.add_row(m["name"], m["role"], m["type"])

    console.print(table)
