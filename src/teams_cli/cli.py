"""Typer CLI application for Microsoft Teams."""

import asyncio
import json

import typer
from rich.console import Console
from rich.table import Table

from .api import (
    get_activity,
    get_messages,
    get_thread_members,
    list_conversations,
    search_messages,
    send_message,
)
from .auth import login
from .client import get_chat_client, get_search_client
from .config import LAST_CHATS_FILE, DATA_DIR
from .formatting import (
    format_chat_list,
    format_member,
    format_message,
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


def _resolve_conversation_id(id_or_index: str) -> str:
    """Resolve an index number (from `teams chats`) or raw conversation ID."""
    if id_or_index.isdigit():
        idx = int(id_or_index)
        chats = _load_chat_index()
        for c in chats:
            if c["index"] == idx:
                return c["id"]
        typer.echo(f"Index {idx} not found. Run `teams chats` first.", err=True)
        raise typer.Exit(1)
    return id_or_index


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


@app.command()
def chats(
    limit: int = typer.Option(30, "--limit", "-n", help="Number of chats to list."),
) -> None:
    """List recent conversations."""

    async def _run() -> list[dict]:
        async with get_chat_client() as client:
            return await list_conversations(client, page_size=limit)

    convos = asyncio.run(_run())
    formatted = format_chat_list(convos)
    _save_chat_index(formatted)

    table = Table(title="Recent Chats")
    table.add_column("#", style="cyan", justify="right")
    table.add_column("Name", style="bold")
    table.add_column("Type", style="dim")
    table.add_column("Last Message")
    table.add_column("Time", style="green")

    for c in formatted:
        table.add_row(
            str(c["index"]),
            c["name"],
            c["type"],
            c["preview"][:60],
            c["time"],
        )
    console.print(table)


@app.command()
def messages(
    chat: str = typer.Argument(help="Chat index (from `teams chats`) or conversation ID."),
    limit: int = typer.Option(20, "--limit", "-n", help="Number of messages to fetch."),
) -> None:
    """Read messages from a conversation."""
    conv_id = _resolve_conversation_id(chat)

    async def _run() -> list[dict]:
        async with get_chat_client() as client:
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

    async def _run() -> dict:
        async with get_chat_client() as client:
            return await send_message(client, conv_id, message)

    result = asyncio.run(_run())
    console.print(f"[green]Sent.[/green] Arrival: {result.get('OriginalArrivalTime', 'ok')}")


@app.command()
def search(
    query: str = typer.Argument(help="Search query."),
    limit: int = typer.Option(25, "--limit", "-n", help="Max results."),
) -> None:
    """Search messages across all conversations."""

    async def _run() -> dict:
        async with get_search_client() as client:
            return await search_messages(client, query, size=limit)

    data = asyncio.run(_run())

    results = []
    for entity_set in data.get("entitySets", []):
        for result in entity_set.get("resultSets", []):
            for hit in result.get("results", []):
                body = strip_html(hit.get("preview", ""))
                sender = hit.get("extensions", {}).get(
                    "Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String", ""
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
    feed: str = typer.Option("notifications", "--feed", "-f", help="Feed: notifications, mentions, or calllogs."),
    limit: int = typer.Option(20, "--limit", "-n", help="Number of items."),
) -> None:
    """Show the activity feed (notifications, mentions, or call logs)."""

    async def _run() -> list[dict]:
        async with get_chat_client() as client:
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
def members(
    chat: str = typer.Argument(help="Chat index or conversation ID."),
) -> None:
    """List members of a conversation."""
    conv_id = _resolve_conversation_id(chat)

    async def _run() -> list[dict]:
        async with get_chat_client() as client:
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
