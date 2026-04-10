"""HTTP client construction for Teams APIs."""

import sys

import httpx
import typer

from .auth import get_region, get_token
from .config import CLIENT_INFO


def _require_token(name: str, label: str) -> str:
    """Get a token or exit with a helpful error."""
    token = get_token(name)
    if not token:
        typer.echo(
            f"No valid {label} token found. Run `teams login` to authenticate.",
            err=True,
        )
        raise typer.Exit(1)
    return token


def get_chatsvc_base_url() -> str:
    """Build the chatsvc base URL from stored region."""
    region = get_region() or "amer"
    return f"https://teams.cloud.microsoft/api/chatsvc/{region}/v1/users/ME"


def get_chat_client() -> httpx.AsyncClient:
    """Create an HTTP client for the chatsvc API."""
    token = _require_token("ic3", "Teams chat")
    return httpx.AsyncClient(
        headers={
            "Authorization": f"Bearer {token}",
            "x-ms-test-user": "False",
            "x-ms-migration": "True",
            "behavioroverride": "redirectAs404",
            "clientinfo": CLIENT_INFO,
        },
        timeout=30,
    )


def get_search_client() -> httpx.AsyncClient:
    """Create an HTTP client for the Substrate Search API."""
    token = _require_token("search", "search")
    return httpx.AsyncClient(
        base_url="https://substrate.office.com",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "x-client-version": "T2.1",
        },
        timeout=30,
    )


def get_presence_client() -> httpx.AsyncClient:
    """Create an HTTP client for the Presence API."""
    token = _require_token("presence", "presence")
    region = get_region() or "amer"
    return httpx.AsyncClient(
        base_url=f"https://teams.cloud.microsoft/ups/{region}/v1",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        timeout=30,
    )
