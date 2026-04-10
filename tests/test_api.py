"""Tests for API module."""

import pytest
import httpx
from pytest_httpx import HTTPXMock

from teams_cli.api import (
    get_activity,
    get_messages,
    get_thread_members,
    list_conversations,
    send_message,
)


BASE = "https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME"


@pytest.fixture(autouse=True)
def _patch_region(monkeypatch):
    monkeypatch.setattr("teams_cli.api.get_chatsvc_base_url", lambda: BASE)


@pytest.fixture()
def client():
    return httpx.AsyncClient()


class TestListConversations:
    @pytest.mark.asyncio
    async def test_returns_conversations(self, client, httpx_mock: HTTPXMock):
        httpx_mock.add_response(
            url=f"{BASE}/conversations?view=msnp24Equivalent&pageSize=10",
            json={"conversations": [{"id": "19:abc"}]},
        )
        result = await list_conversations(client, page_size=10)
        assert len(result) == 1
        assert result[0]["id"] == "19:abc"


class TestGetMessages:
    @pytest.mark.asyncio
    async def test_returns_messages(self, client, httpx_mock: HTTPXMock):
        httpx_mock.add_response(json={"messages": [{"id": "1", "content": "hi"}]})
        result = await get_messages(client, "19:abc", page_size=5)
        assert len(result) == 1


class TestSendMessage:
    @pytest.mark.asyncio
    async def test_sends(self, client, httpx_mock: HTTPXMock):
        httpx_mock.add_response(json={"OriginalArrivalTime": "123"})
        result = await send_message(client, "19:abc", "hello")
        assert "OriginalArrivalTime" in result

    @pytest.mark.asyncio
    async def test_special_characters_in_message(self, client, httpx_mock: HTTPXMock):
        """Verify messages with quotes and braces produce valid JSON."""
        httpx_mock.add_response(json={"OriginalArrivalTime": "456"})
        result = await send_message(client, "19:abc", 'He said "hello" & {test}')
        assert "OriginalArrivalTime" in result
        request = httpx_mock.get_requests()[0]
        import json
        body = json.loads(request.content)
        assert body["content"] == '<p>He said "hello" & {test}</p>'
        assert body["messagetype"] == "RichText/Html"


class TestGetThreadMembers:
    @pytest.mark.asyncio
    async def test_returns_members(self, client, httpx_mock: HTTPXMock):
        httpx_mock.add_response(
            json={"members": [{"id": "8:orgid:user1"}]},
        )
        result = await get_thread_members(client, "19:abc")
        assert len(result) == 1


class TestGetActivity:
    @pytest.mark.asyncio
    async def test_returns_activity(self, client, httpx_mock: HTTPXMock):
        httpx_mock.add_response(json={"messages": [{"id": "n1"}]})
        result = await get_activity(client, feed="notifications", page_size=5)
        assert len(result) == 1
