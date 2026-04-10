"""Teams API functions."""

import random
import uuid

import httpx

from .client import get_chatsvc_base_url


async def list_conversations(
    client: httpx.AsyncClient, page_size: int = 30
) -> list[dict]:
    """List recent conversations."""
    base = get_chatsvc_base_url()
    resp = await client.get(
        f"{base}/conversations",
        params={"view": "msnp24Equivalent", "pageSize": str(page_size)},
    )
    resp.raise_for_status()
    return resp.json().get("conversations", [])


async def get_conversation(
    client: httpx.AsyncClient, conversation_id: str
) -> dict:
    """Get a single conversation's details."""
    base = get_chatsvc_base_url()
    resp = await client.get(
        f"{base}/conversations/{conversation_id}",
        params={"view": "msnp24Equivalent"},
    )
    resp.raise_for_status()
    return resp.json()


async def get_messages(
    client: httpx.AsyncClient, conversation_id: str, page_size: int = 20
) -> list[dict]:
    """Get messages from a conversation."""
    base = get_chatsvc_base_url()
    resp = await client.get(
        f"{base}/conversations/{conversation_id}/messages",
        params={
            "view": "msnp24Equivalent|supportsMessageProperties",
            "pageSize": str(page_size),
        },
    )
    resp.raise_for_status()
    return resp.json().get("messages", [])


async def send_message(
    client: httpx.AsyncClient, conversation_id: str, content: str
) -> dict:
    """Send a message to a conversation.

    Returns dict with OriginalArrivalTime (the message ID).
    """
    base = get_chatsvc_base_url()
    client_msg_id = str(random.randint(10**18, 10**19 - 1))
    resp = await client.post(
        f"{base}/conversations/{conversation_id}/messages",
        headers={"Content-Type": "text/json"},
        content=f'{{"content":"<p>{content}</p>","messagetype":"RichText/Html","contenttype":"text","clientmessageid":"{client_msg_id}"}}',
    )
    resp.raise_for_status()
    return resp.json()


async def get_thread_members(
    client: httpx.AsyncClient, conversation_id: str
) -> list[dict]:
    """Get members of a conversation thread."""
    base = get_chatsvc_base_url()
    # Members are under /threads/, not /conversations/
    threads_base = base.replace("/users/ME", "")
    resp = await client.get(
        f"{threads_base}/threads/{conversation_id}",
    )
    resp.raise_for_status()
    data = resp.json()
    return data.get("members", [])


async def get_activity(
    client: httpx.AsyncClient,
    feed: str = "notifications",
    page_size: int = 20,
) -> list[dict]:
    """Get activity feed messages.

    Args:
        feed: One of "notifications", "mentions", "calllogs".
    """
    return await get_messages(client, f"48:{feed}", page_size)


async def mark_as_read(
    client: httpx.AsyncClient,
    conversation_id: str,
    message_id: str,
    message_version: str,
    client_message_id: str = "0",
) -> None:
    """Mark a conversation as read up to a specific message."""
    base = get_chatsvc_base_url()
    horizon = f"{message_id};{message_version};{client_message_id}"
    resp = await client.put(
        f"{base}/conversations/{conversation_id}/properties",
        params={"name": "consumptionhorizon"},
        headers={"Content-Type": "text/json"},
        content=f'{{"consumptionhorizon":"{horizon}"}}',
    )
    resp.raise_for_status()


async def search_messages(
    client: httpx.AsyncClient, query: str, size: int = 25, offset: int = 0
) -> dict:
    """Search for messages using the Substrate Search API.

    Returns the raw response including entitySets with results.
    """
    cvid = str(uuid.uuid4())
    logical_id = str(uuid.uuid4())
    body = {
        "entityRequests": [
            {
                "entityType": "Message",
                "contentSources": ["Teams"],
                "fields": [
                    "Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String",
                    "Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String",
                    "Extension_SkypeSpaces_ConversationPost_Extension_SkypeGroupId_String",
                ],
                "propertySet": "Optimized",
                "query": {
                    "queryString": query,
                    "displayQueryString": query,
                },
                "from": offset,
                "size": size,
                "topResultsCount": 5,
            }
        ],
        "QueryAlterationOptions": {
            "EnableAlteration": True,
            "EnableSuggestion": True,
            "SupportedRecourseDisplayTypes": ["Suggestion"],
        },
        "cvid": cvid,
        "logicalId": logical_id,
        "scenario": {
            "Dimensions": [
                {"DimensionName": "QueryType", "DimensionValue": "Messages"},
                {
                    "DimensionName": "FormFactor",
                    "DimensionValue": "general.web.reactSearch",
                },
            ],
            "Name": "powerbar",
        },
        "timezone": "UTC",
    }
    resp = await client.post("/searchservice/api/v2/query", json=body)
    resp.raise_for_status()
    return resp.json()
