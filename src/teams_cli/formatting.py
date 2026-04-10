"""Output formatting for terminal display."""

import re
from datetime import datetime


def strip_html(text: str) -> str:
    """Remove HTML tags from text."""
    return re.sub(r"<[^>]+>", "", text).strip()


def format_timestamp(ts: str | None) -> str:
    """Format an ISO timestamp to a short readable form."""
    if not ts:
        return ""
    try:
        dt = datetime.fromisoformat(ts.replace("Z", "+00:00").split(".")[0])
        return dt.strftime("%Y-%m-%d %H:%M")
    except (ValueError, AttributeError):
        return ts[:16] if ts else ""


def get_conversation_display_name(conv: dict) -> str:
    """Extract a display name from a conversation object."""
    thread_props = conv.get("threadProperties", {})
    topic = thread_props.get("topic", "")
    if topic:
        return topic.strip()

    # For 1:1 chats, the topic is often empty
    conv_id = conv.get("id", "")
    last_msg = conv.get("lastMessage", {})
    sender = last_msg.get("imdisplayname", "")
    if sender:
        return f"Chat with {sender}"

    # Fallback to truncated ID
    return conv_id[:50]


def get_conversation_type(conv: dict) -> str:
    """Determine conversation type from thread properties."""
    thread_props = conv.get("threadProperties", {})
    thread_type = thread_props.get("threadType", "")
    product_type = thread_props.get("productThreadType", "")

    if product_type:
        return product_type

    type_map = {
        "meeting": "Meeting",
        "chat": "Chat",
        "space": "Channel",
        "streamofnotifications": "System",
        "streamofmentions": "System",
    }
    return type_map.get(thread_type, "Chat")


def make_short_id(conv_id: str) -> str:
    """Extract the unique hash portion of a conversation ID.

    Conversation IDs look like '19:{hash}@thread.v2' or '19:{hash}@thread.tacv2'.
    Returns the hash part (e.g. '73e9410b1e38453d94aa78274efcf175').
    Falls back to the full ID if parsing fails.
    """
    # Strip prefix like "19:" and suffix like "@thread.v2"
    body = conv_id.split(":", 1)[-1] if ":" in conv_id else conv_id
    body = body.split("@")[0] if "@" in body else body
    return body


def _assign_short_ids(entries: list[dict], min_len: int = 4) -> None:
    """Assign unique short IDs to each entry using the last N chars of the hash.

    Increases length until all IDs are unique.
    """
    hashes = [make_short_id(e["id"]) for e in entries]
    length = min_len
    while length <= max((len(h) for h in hashes), default=min_len):
        short_ids = [h[-length:] for h in hashes]
        if len(set(short_ids)) == len(short_ids):
            break
        length += 1
    for entry, h in zip(entries, hashes):
        entry["short_id"] = h[-length:]


def format_chat_list(conversations: list[dict]) -> list[dict]:
    """Format conversations into a list of display dicts.

    Returns list of {short_id, name, type, preview, time, id}.
    """
    results = []
    for conv in conversations:
        conv_id = conv.get("id", "")

        # Skip system conversations
        if conv_id.startswith("48:"):
            continue

        last_msg = conv.get("lastMessage", {})
        last_body = strip_html(last_msg.get("content", ""))
        last_sender = last_msg.get("imdisplayname", "")
        last_time = format_timestamp(last_msg.get("originalarrivaltime"))

        preview = ""
        if last_sender and last_body:
            preview = f"{last_sender}: {last_body[:60]}"
        elif last_body:
            preview = last_body[:80]

        results.append(
            {
                "short_id": "",
                "name": get_conversation_display_name(conv),
                "type": get_conversation_type(conv),
                "preview": preview,
                "time": last_time,
                "id": conv_id,
            }
        )

    _assign_short_ids(results)
    return results


def format_message(msg: dict) -> dict | None:
    """Format a single message for display.

    Returns None for non-displayable messages (system events, etc).
    """
    msg_type = msg.get("messagetype", "")
    if msg_type not in ("RichText/Html", "Text", "RichText"):
        return None

    sender = msg.get("imdisplayname", "Unknown")
    content = strip_html(msg.get("content", ""))
    if not content:
        return None

    return {
        "time": format_timestamp(msg.get("originalarrivaltime")),
        "sender": sender,
        "body": content,
        "id": msg.get("id", ""),
    }


def format_person(person: dict) -> dict:
    """Format a person search result for display."""
    emails = person.get("EmailAddresses", [])
    return {
        "name": person.get("DisplayName", ""),
        "email": emails[0] if emails else "",
        "mri": person.get("MRI", ""),
        "title": person.get("JobTitle", ""),
        "department": person.get("Department", ""),
        "company": person.get("CompanyName", ""),
        "type": person.get("PeopleSubtype", ""),
    }


def format_member(member: dict) -> dict:
    """Format a member for display."""
    mri = member.get("id", "")
    name = member.get("friendlyName", "")
    role = member.get("role", "User")

    # Determine type from MRI prefix
    if mri.startswith("28:"):
        member_type = "Bot"
    elif mri.startswith("8:orgid:"):
        member_type = "User"
    else:
        member_type = "Other"

    return {
        "name": name or mri,
        "role": role,
        "type": member_type,
        "mri": mri,
    }
