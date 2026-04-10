"""Tests for formatting module."""

from teams_cli.formatting import (
    format_chat_list,
    format_member,
    format_message,
    format_timestamp,
    get_conversation_display_name,
    get_conversation_type,
    strip_html,
)


class TestStripHtml:
    def test_removes_tags(self):
        assert strip_html("<p>Hello <b>world</b></p>") == "Hello world"

    def test_handles_empty(self):
        assert strip_html("") == ""

    def test_no_tags(self):
        assert strip_html("plain text") == "plain text"


class TestFormatTimestamp:
    def test_iso_format(self):
        result = format_timestamp("2025-01-15T10:30:00.000Z")
        assert result == "2025-01-15 10:30"

    def test_none(self):
        assert format_timestamp(None) == ""

    def test_empty(self):
        assert format_timestamp("") == ""

    def test_invalid_falls_back(self):
        result = format_timestamp("not-a-date-but-long-enough")
        assert result == "not-a-date-but-l"


class TestGetConversationDisplayName:
    def test_uses_topic(self, sample_conversation):
        assert get_conversation_display_name(sample_conversation) == "Project Alpha"

    def test_falls_back_to_sender(self):
        conv = {
            "id": "19:abc@thread.v2",
            "threadProperties": {},
            "lastMessage": {"imdisplayname": "Alice"},
        }
        assert get_conversation_display_name(conv) == "Chat with Alice"

    def test_falls_back_to_id(self):
        conv = {"id": "19:abc@thread.v2", "threadProperties": {}, "lastMessage": {}}
        assert get_conversation_display_name(conv) == "19:abc@thread.v2"


class TestGetConversationType:
    def test_product_type(self):
        conv = {"threadProperties": {"productThreadType": "TeamChannel"}}
        assert get_conversation_type(conv) == "TeamChannel"

    def test_thread_type_mapping(self):
        conv = {"threadProperties": {"threadType": "meeting"}}
        assert get_conversation_type(conv) == "Meeting"

    def test_default_chat(self):
        conv = {"threadProperties": {}}
        assert get_conversation_type(conv) == "Chat"


class TestFormatChatList:
    def test_skips_system(self, sample_conversation, system_conversation):
        result = format_chat_list([sample_conversation, system_conversation])
        assert len(result) == 1
        assert result[0]["name"] == "Project Alpha"

    def test_assigns_indices(self, sample_conversation):
        result = format_chat_list([sample_conversation])
        assert result[0]["index"] == 1

    def test_preview(self, sample_conversation):
        result = format_chat_list([sample_conversation])
        assert "Alice" in result[0]["preview"]
        assert "Hello world" in result[0]["preview"]


class TestFormatMessage:
    def test_formats_html_message(self, sample_message):
        result = format_message(sample_message)
        assert result is not None
        assert result["sender"] == "Bob"
        assert result["body"] == "Hey there!"

    def test_skips_system_message(self):
        msg = {"messagetype": "Event/Call", "content": "call started"}
        assert format_message(msg) is None

    def test_skips_empty_content(self):
        msg = {"messagetype": "Text", "content": ""}
        assert format_message(msg) is None


class TestFormatMember:
    def test_user(self, sample_member):
        result = format_member(sample_member)
        assert result["name"] == "Alice Smith"
        assert result["type"] == "User"
        assert result["role"] == "Admin"

    def test_bot(self):
        member = {"id": "28:bot-id", "friendlyName": "Helper Bot", "role": "User"}
        result = format_member(member)
        assert result["type"] == "Bot"

    def test_name_fallback(self):
        member = {"id": "8:orgid:xyz", "role": "User"}
        result = format_member(member)
        assert result["name"] == "8:orgid:xyz"
