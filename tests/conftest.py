"""Shared fixtures for tests."""

import pytest


@pytest.fixture()
def sample_conversation():
    return {
        "id": "19:abc123@thread.v2",
        "threadProperties": {
            "topic": "Project Alpha",
            "threadType": "chat",
        },
        "lastMessage": {
            "imdisplayname": "Alice",
            "content": "<p>Hello world</p>",
            "originalarrivaltime": "2025-01-15T10:30:00.000Z",
        },
    }


@pytest.fixture()
def sample_message():
    return {
        "id": "1705312200000",
        "messagetype": "RichText/Html",
        "imdisplayname": "Bob",
        "content": "<p>Hey there!</p>",
        "originalarrivaltime": "2025-01-15T10:30:00.000Z",
    }


@pytest.fixture()
def sample_member():
    return {
        "id": "8:orgid:user-uuid-123",
        "friendlyName": "Alice Smith",
        "role": "Admin",
    }


@pytest.fixture()
def system_conversation():
    return {
        "id": "48:notifications",
        "threadProperties": {"threadType": "streamofnotifications"},
        "lastMessage": {},
    }
