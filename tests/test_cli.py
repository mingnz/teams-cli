"""Tests for CLI commands."""

import json

from typer.testing import CliRunner

from teams_cli.cli import app

runner = CliRunner()


def test_help():
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "chats" in result.output
    assert "messages" in result.output
    assert "send" in result.output
    assert "search" in result.output
    assert "activity" in result.output
    assert "members" in result.output


def test_chats_requires_auth(monkeypatch):
    monkeypatch.setattr("teams_cli.client.get_token", lambda name: None)
    result = runner.invoke(app, ["chats"])
    assert result.exit_code == 1
    assert "token" in result.output.lower() or result.exit_code == 1


def test_resolve_short_id(tmp_path, monkeypatch):
    chats_file = tmp_path / "last_chats.json"
    chats_file.write_text(
        json.dumps(
            [
                {
                    "short_id": "lved",
                    "id": "19:resolved@thread.v2",
                    "name": "Test",
                },
            ]
        )
    )
    monkeypatch.setattr("teams_cli.cli.LAST_CHATS_FILE", chats_file)

    from teams_cli.cli import _resolve_conversation_id

    assert _resolve_conversation_id("lved") == "19:resolved@thread.v2"


def test_resolve_raw_id():
    from teams_cli.cli import _resolve_conversation_id

    assert _resolve_conversation_id("19:abc@thread.v2") == "19:abc@thread.v2"
