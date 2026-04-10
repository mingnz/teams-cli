"""Tests for auth module."""

import time

import pytest

from teams_cli.auth import get_region, get_token, is_expired, load_tokens, save_tokens


@pytest.fixture()
def tokens_dir(tmp_path, monkeypatch):
    tokens_file = tmp_path / "tokens.json"
    monkeypatch.setattr("teams_cli.auth.TOKENS_FILE", tokens_file)
    monkeypatch.setattr("teams_cli.auth.DATA_DIR", tmp_path)
    return tmp_path, tokens_file


class TestSaveLoadTokens:
    def test_round_trip(self, tokens_dir):
        _, tokens_file = tokens_dir
        data = {"ic3": {"secret": "tok123", "expires_on": str(int(time.time()) + 3600)}}
        save_tokens(data)
        loaded = load_tokens()
        assert loaded["ic3"]["secret"] == "tok123"

    def test_load_missing(self, tokens_dir):
        assert load_tokens() is None


class TestIsExpired:
    def test_expired(self):
        entry = {"expires_on": str(int(time.time()) - 100)}
        assert is_expired(entry) is True

    def test_valid(self):
        entry = {"expires_on": str(int(time.time()) + 3600)}
        assert is_expired(entry) is False


class TestGetToken:
    def test_returns_secret(self, tokens_dir):
        _, tokens_file = tokens_dir
        data = {
            "ic3": {"secret": "mytoken", "expires_on": str(int(time.time()) + 3600)}
        }
        save_tokens(data)
        assert get_token("ic3") == "mytoken"

    def test_returns_none_if_expired(self, tokens_dir):
        _, tokens_file = tokens_dir
        data = {"ic3": {"secret": "old", "expires_on": str(int(time.time()) - 100)}}
        save_tokens(data)
        assert get_token("ic3") is None

    def test_returns_none_if_missing(self, tokens_dir):
        save_tokens({})
        assert get_token("ic3") is None


class TestGetRegion:
    def test_returns_region(self, tokens_dir):
        save_tokens({"region": "au"})
        assert get_region() == "au"

    def test_returns_none_if_no_tokens(self, tokens_dir):
        assert get_region() is None
