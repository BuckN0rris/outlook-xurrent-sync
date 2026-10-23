"""Validation of the settings that depend on each other. No network."""

from __future__ import annotations

import os

import pytest

from outlook_xurrent.config import Config, ConfigError

IN_PROGRESS = "In Bearbeitung"


@pytest.fixture
def env(monkeypatch: pytest.MonkeyPatch) -> pytest.MonkeyPatch:
    for key in list(os.environ):
        if key.startswith(
            ("XURRENT_", "GRAPH_", "OPEN_", "DONE_", "WAITING_", "MENTION_", "LOOKBACK_")
        ):
            monkeypatch.delenv(key, raising=False)
    monkeypatch.setenv("XURRENT_TOKEN", "test-token")
    monkeypatch.setenv("XURRENT_ACCOUNT", "example-it")
    monkeypatch.setenv("OPEN_CATEGORY", f"Nicht erledigt,{IN_PROGRESS}")
    monkeypatch.setenv("DONE_CATEGORY", "Erledigt")
    return monkeypatch


def test_waiting_is_off_by_default(env) -> None:
    config = Config.from_env()
    assert config.waiting_category == ""
    assert config.waiting_statuses == frozenset()


def test_waiting_statuses_default_to_both_waiting_states(env) -> None:
    env.setenv("WAITING_CATEGORY", IN_PROGRESS)
    assert Config.from_env().waiting_statuses == {"waiting_for", "waiting_for_customer"}


def test_waiting_category_must_be_an_open_category(env) -> None:
    """Otherwise a mail moved there would never be set to done."""
    env.setenv("WAITING_CATEGORY", "Wartet")
    with pytest.raises(ConfigError, match="OPEN_CATEGORY"):
        Config.from_env()


def test_waiting_category_without_open_categories_is_accepted(env) -> None:
    """Empty OPEN_CATEGORY already treats every mail without done as open."""
    env.delenv("OPEN_CATEGORY")
    env.setenv("WAITING_CATEGORY", IN_PROGRESS)
    assert Config.from_env().waiting_category == IN_PROGRESS


def test_waiting_category_must_differ_from_done(env) -> None:
    env.setenv("WAITING_CATEGORY", "erledigt")
    with pytest.raises(ConfigError, match="DONE_CATEGORY"):
        Config.from_env()


def test_waiting_and_done_statuses_must_not_overlap(env) -> None:
    env.setenv("WAITING_CATEGORY", IN_PROGRESS)
    env.setenv("DONE_STATUSES", "completed,waiting_for")
    with pytest.raises(ConfigError, match="overlap"):
        Config.from_env()


def test_unknown_waiting_status_is_rejected(env) -> None:
    env.setenv("WAITING_CATEGORY", IN_PROGRESS)
    env.setenv("WAITING_STATUSES", "waiting")
    with pytest.raises(ConfigError, match="WAITING_STATUSES"):
        Config.from_env()


def test_invalid_mention_regex_is_a_config_error(env) -> None:
    env.setenv("MENTION_REGEX", "you were (mentioned")
    with pytest.raises(ConfigError, match="MENTION_REGEX"):
        Config.from_env()


def test_empty_mention_regex_is_accepted(env) -> None:
    env.setenv("MENTION_REGEX", "")
    assert Config.from_env().mention_regex == ""


def test_token_does_not_show_up_in_the_repr(env) -> None:
    assert "test-token" not in repr(Config.from_env())
