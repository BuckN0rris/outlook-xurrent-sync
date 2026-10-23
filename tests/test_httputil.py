"""Retry and error mapping of the shared HTTP helper. No network: a scripted session."""

from __future__ import annotations

import logging

import pytest
import requests

from outlook_xurrent import httputil
from outlook_xurrent.httputil import AuthError, HttpError, request_with_retry

LOG = logging.getLogger("test")


class FakeResponse:
    def __init__(self, status: int, headers: dict[str, str] | None = None) -> None:
        self.status_code = status
        self.headers = headers or {}
        self.text = f"body {status}"


class ScriptedSession:
    """Hands out the given responses (or raises the given exceptions) in order."""

    def __init__(self, *outcomes: FakeResponse | Exception) -> None:
        self.outcomes = list(outcomes)
        self.calls = 0

    def request(self, method: str, url: str, **kwargs):
        self.calls += 1
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome


@pytest.fixture
def sleeps(monkeypatch: pytest.MonkeyPatch) -> list[float]:
    slept: list[float] = []
    monkeypatch.setattr(httputil.time, "sleep", slept.append)
    return slept


def send(session: ScriptedSession, retries: int = 3):
    return request_with_retry(
        session, "GET", "https://example.test/x", timeout=1, max_retries=retries, logger=LOG
    )  # type: ignore[arg-type]


def test_success_needs_no_retry(sleeps) -> None:
    session = ScriptedSession(FakeResponse(200))
    assert send(session).status_code == 200
    assert sleeps == []


def test_server_error_is_retried_with_backoff(sleeps) -> None:
    session = ScriptedSession(FakeResponse(503), FakeResponse(502), FakeResponse(200))
    assert send(session).status_code == 200
    assert sleeps == [1.0, 2.0]


def test_retry_after_is_honoured(sleeps) -> None:
    session = ScriptedSession(FakeResponse(429, {"Retry-After": "7"}), FakeResponse(200))
    send(session)
    assert sleeps == [7.0]


def test_retry_after_is_capped(sleeps) -> None:
    session = ScriptedSession(FakeResponse(429, {"retry-after": "9999"}), FakeResponse(200))
    send(session)
    assert sleeps == [httputil.MAX_SLEEP_SECONDS]


def test_gives_up_after_max_retries(sleeps) -> None:
    session = ScriptedSession(*[FakeResponse(503)] * 3)
    with pytest.raises(HttpError) as raised:
        send(session)
    assert raised.value.status == 503
    assert session.calls == 3


def test_client_error_is_not_retried(sleeps) -> None:
    session = ScriptedSession(FakeResponse(404))
    with pytest.raises(HttpError) as raised:
        send(session)
    assert not isinstance(raised.value, AuthError)
    assert session.calls == 1


@pytest.mark.parametrize("status", [401, 403])
def test_auth_failures_raise_auth_error_without_retry(sleeps, status: int) -> None:
    session = ScriptedSession(FakeResponse(status))
    with pytest.raises(AuthError):
        send(session)
    assert session.calls == 1


def test_transport_error_is_retried(sleeps) -> None:
    session = ScriptedSession(requests.ConnectionError("down"), FakeResponse(200))
    assert send(session).status_code == 200


def test_transport_error_every_time_ends_as_http_error(sleeps) -> None:
    session = ScriptedSession(*[requests.Timeout("slow")] * 3)
    with pytest.raises(HttpError) as raised:
        send(session)
    assert raised.value.status == 0
