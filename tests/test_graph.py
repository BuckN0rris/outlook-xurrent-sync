"""Graph client: query building and sign-in modes. No network, msal is faked."""

from __future__ import annotations

import pytest

from outlook_xurrent import graph as graph_module
from outlook_xurrent.graph import GraphAuthError, GraphClient
from outlook_xurrent.httputil import HttpError


@pytest.fixture
def client(app_config) -> GraphClient:
    return GraphClient(app_config)


def record_queries(client: GraphClient, monkeypatch, *pages):
    """Replace paging by a recorder; each call returns the next page (or raises it)."""
    calls: list[tuple[str, dict | None]] = []
    outcomes = list(pages)

    def fake(url, params=None):
        calls.append((url, params))
        outcome = outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return iter(outcome)

    monkeypatch.setattr(client, "iter_paged", fake)
    return calls


def test_category_filter_is_quoted_joined_and_combined_with_the_date(client, monkeypatch) -> None:
    calls = record_queries(client, monkeypatch, [])

    client.list_messages("folder/id", "2026-01-01T00:00:00Z", ("Nicht erledigt", "O'Neil"))

    url, params = calls[0]
    assert "folder%2Fid" in url
    assert params["$filter"] == (
        "(categories/any(c:c eq 'Nicht erledigt') or categories/any(c:c eq 'O''Neil'))"
        " and receivedDateTime ge 2026-01-01T00:00:00Z"
    )


def test_rejected_category_filter_falls_back_to_listing_the_folder(client, monkeypatch) -> None:
    calls = record_queries(client, monkeypatch, HttpError(400, "u", "no"), [{"id": "m1"}])

    assert client.list_messages("f", None, ("Nicht erledigt",)) == [{"id": "m1"}]
    assert "$filter" in calls[0][1]
    assert "$filter" not in calls[1][1]


def test_a_quote_in_the_search_term_cannot_end_the_phrase(client, monkeypatch) -> None:
    calls = record_queries(client, monkeypatch, [])

    client.search_messages('say "hi"')

    assert calls[0][1]["$search"] == '"say  hi "'


class FakeApp:
    """Stands in for msal.PublicClientApplication."""

    def __init__(self, *, accounts, silent):
        self.accounts, self.silent = accounts, silent
        self.interactive_calls = 0

    def get_accounts(self):
        return self.accounts

    def acquire_token_silent(self, scopes, account):
        return self.silent

    def acquire_token_interactive(self, scopes):
        self.interactive_calls += 1
        return {"access_token": "interactive-token"}


def use_app(monkeypatch, app: FakeApp) -> None:
    monkeypatch.setattr(graph_module.msal, "PublicClientApplication", lambda *a, **k: app)


def test_no_login_refuses_to_open_a_browser(client, monkeypatch) -> None:
    app = FakeApp(accounts=[], silent=None)
    use_app(monkeypatch, app)

    with pytest.raises(GraphAuthError, match="--no-login"):
        client.sign_in(interactive=False)
    assert app.interactive_calls == 0


def test_no_login_still_uses_a_cached_token(client, monkeypatch) -> None:
    app = FakeApp(accounts=[{"username": "me@example.com"}], silent={"access_token": "cached"})
    use_app(monkeypatch, app)

    assert client.sign_in(interactive=False) == "me@example.com"
    assert app.interactive_calls == 0


def test_default_sign_in_opens_the_browser_when_nothing_is_cached(client, monkeypatch) -> None:
    app = FakeApp(accounts=[], silent=None)
    use_app(monkeypatch, app)

    client.sign_in()

    assert app.interactive_calls == 1
