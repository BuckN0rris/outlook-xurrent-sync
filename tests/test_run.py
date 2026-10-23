"""End-to-end test of the apply step against a stub Graph client.

Proves the guarantees from the README without touching the network:
dry-run writes nothing, only done mails are written, mails without a request id
and notes on closed requests are never written, and a second pass is a no-op.
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Sequence

import pytest

from outlook_xurrent import report as rp
from outlook_xurrent.cli import apply_decisions
from outlook_xurrent.config import Config
from outlook_xurrent.daily import DailyLog, KIND_APPLIED, KIND_RUN, format_run, local_now
from outlook_xurrent.decide import is_candidate, parse_iso
from outlook_xurrent.state import State
from outlook_xurrent.xurrent import status_map

FIXTURES = Path(__file__).parent / "fixtures"
LOG = logging.getLogger("test")

OPEN = "Nicht erledigt"
DONE = "Erledigt"


class StubGraph:
    """Records PATCH calls instead of performing them."""

    def __init__(self) -> None:
        self.writes: list[tuple[str, list[str]]] = []

    def patch_categories(self, message_id: str, categories: Sequence[str]) -> None:
        self.writes.append((message_id, list(categories)))


@pytest.fixture
def config(monkeypatch: pytest.MonkeyPatch) -> Config:
    for key in list(__import__("os").environ):
        if key.startswith(
            ("XURRENT_", "GRAPH_", "OPEN_", "DONE_", "WAITING_", "MENTION_", "LOOKBACK_")
        ):
            monkeypatch.delenv(key, raising=False)
    monkeypatch.setenv("XURRENT_TOKEN", "test-token")
    monkeypatch.setenv("XURRENT_ACCOUNT", "example-it")
    monkeypatch.setenv("OPEN_CATEGORY", OPEN)
    monkeypatch.setenv("DONE_CATEGORY", DONE)
    return Config.from_env()


@pytest.fixture
def xurrent_data() -> tuple[dict[str, str], dict[str, Any]]:
    raw = json.loads((FIXTURES / "xurrent_requests.json").read_text(encoding="utf-8"))
    records = {str(r["id"]): r for r in raw}
    completions = {i: parse_iso(r.get("completed_at")) for i, r in records.items()}
    return status_map(records), completions


def mail(
    mail_id: str,
    request_ids: tuple[str, ...],
    received: str,
    categories: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "id": mail_id,
        "subject": f"mail {mail_id}",
        "receivedDateTime": received,
        "categories": categories if categories is not None else [OPEN],
        "internetMessageId": f"<{mail_id}@test>",
        "_request_ids": request_ids,
        "_source": "body-link",
        "_folder": "Inbox",
    }


def run(
    graph: StubGraph,
    config: Config,
    candidates: list[dict[str, Any]],
    xurrent_data: tuple[dict[str, str], dict[str, Any]],
    tmp_path: Path,
    *,
    apply: bool,
) -> rp.Counters:
    statuses, completions = xurrent_data
    counters = rp.Counters()
    apply_decisions(
        graph,  # type: ignore[arg-type]  - only patch_categories is used
        config,
        State(path=tmp_path / "state.json"),
        candidates,
        statuses,
        completions,
        DailyLog(tmp_path / "daily"),
        apply=apply,
        limit=None,
        counters=counters,
        report=rp.Report(counters=counters),
        log=LOG,
    )
    return counters


# --- the completed request, mail received before the closure ---------------

DONE_MAIL = ("m-done", ("80000001",), "2026-09-21T10:12:00Z")


def test_dry_run_writes_nothing(config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    counters = run(graph, config, [mail(*DONE_MAIL)], xurrent_data, tmp_path, apply=False)
    assert graph.writes == []
    assert counters.would_apply == 1
    assert counters.applied == 0


def test_apply_writes_exactly_the_done_mail(config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    counters = run(graph, config, [mail(*DONE_MAIL)], xurrent_data, tmp_path, apply=True)
    assert graph.writes == [("m-done", [DONE])]
    assert counters.applied == 1


def test_other_categories_survive_the_write(config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    candidate = mail(*DONE_MAIL, categories=["Kunde A", OPEN, "Wichtig"])
    run(graph, config, [candidate], xurrent_data, tmp_path, apply=True)
    assert graph.writes == [("m-done", ["Kunde A", "Wichtig", DONE])]


def test_second_pass_writes_nothing(config, xurrent_data, tmp_path) -> None:
    """Idempotency: after the first write the mail carries the done category."""
    graph = StubGraph()
    candidate = mail(*DONE_MAIL)
    run(graph, config, [candidate], xurrent_data, tmp_path, apply=True)
    candidate["categories"] = graph.writes[0][1]  # what Outlook now holds
    graph.writes.clear()
    counters = run(graph, config, [candidate], xurrent_data, tmp_path, apply=True)
    assert graph.writes == []
    assert counters.unchanged == 1


# --- everything that must never be written ---------------------------------


def test_mail_without_request_id_is_never_written(config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    candidate = mail("m-no-id", (), "2026-09-18T09:20:00Z")
    candidate["_source"] = "none"
    counters = run(graph, config, [candidate], xurrent_data, tmp_path, apply=True)
    assert graph.writes == []
    assert counters.applied == 0


def test_note_after_closure_is_never_written(config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    candidate = mail("m-note", ("80000001",), "2026-09-22T08:00:00Z")
    counters = run(graph, config, [candidate], xurrent_data, tmp_path, apply=True)
    assert graph.writes == []
    assert counters.after_completion == 1


def test_open_request_is_never_written(config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    candidate = mail("m-open", ("80000002",), "2026-09-19T14:41:00Z")
    counters = run(graph, config, [candidate], xurrent_data, tmp_path, apply=True)
    assert graph.writes == []
    assert counters.still_open == 1


def test_unreadable_request_is_never_written(config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    candidate = mail("m-403", ("99999999",), "2026-09-19T14:41:00Z")
    counters = run(graph, config, [candidate], xurrent_data, tmp_path, apply=True)
    assert graph.writes == []
    assert counters.unresolved == 1


def test_limit_caps_the_number_of_writes(config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    candidates = [
        mail(f"m-{n}", ("80000001",), "2026-09-21T10:12:00Z") for n in range(4)
    ]
    statuses, completions = xurrent_data
    counters = rp.Counters()
    apply_decisions(
        graph,  # type: ignore[arg-type]
        config,
        State(path=tmp_path / "state.json"),
        candidates,
        statuses,
        completions,
        DailyLog(tmp_path / "daily"),
        apply=True,
        limit=2,
        counters=counters,
        report=rp.Report(counters=counters),
        log=LOG,
    )
    assert len(graph.writes) == 2
    assert counters.applied == 2
    assert counters.would_apply == 2


def test_all_mails_of_one_request_are_updated(config, xurrent_data, tmp_path) -> None:
    """Assignment mail plus note mails of the same request: all get updated."""
    graph = StubGraph()
    candidates = [
        mail("m-assigned", ("80000001",), "2026-09-21T09:00:00Z"),
        mail("m-note-1", ("80000001",), "2026-09-21T10:00:00Z"),
        mail("m-note-2", ("80000001",), "2026-09-21T11:00:00Z"),
    ]
    counters = run(graph, config, candidates, xurrent_data, tmp_path, apply=True)
    assert [w[0] for w in graph.writes] == ["m-assigned", "m-note-1", "m-note-2"]
    assert counters.applied == 3


# --- waiting requests move to "In Bearbeitung" ------------------------------
#
# request 80000003 is waiting_for_customer.

IN_PROGRESS = "In Bearbeitung"
WAITING_MAIL = ("m-waiting", ("80000003",), "2026-09-20T08:00:00Z")


@pytest.fixture
def waiting_config(config: Config, monkeypatch: pytest.MonkeyPatch) -> Config:
    monkeypatch.setenv("OPEN_CATEGORY", f"{OPEN},{IN_PROGRESS}")
    monkeypatch.setenv("WAITING_CATEGORY", IN_PROGRESS)
    return Config.from_env()


def test_waiting_mail_moves_to_in_progress(
    waiting_config, xurrent_data, tmp_path
) -> None:
    graph = StubGraph()
    candidate = mail(*WAITING_MAIL, categories=["Kunde A", OPEN])
    counters = run(graph, waiting_config, [candidate], xurrent_data, tmp_path, apply=True)
    assert graph.writes == [("m-waiting", ["Kunde A", IN_PROGRESS])]
    assert counters.waiting == 1
    assert counters.applied == 1


def test_waiting_move_respects_dry_run(waiting_config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    counters = run(
        graph, waiting_config, [mail(*WAITING_MAIL)], xurrent_data, tmp_path, apply=False
    )
    assert graph.writes == []
    assert counters.would_apply == 1


def test_mail_already_in_progress_is_not_written(
    waiting_config, xurrent_data, tmp_path
) -> None:
    graph = StubGraph()
    candidate = mail(*WAITING_MAIL, categories=[IN_PROGRESS, "Kunde A"])
    counters = run(graph, waiting_config, [candidate], xurrent_data, tmp_path, apply=True)
    assert graph.writes == []
    assert counters.unchanged == 1


def test_waiting_move_is_off_without_waiting_category(
    config, xurrent_data, tmp_path
) -> None:
    graph = StubGraph()
    counters = run(graph, config, [mail(*WAITING_MAIL)], xurrent_data, tmp_path, apply=True)
    assert graph.writes == []
    assert counters.still_open == 1
    assert counters.waiting == 0


def test_waiting_mail_is_set_to_done_once_the_request_closes(
    waiting_config, xurrent_data, tmp_path
) -> None:
    graph = StubGraph()
    candidate = mail(*WAITING_MAIL)
    run(graph, waiting_config, [candidate], xurrent_data, tmp_path, apply=True)
    candidate["categories"] = graph.writes[0][1]  # what Outlook now holds
    assert is_candidate(
        candidate["categories"], waiting_config.open_categories, DONE
    )

    statuses, completions = xurrent_data
    closed = {**statuses, "80000003": "completed"}
    graph.writes.clear()
    counters = run(
        graph, waiting_config, [candidate], (closed, completions), tmp_path, apply=True
    )
    assert graph.writes == [("m-waiting", [DONE])]
    assert counters.done == 1


def test_mention_on_a_waiting_request_keeps_its_category(
    waiting_config, xurrent_data, tmp_path
) -> None:
    graph = StubGraph()
    mention = mail(*WAITING_MAIL)
    mention["subject"] = "👉You were mentioned in Request #80000003 - Onboarding"
    counters = run(graph, waiting_config, [mention], xurrent_data, tmp_path, apply=True)
    assert graph.writes == []
    assert counters.mentions == 1
    assert counters.waiting == 0


def test_only_the_mention_of_a_waiting_request_stays(
    waiting_config, xurrent_data, tmp_path
) -> None:
    """Assignment mail moves along, the mention next to it does not."""
    graph = StubGraph()
    assigned = mail("m-assigned", ("80000003",), "2026-09-20T08:00:00Z")
    assigned["subject"] = "📫 New assigned Request #80000003 - Onboarding"
    mention = mail("m-mention", ("80000003",), "2026-09-20T09:00:00Z")
    mention["subject"] = "👉You were mentioned in Request #80000003 - Onboarding"
    run(graph, waiting_config, [assigned, mention], xurrent_data, tmp_path, apply=True)
    assert graph.writes == [("m-assigned", [IN_PROGRESS])]


def test_mention_is_still_set_to_done(waiting_config, xurrent_data, tmp_path) -> None:
    graph = StubGraph()
    mention = mail(*DONE_MAIL)
    mention["subject"] = "👉You were mentioned in Request #80000001 - Zugangsdaten"
    run(graph, waiting_config, [mention], xurrent_data, tmp_path, apply=True)
    assert graph.writes == [("m-done", [DONE])]
