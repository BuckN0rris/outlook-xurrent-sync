"""Unit tests for the done-decision and the category rewrite. No network."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from outlook_xurrent.decide import (
    REASON_AFTER_COMPLETION,
    REASON_DONE,
    REASON_MENTION,
    REASON_NO_REQUEST_ID,
    REASON_STILL_OPEN,
    REASON_UNRESOLVED,
    REASON_WAITING,
    categories_changed,
    decide_mail,
    is_candidate,
    next_categories,
    parse_iso,
    waiting_categories,
)
from outlook_xurrent.xurrent import status_map

FIXTURES = Path(__file__).parent / "fixtures"
OPEN = "Nicht erledigt"
DONE = "Erledigt"
DONE_STATUSES = frozenset({"completed"})


@pytest.fixture(scope="module")
def records() -> dict[str, dict]:
    raw = json.loads((FIXTURES / "xurrent_requests.json").read_text(encoding="utf-8"))
    return {str(r["id"]): r for r in raw}


@pytest.fixture(scope="module")
def statuses(records: dict[str, dict]) -> dict[str, str]:
    return status_map(records)


@pytest.fixture(scope="module")
def completions(records: dict[str, dict]) -> dict[str, object]:
    return {i: parse_iso(r.get("completed_at")) for i, r in records.items()}


# --- decision --------------------------------------------------------------


def test_completed_request_is_done(statuses: dict[str, str]) -> None:
    decision = decide_mail(("80000001",), statuses, DONE_STATUSES)
    assert decision.is_done
    assert decision.reason == REASON_DONE
    assert decision.modifiable


def test_in_progress_request_is_still_open(statuses: dict[str, str]) -> None:
    decision = decide_mail(("80000002",), statuses, DONE_STATUSES)
    assert not decision.is_done
    assert decision.reason == REASON_STILL_OPEN
    assert not decision.modifiable


def test_waiting_for_customer_is_open_by_default(statuses: dict[str, str]) -> None:
    assert decide_mail(("80000003",), statuses, DONE_STATUSES).reason == REASON_STILL_OPEN


def test_declined_counts_as_done_only_when_configured(statuses: dict[str, str]) -> None:
    assert decide_mail(("70000001",), statuses, DONE_STATUSES).reason == REASON_STILL_OPEN
    wider = frozenset({"completed", "declined"})
    assert decide_mail(("70000001",), statuses, wider).reason == REASON_DONE


def test_unknown_request_is_unresolved_not_done(statuses: dict[str, str]) -> None:
    """A 403/404 request never comes back in the batch, so it stays unresolved."""
    decision = decide_mail(("80000001", "99999999"), statuses, DONE_STATUSES)
    assert decision.reason == REASON_UNRESOLVED
    assert decision.unresolved_ids == ("99999999",)
    assert not decision.modifiable


# --- note on an already closed ticket --------------------------------------
#
# Mails only arrive when a request is assigned, or when a customer or colleague
# writes a note. A note on a request that is already closed must not make the
# mail disappear from the to-do list.
#
# request 80000001 was completed at 2026-09-21T11:40:12Z.

BEFORE_CLOSURE = parse_iso("2026-09-21T10:12:00Z")
AFTER_CLOSURE = parse_iso("2026-09-21T14:00:00Z")


def test_mail_before_the_closure_is_done(
    statuses: dict[str, str], completions: dict[str, object]
) -> None:
    decision = decide_mail(
        ("80000001",),
        statuses,
        DONE_STATUSES,
        received_at=BEFORE_CLOSURE,
        completed_by_id=completions,
    )
    assert decision.reason == REASON_DONE
    assert decision.modifiable


def test_note_arriving_after_the_closure_is_left_open(
    statuses: dict[str, str], completions: dict[str, object]
) -> None:
    decision = decide_mail(
        ("80000001",),
        statuses,
        DONE_STATUSES,
        received_at=AFTER_CLOSURE,
        completed_by_id=completions,
    )
    assert decision.reason == REASON_AFTER_COMPLETION
    assert not decision.is_done
    assert not decision.modifiable


def test_without_timestamps_the_status_alone_decides(statuses: dict[str, str]) -> None:
    """Backwards compatible: no received_at means no timestamp guard."""
    assert decide_mail(("80000001",), statuses, DONE_STATUSES).reason == REASON_DONE


def test_missing_completed_at_falls_back_to_the_status(
    statuses: dict[str, str]
) -> None:
    decision = decide_mail(
        ("80000001",),
        statuses,
        DONE_STATUSES,
        received_at=AFTER_CLOSURE,
        completed_by_id={"80000001": None},
    )
    assert decision.reason == REASON_DONE


def test_closure_exactly_at_arrival_counts_as_done(
    statuses: dict[str, str], completions: dict[str, object]
) -> None:
    decision = decide_mail(
        ("80000001",),
        statuses,
        DONE_STATUSES,
        received_at=parse_iso("2026-09-21T11:40:12Z"),
        completed_by_id=completions,
    )
    assert decision.reason == REASON_DONE


def test_open_request_beats_the_timestamp_guard(
    statuses: dict[str, str], completions: dict[str, object]
) -> None:
    """A mail on a still-open request stays STILL-OPEN, not AFTER-CLOSE."""
    decision = decide_mail(
        ("80000001", "80000002"),
        statuses,
        DONE_STATUSES,
        received_at=AFTER_CLOSURE,
        completed_by_id=completions,
    )
    assert decision.reason == REASON_STILL_OPEN


@pytest.mark.parametrize(
    ("raw", "expected_none"),
    [
        ("2026-09-21T11:40:12Z", False),
        ("2026-09-21T11:40:12+02:00", False),
        ("2026-09-21T11:40:12", False),  # naive input is read as UTC
        ("", True),
        (None, True),
        ("not a date", True),
    ],
)
def test_parse_iso(raw: str | None, expected_none: bool) -> None:
    parsed = parse_iso(raw)
    assert (parsed is None) is expected_none
    if parsed is not None:
        assert parsed.tzinfo is not None


def test_no_request_id_is_never_modifiable() -> None:
    decision = decide_mail((), {}, DONE_STATUSES)
    assert decision.reason == REASON_NO_REQUEST_ID
    assert not decision.modifiable


# --- multi-request policy --------------------------------------------------


def test_policy_all_requires_every_request_done(statuses: dict[str, str]) -> None:
    decision = decide_mail(("80000001", "80000002"), statuses, DONE_STATUSES, "all")
    assert decision.reason == REASON_STILL_OPEN


def test_policy_any_accepts_one_done_request(statuses: dict[str, str]) -> None:
    decision = decide_mail(("80000001", "80000002"), statuses, DONE_STATUSES, "any")
    assert decision.reason == REASON_DONE


def test_policy_any_ignores_unresolved_when_one_is_done(statuses: dict[str, str]) -> None:
    decision = decide_mail(("80000001", "99999999"), statuses, DONE_STATUSES, "any")
    assert decision.reason == REASON_DONE


def test_policy_all_with_two_completed(statuses: dict[str, str]) -> None:
    both = dict(statuses)
    both["80000002"] = "completed"
    assert decide_mail(("80000001", "80000002"), both, DONE_STATUSES, "all").reason == REASON_DONE


# --- candidate selection ---------------------------------------------------


@pytest.mark.parametrize(
    ("categories", "expected"),
    [
        ([OPEN], True),
        ([OPEN, "Kunde A"], True),
        ([DONE], False),
        ([OPEN, DONE], False),
        (["nicht erledigt"], True),  # category matching is case-insensitive
        ([], False),
        (["Kunde A"], False),
    ],
)
def test_is_candidate(categories: list[str], expected: bool) -> None:
    assert is_candidate(categories, OPEN, DONE) is expected


def test_empty_open_category_means_any_mail_without_done() -> None:
    assert is_candidate([], "", DONE) is True
    assert is_candidate(["Kunde A"], "", DONE) is True
    assert is_candidate([DONE], "", DONE) is False


# --- category rewrite ------------------------------------------------------


def test_open_is_removed_and_done_is_added() -> None:
    assert next_categories([OPEN], OPEN, DONE) == (DONE,)


def test_other_categories_are_preserved_in_order() -> None:
    result = next_categories(["Kunde A", OPEN, "Wichtig"], OPEN, DONE)
    assert result == ("Kunde A", "Wichtig", DONE)


def test_rewrite_is_idempotent() -> None:
    once = next_categories(["Kunde A", OPEN], OPEN, DONE)
    twice = next_categories(once, OPEN, DONE)
    assert once == twice
    assert not categories_changed(once, twice)


def test_second_run_produces_no_change() -> None:
    """The guarantee that running twice writes nothing the second time."""
    current = ["Kunde A", DONE]
    proposed = next_categories(current, OPEN, DONE)
    assert not categories_changed(current, proposed)


def test_differently_cased_open_category_is_removed() -> None:
    assert next_categories(["NICHT ERLEDIGT"], OPEN, DONE) == (DONE,)


def test_duplicate_done_category_is_collapsed() -> None:
    assert next_categories([DONE, "Kunde A", DONE], OPEN, DONE) == ("Kunde A", DONE)


# --- several open categories ------------------------------------------------
#
# Mails are tracked with "Nicht erledigt" or "In Bearbeitung"; both mean the
# work is not finished, and both must give way to the done category.

IN_PROGRESS = "In Bearbeitung"
OPEN_BOTH = (OPEN, IN_PROGRESS)


@pytest.mark.parametrize(
    ("categories", "expected"),
    [
        ([OPEN], True),
        ([IN_PROGRESS], True),
        ([OPEN, IN_PROGRESS], True),
        ([IN_PROGRESS, "Kunde A"], True),
        ([DONE], False),
        ([IN_PROGRESS, DONE], False),
        (["in bearbeitung"], True),  # still case-insensitive
        ([], False),
        (["Kunde A"], False),
    ],
)
def test_is_candidate_with_two_open_categories(
    categories: list[str], expected: bool
) -> None:
    assert is_candidate(categories, OPEN_BOTH, DONE) is expected


def test_in_progress_is_replaced_by_done() -> None:
    assert next_categories([IN_PROGRESS], OPEN_BOTH, DONE) == (DONE,)


def test_both_open_categories_are_removed_together() -> None:
    result = next_categories(["Kunde A", OPEN, IN_PROGRESS], OPEN_BOTH, DONE)
    assert result == ("Kunde A", DONE)


def test_two_category_rewrite_is_idempotent() -> None:
    once = next_categories([IN_PROGRESS, "Kunde A"], OPEN_BOTH, DONE)
    twice = next_categories(once, OPEN_BOTH, DONE)
    assert once == twice
    assert not categories_changed(once, twice)


def test_single_name_as_string_still_works() -> None:
    """Backwards compatible: one category may be passed as a plain string."""
    assert is_candidate([OPEN], OPEN, DONE) is True
    assert is_candidate([IN_PROGRESS], OPEN, DONE) is False
    assert next_categories([OPEN], OPEN, DONE) == (DONE,)


# --- grace window for the closing notification ------------------------------
#
# Xurrent sends the "completed" notification seconds after completed_at, so
# without a grace window the one mail announcing the closure is the one left
# open. Measured against a real mailbox: notifications 12-75 s after the
# closure, closing notes within minutes, genuine follow-ups 28 h and later.

CLOSED_AT = "2026-09-04T09:37:03+02:00"
GRACE_2H = 7200


def _decide_at(received_iso: str, grace: int):
    return decide_mail(
        ("80000001",),
        {"80000001": "completed"},
        DONE_STATUSES,
        received_at=parse_iso(received_iso),
        completed_by_id={"80000001": parse_iso(CLOSED_AT)},
        grace_seconds=grace,
    )


def test_closing_notification_22_seconds_later_counts_as_done() -> None:
    assert _decide_at("2026-09-04T09:37:25+02:00", GRACE_2H).reason == REASON_DONE


def test_closing_note_23_minutes_later_counts_as_done() -> None:
    assert _decide_at("2026-09-04T10:00:00+02:00", GRACE_2H).reason == REASON_DONE


def test_follow_up_the_next_day_is_still_held_back() -> None:
    decision = _decide_at("2026-09-05T14:00:00+02:00", GRACE_2H)
    assert decision.reason == REASON_AFTER_COMPLETION
    assert not decision.modifiable


def test_mail_exactly_at_the_grace_boundary_counts_as_done() -> None:
    assert _decide_at("2026-09-04T11:37:03+02:00", GRACE_2H).reason == REASON_DONE


def test_one_second_past_the_boundary_is_held_back() -> None:
    assert (
        _decide_at("2026-09-04T11:37:04+02:00", GRACE_2H).reason
        == REASON_AFTER_COMPLETION
    )


def test_zero_grace_restores_the_strict_rule() -> None:
    assert (
        _decide_at("2026-09-04T09:37:25+02:00", 0).reason == REASON_AFTER_COMPLETION
    )


def test_mail_before_the_closure_is_unaffected_by_the_grace() -> None:
    assert _decide_at("2026-09-04T08:00:00+02:00", GRACE_2H).reason == REASON_DONE


# --- waiting on someone else ------------------------------------------------
#
# A request in waiting_for or waiting_for_customer is not finished, but the
# next step lies with someone else. Its mails move from "Nicht erledigt" to
# "In Bearbeitung" and stay in the open set, so they still reach "Erledigt"
# once the request closes.

WAITING_STATUSES = frozenset({"waiting_for", "waiting_for_customer"})


def _decide_waiting(statuses_by_id: dict[str, str], policy: str = "all", **kwargs):
    return decide_mail(
        tuple(statuses_by_id),
        statuses_by_id,
        DONE_STATUSES,
        policy,
        waiting_statuses=WAITING_STATUSES,
        **kwargs,
    )


def test_waiting_for_customer_is_waiting_when_configured(
    statuses: dict[str, str]
) -> None:
    decision = decide_mail(
        ("80000003",), statuses, DONE_STATUSES, waiting_statuses=WAITING_STATUSES
    )
    assert decision.reason == REASON_WAITING
    assert not decision.is_done
    assert decision.modifiable


def test_waiting_for_is_waiting_too() -> None:
    assert _decide_waiting({"1": "waiting_for"}).reason == REASON_WAITING


@pytest.mark.parametrize("status", ["assigned", "accepted", "in_progress", "on_backlog"])
def test_statuses_that_need_action_stay_open(status: str) -> None:
    assert _decide_waiting({"1": status}).reason == REASON_STILL_OPEN


def test_completed_is_still_done_with_waiting_configured() -> None:
    assert _decide_waiting({"1": "completed"}).reason == REASON_DONE


def test_policy_all_done_plus_waiting_is_waiting() -> None:
    decision = _decide_waiting({"1": "completed", "2": "waiting_for"}, "all")
    assert decision.reason == REASON_WAITING


def test_policy_any_done_beats_waiting() -> None:
    decision = _decide_waiting({"1": "completed", "2": "waiting_for"}, "any")
    assert decision.reason == REASON_DONE


@pytest.mark.parametrize("policy", ["all", "any"])
def test_all_requests_waiting_is_waiting(policy: str) -> None:
    decision = _decide_waiting(
        {"1": "waiting_for", "2": "waiting_for_customer"}, policy
    )
    assert decision.reason == REASON_WAITING


@pytest.mark.parametrize("policy", ["all", "any"])
def test_one_request_that_needs_action_keeps_the_mail_open(policy: str) -> None:
    decision = _decide_waiting({"1": "waiting_for", "2": "in_progress"}, policy)
    assert decision.reason == REASON_STILL_OPEN
    assert not decision.modifiable


@pytest.mark.parametrize("policy", ["all", "any"])
def test_unreadable_request_blocks_the_waiting_move(policy: str) -> None:
    decision = decide_mail(
        ("1", "99999999"),
        {"1": "waiting_for"},
        DONE_STATUSES,
        policy,
        waiting_statuses=WAITING_STATUSES,
    )
    assert decision.reason == REASON_UNRESOLVED
    assert not decision.modifiable


@pytest.mark.parametrize("policy", ["all", "any"])
def test_note_on_a_closed_request_beats_waiting(policy: str) -> None:
    """The mail is a note on the closed request and may need an answer."""
    decision = _decide_waiting(
        {"1": "waiting_for", "2": "completed"},
        policy,
        received_at=parse_iso("2026-09-25T10:00:00Z"),
        completed_by_id={"2": parse_iso("2026-09-21T11:40:12Z")},
    )
    assert decision.reason == REASON_AFTER_COMPLETION
    assert not decision.modifiable


# --- waiting rewrite --------------------------------------------------------


def test_open_is_swapped_for_the_waiting_category() -> None:
    assert waiting_categories([OPEN], OPEN_BOTH, IN_PROGRESS) == (IN_PROGRESS,)


def test_waiting_rewrite_keeps_other_categories_in_order() -> None:
    result = waiting_categories(["Kunde A", OPEN, "Wichtig"], OPEN_BOTH, IN_PROGRESS)
    assert result == ("Kunde A", "Wichtig", IN_PROGRESS)


def test_mail_already_waiting_is_not_rewritten() -> None:
    """The waiting category keeps its place, so nothing is written."""
    current = ["Kunde A", IN_PROGRESS, "Wichtig"]
    proposed = waiting_categories(current, OPEN_BOTH, IN_PROGRESS)
    assert not categories_changed(current, proposed)


def test_open_and_waiting_together_keep_only_waiting() -> None:
    assert waiting_categories([IN_PROGRESS, OPEN], OPEN_BOTH, IN_PROGRESS) == (
        IN_PROGRESS,
    )


def test_duplicate_waiting_category_is_collapsed() -> None:
    result = waiting_categories(
        [IN_PROGRESS, "Kunde A", "in bearbeitung"], OPEN_BOTH, IN_PROGRESS
    )
    assert result == (IN_PROGRESS, "Kunde A")


def test_waiting_rewrite_is_idempotent() -> None:
    once = waiting_categories(["Kunde A", OPEN], OPEN_BOTH, IN_PROGRESS)
    assert waiting_categories(once, OPEN_BOTH, IN_PROGRESS) == once


def test_waiting_mail_is_set_to_done_later() -> None:
    waiting = waiting_categories(["Kunde A", OPEN], OPEN_BOTH, IN_PROGRESS)
    assert is_candidate(waiting, OPEN_BOTH, DONE)
    assert next_categories(waiting, OPEN_BOTH, DONE) == ("Kunde A", DONE)


# --- mentions stay open -------------------------------------------------------
#
# "You were mentioned in Request #..." means someone asked you something. On a
# waiting request that mail keeps "Nicht erledigt" instead of moving along.


def test_mention_on_a_waiting_request_is_left_alone() -> None:
    decision = _decide_waiting({"1": "waiting_for_customer"}, is_mention=True)
    assert decision.reason == REASON_MENTION
    assert not decision.modifiable


@pytest.mark.parametrize("policy", ["all", "any"])
def test_mention_beats_waiting_under_both_policies(policy: str) -> None:
    decision = _decide_waiting(
        {"1": "waiting_for", "2": "waiting_for_customer"}, policy, is_mention=True
    )
    assert decision.reason == REASON_MENTION


def test_mention_on_a_done_request_is_still_done() -> None:
    assert _decide_waiting({"1": "completed"}, is_mention=True).reason == REASON_DONE


def test_mention_on_an_open_request_is_still_open() -> None:
    assert (
        _decide_waiting({"1": "in_progress"}, is_mention=True).reason
        == REASON_STILL_OPEN
    )


def test_mention_without_waiting_configured_is_still_open() -> None:
    decision = decide_mail(
        ("1",), {"1": "waiting_for_customer"}, DONE_STATUSES, is_mention=True
    )
    assert decision.reason == REASON_STILL_OPEN


def test_parse_iso_reads_the_utc_suffix_graph_sends() -> None:
    from datetime import datetime, timezone

    from outlook_xurrent.decide import parse_iso

    assert parse_iso("2026-09-21T10:12:00Z") == datetime(2026, 9, 21, 10, 12, tzinfo=timezone.utc)
    assert parse_iso("2026-09-21T12:12:00+02:00") == parse_iso("2026-09-21T10:12:00Z")
