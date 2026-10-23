"""Pure decision logic: is a mail done, and which categories should it carry?

No I/O, no network, no logging. Everything here is unit-tested against
fixtures.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Mapping, Sequence

REASON_NO_REQUEST_ID = "no_request_id"
REASON_UNRESOLVED = "unresolved"
REASON_DONE = "done"
REASON_STILL_OPEN = "still_open"
REASON_AFTER_COMPLETION = "after_completion"
REASON_WAITING = "waiting"
REASON_MENTION = "mention"

# per-request verdict used to build the mail decision
_DONE = "done"
_OPEN = "open"
_WAITING = "waiting"
_CLOSED_BEFORE_MAIL = "closed_before_mail"


def parse_iso(value: str | None) -> datetime | None:
    """Parse an ISO 8601 timestamp from either API. Naive input is read as UTC."""
    if not value:
        return None
    text = str(value).strip()
    try:
        # Graph sends a trailing "Z", which fromisoformat only accepts from 3.11
        parsed = datetime.fromisoformat(text[:-1] + "+00:00" if text.endswith("Z") else text)
    except ValueError:
        return None
    return parsed.replace(tzinfo=timezone.utc) if parsed.tzinfo is None else parsed


@dataclass(frozen=True)
class Decision:
    """Outcome for a single mail."""

    is_done: bool
    reason: str
    resolved_ids: tuple[str, ...] = ()
    unresolved_ids: tuple[str, ...] = ()

    @property
    def modifiable(self) -> bool:
        """A mail is only ever touched when a request id was found AND its
        requests are either done or waiting on someone else."""
        return self.reason in (REASON_DONE, REASON_WAITING)


def _request_state(
    status: str,
    done_statuses: frozenset[str],
    completed_at: datetime | None,
    received_at: datetime | None,
    grace_seconds: int = 0,
    waiting_statuses: frozenset[str] = frozenset(),
) -> str:
    """Judge a single request from the point of view of one mail.

    A request that was already closed when the mail arrived does not make that
    mail done: mail like that is a note someone wrote on a finished ticket, and
    it may well still be waiting for an answer.

    ``grace_seconds`` carves out the mail that *reports* the closure. Xurrent
    sends it seconds after completed_at, so without a grace window the one mail
    announcing "completed" would be the one left open. Measured against a real
    mailbox the closing notifications land 12-75 s after the closure and the
    closing notes within a few minutes, while genuine follow-ups arrive hours
    or days later.

    Without both timestamps the status alone decides.

    A request in one of ``waiting_statuses`` is not finished, but the next step
    lies with someone else (customer, supplier, colleague).
    """
    if status.lower() in waiting_statuses:
        return _WAITING
    if status.lower() not in done_statuses:
        return _OPEN
    if received_at is None or completed_at is None:
        return _DONE
    deadline = completed_at + timedelta(seconds=max(0, grace_seconds))
    return _DONE if deadline >= received_at else _CLOSED_BEFORE_MAIL


def decide_mail(
    request_ids: Sequence[str],
    status_by_id: Mapping[str, str],
    done_statuses: frozenset[str],
    policy: str = "all",
    *,
    received_at: datetime | None = None,
    completed_by_id: Mapping[str, datetime | None] | None = None,
    grace_seconds: int = 0,
    waiting_statuses: frozenset[str] = frozenset(),
    is_mention: bool = False,
) -> Decision:
    """Decide whether a mail counts as done.

    ``status_by_id`` contains only requests that were actually readable. Ids
    missing from it are inaccessible (404/403) or simply unknown, and are
    reported as unresolved instead of being guessed at.

    policy ``all``: every referenced request must be readable and done.
    policy ``any``: one readable, done request is enough.

    Passing ``received_at`` and ``completed_by_id`` enables the note-on-a-closed
    ticket guard described in ``_request_state``, with ``grace_seconds`` as its
    tolerance. Without them the decision rests on the status alone.

    ``waiting_statuses`` enables the waiting verdict: a mail whose requests are
    all done or waiting, with at least one waiting, is reported as waiting. A
    single request that still needs action, or a note on a closed request,
    keeps the mail where it is. Empty means the verdict never occurs.

    ``is_mention`` marks an @-mention notification. Someone asked you
    something, so instead of waiting the mail is reported as a mention and left
    where it is. Done and every other verdict are unaffected.
    """
    ids = tuple(request_ids)
    if not ids:
        return Decision(is_done=False, reason=REASON_NO_REQUEST_ID)

    resolved = tuple(i for i in ids if i in status_by_id)
    unresolved = tuple(i for i in ids if i not in status_by_id)
    waiting = REASON_MENTION if is_mention else REASON_WAITING
    completions = completed_by_id or {}
    states = [
        _request_state(
            status_by_id[i],
            done_statuses,
            completions.get(i),
            received_at,
            grace_seconds,
            waiting_statuses,
        )
        for i in resolved
    ]

    if unresolved and not (policy == "any" and _DONE in states):
        return Decision(False, REASON_UNRESOLVED, resolved, unresolved)

    if policy == "any":
        if _DONE in states:
            return Decision(True, REASON_DONE, resolved, unresolved)
        if _CLOSED_BEFORE_MAIL in states:
            return Decision(False, REASON_AFTER_COMPLETION, resolved, unresolved)
        if _OPEN in states:
            return Decision(False, REASON_STILL_OPEN, resolved, unresolved)
        return Decision(False, waiting, resolved, unresolved)

    # policy == "all"
    if states and all(state == _DONE for state in states):
        return Decision(True, REASON_DONE, resolved, unresolved)
    if _OPEN in states:
        return Decision(False, REASON_STILL_OPEN, resolved, unresolved)
    if _CLOSED_BEFORE_MAIL in states:
        return Decision(False, REASON_AFTER_COMPLETION, resolved, unresolved)
    # what is left is done and waiting, with at least one waiting
    return Decision(False, waiting, resolved, unresolved)


def _open_folded(open_categories: str | Sequence[str]) -> set[str]:
    """Normalise one category name or several into a case-folded set."""
    if isinstance(open_categories, str):
        names = [open_categories] if open_categories else []
    else:
        names = [name for name in open_categories if name]
    return {name.casefold() for name in names}


def is_candidate(
    categories: Sequence[str],
    open_categories: str | Sequence[str],
    done_category: str,
) -> bool:
    """True when the mail carries any open category and not the done category.

    ``open_categories`` is one name or several ("Nicht erledigt" plus
    "In Bearbeitung"); empty means "every mail without the done category counts
    as open". Comparison is case-insensitive; writing uses the exact configured
    spelling.
    """
    folded = {c.casefold() for c in categories or ()}
    if done_category.casefold() in folded:
        return False
    wanted = _open_folded(open_categories)
    if not wanted:
        return True
    return bool(folded & wanted)


def next_categories(
    current: Sequence[str],
    open_categories: str | Sequence[str],
    done_category: str,
) -> tuple[str, ...]:
    """Return the category list to write: open removed, done added, rest kept.

    Every configured open category is removed, so a mail that carried both
    "Nicht erledigt" and "In Bearbeitung" ends up with neither. Order of the
    untouched categories is preserved. Calling this on its own result returns
    that result unchanged, which is what makes the whole run idempotent.
    """
    drop = _open_folded(open_categories) | {done_category.casefold()}
    result = [c for c in (current or ()) if c.casefold() not in drop]
    result.append(done_category)
    return tuple(result)


def waiting_categories(
    current: Sequence[str],
    open_categories: str | Sequence[str],
    waiting_category: str,
) -> tuple[str, ...]:
    """Return the category list to write for a waiting mail.

    Every other open category is removed and the waiting category added, so
    "Nicht erledigt" becomes "In Bearbeitung". A waiting category the mail
    already carries keeps its place, which makes a mail that is already right
    produce no write at all. Calling this on its own result returns that
    result unchanged.
    """
    target = waiting_category.casefold()
    drop = _open_folded(open_categories) - {target}
    result: list[str] = []
    has_target = False
    for category in current or ():
        folded = category.casefold()
        if folded in drop or (folded == target and has_target):
            continue
        has_target = has_target or folded == target
        result.append(category)
    if not has_target:
        result.append(waiting_category)
    return tuple(result)


def categories_changed(
    current: Sequence[str], proposed: Sequence[str]
) -> bool:
    """True when the proposed list differs from the current one (order-sensitive)."""
    return tuple(current or ()) != tuple(proposed or ())
