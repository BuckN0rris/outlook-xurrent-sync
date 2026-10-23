"""One log file per day recording what was actually changed.

This is the journal you read after the fact: which mail was recategorised when,
which request it belonged to, what the categories were before and after, and
every error that occurred. ``logs/sync.log`` is the technical log and rotates
away; these files are keyed by date and stay.

Timestamps and file names use local time on purpose - this is read by a person
asking "what happened yesterday afternoon", not by a machine.

Format is tab-separated so it stays greppable:

    2026-10-08T15:45:02+02:00<TAB>RUN<TAB>mode=APPLY scanned=109 ...
    2026-10-08T15:45:02+02:00<TAB>APPLIED<TAB>80000004<TAB>completed<TAB>...
    2026-10-08T15:45:03+02:00<TAB>ERROR<TAB>80000005<TAB>HTTP 412 ...
"""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Sequence

LOG = logging.getLogger(__name__)

KIND_RUN = "RUN"
KIND_APPLIED = "APPLIED"
KIND_WOULD_APPLY = "WOULD-APPLY"
KIND_ERROR = "ERROR"


def local_now() -> datetime:
    """Current local time, timezone-aware."""
    return datetime.now().astimezone()


def stamp(moment: datetime) -> str:
    return moment.isoformat(timespec="seconds")


def _clean(value: str) -> str:
    """Collapse whitespace so one record stays on one line."""
    return " ".join((value or "").split())


def format_run(
    moment: datetime, *, apply: bool, counters: str, duration: float
) -> str:
    mode = "APPLY" if apply else "DRY-RUN"
    return "\t".join(
        (stamp(moment), KIND_RUN, f"mode={mode} {counters} duration={duration:.1f}s")
    )


def format_change(
    moment: datetime,
    *,
    applied: bool,
    request_ids: Sequence[str],
    status: str,
    before: Sequence[str],
    after: Sequence[str],
    subject: str,
) -> str:
    return "\t".join(
        (
            stamp(moment),
            KIND_APPLIED if applied else KIND_WOULD_APPLY,
            ",".join(request_ids) or "-",
            status or "-",
            f"{list(before)} -> {list(after)}",
            _clean(subject),
        )
    )


def format_error(
    moment: datetime,
    *,
    request_ids: Sequence[str],
    subject: str,
    message: str,
    web_link: str = "",
) -> str:
    parts = [
        stamp(moment),
        KIND_ERROR,
        ",".join(request_ids) or "-",
        _clean(subject),
        _clean(message)[:300],
    ]
    if web_link:
        parts.append(web_link)
    return "\t".join(parts)


class DailyLog:
    """Appends records to ``<directory>/<local date>.log``."""

    def __init__(self, directory: Path) -> None:
        self._directory = directory

    def path_for(self, moment: datetime) -> Path:
        return self._directory / f"{moment:%Y-%m-%d}.log"

    def append(self, line: str, moment: datetime | None = None) -> None:
        """Write one record. A failure here must never abort the run."""
        when = moment or local_now()
        path = self.path_for(when)
        try:
            self._directory.mkdir(parents=True, exist_ok=True)
            with path.open("a", encoding="utf-8") as handle:
                handle.write(line + "\n")
        except OSError as exc:  # pragma: no cover - disk/permission dependent
            LOG.warning("could not write the daily log %s: %s", path, exc)
