"""Unit tests for the per-day journal. Pure formatting plus a tmp_path write."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path

from outlook_xurrent.daily import (
    KIND_APPLIED,
    KIND_ERROR,
    KIND_RUN,
    KIND_WOULD_APPLY,
    DailyLog,
    format_change,
    format_error,
    format_run,
    stamp,
)

CEST = timezone(timedelta(hours=2))
MOMENT = datetime(2026, 10, 8, 15, 45, 2, tzinfo=CEST)


def test_stamp_is_iso_with_offset() -> None:
    assert stamp(MOMENT) == "2026-10-08T15:45:02+02:00"


def test_run_line() -> None:
    line = format_run(MOMENT, apply=True, counters="scanned=5 applied=1", duration=2.34)
    assert line.split("\t") == [
        "2026-10-08T15:45:02+02:00",
        KIND_RUN,
        "mode=APPLY scanned=5 applied=1 duration=2.3s",
    ]


def test_run_line_marks_a_dry_run() -> None:
    assert "mode=DRY-RUN" in format_run(MOMENT, apply=False, counters="", duration=0.0)


def test_applied_change_line() -> None:
    line = format_change(
        MOMENT,
        applied=True,
        request_ids=("80000004",),
        status="completed",
        before=["Nicht erledigt"],
        after=("Erledigt",),
        subject="New assigned Request #80000004 - Srv01",
    )
    fields = line.split("\t")
    assert fields[1] == KIND_APPLIED
    assert fields[2] == "80000004"
    assert fields[3] == "completed"
    assert fields[4] == "['Nicht erledigt'] -> ['Erledigt']"
    assert fields[5] == "New assigned Request #80000004 - Srv01"


def test_waiting_move_is_visible_from_the_categories() -> None:
    """The same record shape covers the move to the waiting category."""
    line = format_change(
        MOMENT,
        applied=True,
        request_ids=("80000006",),
        status="waiting_for_customer",
        before=["Nicht erledigt"],
        after=("In Bearbeitung",),
        subject="PC Reception",
    )
    assert "['Nicht erledigt'] -> ['In Bearbeitung']" in line


def test_dry_run_change_is_marked_differently() -> None:
    line = format_change(
        MOMENT,
        applied=False,
        request_ids=("1",),
        status="completed",
        before=[],
        after=("Erledigt",),
        subject="x",
    )
    assert line.split("\t")[1] == KIND_WOULD_APPLY


def test_several_request_ids_are_joined() -> None:
    line = format_change(
        MOMENT,
        applied=True,
        request_ids=("1", "2"),
        status="completed",
        before=[],
        after=("Erledigt",),
        subject="x",
    )
    assert line.split("\t")[2] == "1,2"


def test_subject_newlines_never_break_the_record() -> None:
    line = format_change(
        MOMENT,
        applied=True,
        request_ids=("1",),
        status="completed",
        before=[],
        after=("Erledigt",),
        subject="erste Zeile\nzweite\tZeile",
    )
    assert "\n" not in line
    assert line.split("\t")[5] == "erste Zeile zweite Zeile"


def test_error_line_carries_the_link() -> None:
    line = format_error(
        MOMENT,
        request_ids=("80000004",),
        subject="Srv01",
        message="HTTP 412 for https://graph.microsoft.com/v1.0/me/messages/AAA",
        web_link="https://outlook.office365.com/owa/?ItemID=AAA",
    )
    fields = line.split("\t")
    assert fields[1] == KIND_ERROR
    assert fields[2] == "80000004"
    assert fields[-1] == "https://outlook.office365.com/owa/?ItemID=AAA"


def test_error_line_without_a_link_has_no_trailing_field() -> None:
    line = format_error(
        MOMENT, request_ids=(), subject="x", message="boom"
    )
    assert line.split("\t") == ["2026-10-08T15:45:02+02:00", KIND_ERROR, "-", "x", "boom"]


# --- writing ----------------------------------------------------------------


def test_file_is_named_after_the_local_date(tmp_path: Path) -> None:
    journal = DailyLog(tmp_path / "daily")
    assert journal.path_for(MOMENT).name == "2026-10-08.log"


def test_append_creates_the_directory_and_adds_lines(tmp_path: Path) -> None:
    journal = DailyLog(tmp_path / "daily")
    journal.append("erste", MOMENT)
    journal.append("zweite", MOMENT)
    written = (tmp_path / "daily" / "2026-10-08.log").read_text(encoding="utf-8")
    assert written == "erste\nzweite\n"


def test_each_day_gets_its_own_file(tmp_path: Path) -> None:
    journal = DailyLog(tmp_path / "daily")
    journal.append("heute", MOMENT)
    journal.append("morgen", MOMENT + timedelta(days=1))
    names = sorted(p.name for p in (tmp_path / "daily").iterdir())
    assert names == ["2026-10-08.log", "2026-10-09.log"]


def test_umlauts_and_emoji_survive(tmp_path: Path) -> None:
    journal = DailyLog(tmp_path / "daily")
    journal.append("\U0001f4eb Anfrage Zugangsdaten für Südtirol", MOMENT)
    written = (tmp_path / "daily" / "2026-10-08.log").read_text(encoding="utf-8")
    assert "\U0001f4eb" in written and "Südtirol" in written


def test_an_unwritable_directory_does_not_raise(tmp_path: Path) -> None:
    """A failing journal must never abort a run."""
    blocker = tmp_path / "daily"
    blocker.write_text("not a directory", encoding="utf-8")
    DailyLog(blocker).append("line", MOMENT)  # must not raise
