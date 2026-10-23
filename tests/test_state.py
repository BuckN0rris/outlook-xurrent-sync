"""The local state file: round trip, damaged files, size limits."""

from __future__ import annotations

import json

from outlook_xurrent import state as state_module
from outlook_xurrent.state import State


def change(state: State, mail_id: str) -> None:
    state.record_change(
        mail_id=mail_id,
        internet_message_id=f"<{mail_id}@test>",
        subject="s",
        request_ids=("80000001",),
        statuses={"80000001": "completed"},
        before=["Nicht erledigt"],
        after=["Erledigt"],
    )


def test_round_trip(tmp_path) -> None:
    path = tmp_path / "state.json"
    state = State(path=path, last_successful_run="2026-10-01T00:00:00Z")
    state.remember_ids("<a@test>", ("80000001",), "subject")
    change(state, "m1")
    state.save()

    loaded = State.load(path)
    assert loaded.last_successful_run == "2026-10-01T00:00:00Z"
    assert loaded.cached_ids("<a@test>") == (("80000001",), "subject")
    assert [e["mail_id"] for e in loaded.audit] == ["m1"]


def test_missing_file_starts_empty(tmp_path) -> None:
    assert State.load(tmp_path / "nope.json").audit == []


def test_damaged_file_is_set_aside_not_overwritten(tmp_path) -> None:
    path = tmp_path / "state.json"
    path.write_text("{not json", encoding="utf-8")

    state = State.load(path)
    assert state.audit == []
    state.save()

    assert json.loads(path.read_text(encoding="utf-8"))["version"] == state_module.STATE_VERSION
    assert (tmp_path / "state.json.bad").read_text(encoding="utf-8") == "{not json"


def test_file_from_another_version_is_set_aside(tmp_path) -> None:
    path = tmp_path / "state.json"
    path.write_text(json.dumps({"version": 99, "audit": [{"mail_id": "old"}]}), encoding="utf-8")

    assert State.load(path).audit == []
    assert (tmp_path / "state.json.bad").is_file()


def test_audit_trail_is_capped_to_the_newest_entries(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(state_module, "MAX_AUDIT_ENTRIES", 3)
    state = State(path=tmp_path / "state.json")
    for n in range(5):
        change(state, f"m{n}")
    state.save()

    assert [e["mail_id"] for e in State.load(state.path).audit] == ["m2", "m3", "m4"]


def test_prune_cache_keeps_only_the_listed_keys(tmp_path) -> None:
    state = State(path=tmp_path / "state.json")
    state.remember_ids("keep", ("1",), "subject")
    state.remember_ids("drop", ("2",), "subject")

    state.prune_cache(["keep", "never-cached"])

    assert list(state.request_id_cache) == ["keep"]
