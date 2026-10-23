"""Whole runs against fake Graph and Xurrent clients: no network, no browser."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

import pytest

from outlook_xurrent import cli
from outlook_xurrent.config import ConfigError
from outlook_xurrent.state import State

FIXTURES = Path(__file__).parent / "fixtures"
REQUESTS = {str(r["id"]): r for r in json.loads((FIXTURES / "xurrent_requests.json").read_text("utf-8"))}


def mail(mail_id: str, request_id: str) -> dict[str, Any]:
    return {
        "id": mail_id,
        "subject": f"Request #{request_id} - test",
        "receivedDateTime": "2026-09-21T10:12:00Z",  # before 80000001 was completed
        "categories": ["Nicht erledigt"],
        "internetMessageId": f"<{mail_id}@test>",
        "webLink": "",
    }


class FakeGraph:
    mails: list[dict[str, Any]] = []
    fail_on_patch: int | None = None  # 1-based number of the PATCH that blows up
    patched: list[tuple[str, list[str]]] = []

    def __init__(self, config, log) -> None:
        pass

    def sign_in(self, interactive: bool = True) -> str:
        return "me@example.com"

    def list_master_categories(self) -> list[str]:
        return ["Nicht erledigt", "Erledigt"]

    def inbox_folders(self) -> list[tuple[str, str]]:
        return [("f1", "Inbox")]

    def list_messages(self, folder_id, since_iso, categories) -> list[dict[str, Any]]:
        return [dict(m) for m in self.mails]

    def search_messages(self, term: str) -> list[dict[str, Any]]:
        return [dict(m) for m in self.mails]

    def get_message_detail(self, mail_id: str) -> dict[str, Any]:
        subject = next(m["subject"] for m in self.mails if m["id"] == mail_id)
        return {"subject": subject, "body": {"content": ""}, "internetMessageHeaders": []}

    def patch_categories(self, mail_id: str, categories) -> None:
        if self.fail_on_patch == len(self.patched) + 1:
            raise RuntimeError("boom")
        self.patched.append((mail_id, list(categories)))


class FakeXurrent:
    def __init__(self, config, log) -> None:
        pass

    def ping(self) -> None:
        pass

    def fetch_requests(self, ids) -> dict[str, dict[str, Any]]:
        return {i: REQUESTS[i] for i in ids if i in REQUESTS}


@pytest.fixture
def fakes(monkeypatch: pytest.MonkeyPatch) -> type[FakeGraph]:
    FakeGraph.mails, FakeGraph.fail_on_patch, FakeGraph.patched = [], None, []
    monkeypatch.setattr(cli, "GraphClient", FakeGraph)
    monkeypatch.setattr(cli, "XurrentClient", FakeXurrent)
    return FakeGraph


@pytest.fixture
def logger():
    log = logging.getLogger("test_cli")
    yield log
    for handler in list(logging.getLogger("outlook_xurrent").handlers):
        handler.close()
        logging.getLogger("outlook_xurrent").removeHandler(handler)


def run(app_config, *flags: str, log) -> int:
    return cli.run(app_config, cli.build_arg_parser().parse_args(list(flags)), log)


def test_apply_run_writes_the_done_mail_and_keeps_the_open_one(app_config, fakes, logger) -> None:
    fakes.mails = [mail("m-done", "80000001"), mail("m-open", "80000002")]

    assert run(app_config, "--apply", log=logger) == cli.EXIT_OK

    assert fakes.patched == [("m-done", ["Erledigt"])]
    state = State.load(app_config.state_file)
    assert [e["mail_id"] for e in state.audit] == ["m-done"]
    assert state.last_successful_run is not None


def test_explain_reports_the_verdict_and_writes_nothing(app_config, fakes, logger, capsys) -> None:
    fakes.mails = [mail("m-done", "80000001")]

    assert run(app_config, "--apply", "--explain", "test", log=logger) == cli.EXIT_OK

    assert "request done" in capsys.readouterr().out
    assert fakes.patched == []


def test_dry_run_changes_nothing_in_outlook(app_config, fakes, logger) -> None:
    fakes.mails = [mail("m-done", "80000001")]

    assert run(app_config, log=logger) == cli.EXIT_OK

    assert fakes.patched == []


def test_state_cache_drops_mails_that_are_no_longer_candidates(app_config, fakes, logger) -> None:
    state = State.load(app_config.state_file)
    state.remember_ids("<gone@test>", ("80000001",), "subject")
    state.save()
    fakes.mails = [mail("m-open", "80000002")]

    run(app_config, log=logger)

    assert list(State.load(app_config.state_file).request_id_cache) == ["<m-open@test>"]


def test_a_crash_half_way_still_saves_what_was_already_written(app_config, fakes, logger) -> None:
    fakes.mails = [mail("m1", "80000001"), mail("m2", "80000001")]
    fakes.fail_on_patch = 2

    with pytest.raises(RuntimeError):
        run(app_config, "--apply", log=logger)

    assert [e["mail_id"] for e in State.load(app_config.state_file).audit] == ["m1"]


def test_main_logs_an_unexpected_crash_and_returns_a_failure_code(
    app_config, monkeypatch, logger, capsys
) -> None:
    monkeypatch.setattr(cli, "load_env_file", lambda: None)  # never read the developer's .env
    monkeypatch.setattr(cli, "run", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom")))

    assert cli.main([]) == cli.EXIT_PARTIAL

    assert "UNEXPECTED ERROR" in capsys.readouterr().err
    assert "RuntimeError: boom" in (app_config.log_dir / "sync.log").read_text("utf-8")


@pytest.mark.parametrize("flags", [["--days", "-1"], ["--limit", "-5"], ["--days", "x"]])
def test_negative_or_non_numeric_limits_are_rejected(flags) -> None:
    with pytest.raises(SystemExit):
        cli.build_arg_parser().parse_args(flags)


def test_no_login_is_off_by_default() -> None:
    assert cli.build_arg_parser().parse_args([]).no_login is False


class CategoryGraph:
    def __init__(self, names: list[str] | None) -> None:
        self.names = names

    def list_master_categories(self) -> list[str] | None:
        return self.names


def test_missing_category_is_a_config_error_that_lists_the_existing_ones(app_config, logger) -> None:
    with pytest.raises(ConfigError) as raised:
        cli.validate_categories(CategoryGraph(["Wichtig"]), app_config, logger)  # type: ignore[arg-type]
    message = str(raised.value)
    assert "'Nicht erledigt'" in message and "'Erledigt'" in message and "Wichtig" in message


def test_category_names_are_compared_case_insensitively(app_config, logger) -> None:
    cli.validate_categories(CategoryGraph(["nicht ERLEDIGT", "erledigt"]), app_config, logger)  # type: ignore[arg-type]


def test_unreadable_category_list_skips_the_check(app_config, logger) -> None:
    cli.validate_categories(CategoryGraph(None), app_config, logger)  # type: ignore[arg-type]
