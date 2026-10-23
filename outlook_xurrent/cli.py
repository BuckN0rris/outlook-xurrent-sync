"""Command line entry point and run orchestration.

Exit codes: 0 = success, 1 = partial failure, 2 = configuration problem.
"""

from __future__ import annotations

import argparse
import logging
import logging.handlers
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from typing import Any, Sequence

from . import report as rp
from .config import Config, ConfigError, load_env_file
from .daily import DailyLog, format_change, format_error, format_run, local_now
from .decide import (
    REASON_AFTER_COMPLETION,
    REASON_DONE,
    REASON_MENTION,
    REASON_NO_REQUEST_ID,
    REASON_STILL_OPEN,
    REASON_UNRESOLVED,
    REASON_WAITING,
    Decision,
    categories_changed,
    decide_mail,
    is_candidate,
    next_categories,
    parse_iso,
    waiting_categories,
)
from .extract import (
    compile_header_pattern,
    compile_link_pattern,
    compile_mention_pattern,
    compile_subject_pattern,
    extract_request_ids,
    is_mention,
)
from .graph import GraphAuthError, GraphClient
from .httputil import HttpError
from .state import State, utc_now_iso
from .xurrent import XurrentAuthError, XurrentClient, status_map

LOG = logging.getLogger("outlook_xurrent")

EXIT_OK = 0
EXIT_PARTIAL = 1
EXIT_CONFIG = 2


# ---------------------------------------------------------------------------
# infrastructure
# ---------------------------------------------------------------------------


def setup_logging(config: Config, verbose: bool) -> logging.Logger:
    level = logging.DEBUG if verbose else logging.INFO
    logger = logging.getLogger("outlook_xurrent")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    fmt = logging.Formatter(
        "%(asctime)s %(levelname)-7s %(name)s %(message)s", "%Y-%m-%d %H:%M:%S"
    )

    stream = logging.StreamHandler(sys.stdout)
    stream.setLevel(level)
    stream.setFormatter(fmt)
    logger.addHandler(stream)

    config.log_dir.mkdir(parents=True, exist_ok=True)
    rotating = logging.handlers.RotatingFileHandler(
        config.log_dir / "sync.log",
        maxBytes=1_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    rotating.setLevel(logging.DEBUG)
    rotating.setFormatter(fmt)
    logger.addHandler(rotating)

    logger.propagate = False
    return logger


def _non_negative(text: str) -> int:
    value = int(text)
    if value < 0:
        raise argparse.ArgumentTypeError("must be 0 or more")
    return value


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sync.py",
        description=(
            "Reconcile Outlook categories with Xurrent request status. "
            "Dry-run by default: nothing is written without --apply."
        ),
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="actually write the categories (default: dry-run, read-only)",
    )
    parser.add_argument(
        "--days",
        type=_non_negative,
        default=None,
        metavar="N",
        help="only look at mails from the last N days (default, and 0: no limit)",
    )
    parser.add_argument(
        "--limit",
        type=_non_negative,
        default=None,
        metavar="N",
        help="with --apply: change at most N mails in this run",
    )
    parser.add_argument(
        "--no-login",
        action="store_true",
        help="never open a sign-in window; fail with exit code 2 instead "
        "(for scheduled runs)",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="validate configuration, sign-in, categories and Xurrent access, then exit",
    )
    parser.add_argument(
        "--explain",
        metavar="TERM",
        default=None,
        help="say why the mails matching TERM in the subject were or were not "
        "touched, then exit (read-only, never writes)",
    )
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="ignore the mail-id to request-id cache and re-parse every candidate",
    )
    parser.add_argument("--verbose", action="store_true", help="debug logging on stdout")
    return parser


def _reconfigure_stdout() -> None:
    """Make output safe on Windows.

    Consoles default to cp1252 and mail subjects are not, so the streams are
    switched to UTF-8. Under pythonw.exe - which the scheduled task uses so no
    console window appears every few minutes - there is no stdout at all and
    print() would raise, so it is replaced by a sink. Everything that matters
    is in the log files either way.
    """
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None:
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
            continue
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):  # pragma: no cover - console dependent
                pass


# ---------------------------------------------------------------------------
# steps
# ---------------------------------------------------------------------------


def _humanise(seconds: float) -> str:
    """Render a duration for the report, coarsely but readably."""
    if seconds < 120:
        return f"{seconds:.0f}s"
    if seconds < 7200:
        return f"{seconds / 60:.0f}min"
    if seconds < 172800:
        return f"{seconds / 3600:.1f}h"
    return f"{seconds / 86400:.1f}d"


def extraction_patterns(config: Config) -> dict[str, "re.Pattern[str]"]:
    """Compile the four ID patterns once per run."""
    subject = compile_subject_pattern(config.subject_regex)
    return {
        "subject_pattern": subject,
        "link_pattern": compile_link_pattern(config.xurrent_instance_hosts),
        # the body-text step reuses the subject pattern: both look for a literal
        # '#' followed by digits, just in different text
        "hash_pattern": subject,
        "header_pattern": compile_header_pattern(config.header_regex),
    }


def completion_times(records: dict[str, dict[str, Any]]) -> dict[str, datetime | None]:
    """{request id: completed_at}, None where the request is not completed."""
    return {
        request_id: parse_iso(str(record.get("completed_at") or ""))
        for request_id, record in records.items()
    }


def decide_for(
    config: Config,
    message: dict[str, Any],
    request_ids: Sequence[str],
    statuses: dict[str, str],
    completions: dict[str, datetime | None],
    mention_pattern: "re.Pattern[str] | None",
) -> Decision:
    """The verdict for one mail, shared by the run and by --explain."""
    return decide_mail(
        request_ids,
        statuses,
        config.done_statuses,
        config.multi_request_policy,
        received_at=parse_iso(str(message.get("receivedDateTime") or "")),
        completed_by_id=completions,
        grace_seconds=config.after_close_grace_seconds,
        waiting_statuses=config.waiting_statuses,
        is_mention=is_mention(str(message.get("subject") or ""), mention_pattern),
    )


def cache_key(message: dict[str, Any]) -> str:
    """internetMessageId survives a folder move, the Graph id does not."""
    return str(message.get("internetMessageId") or message.get("id", ""))


def validate_categories(graph: GraphClient, config: Config, log: logging.Logger) -> None:
    """Fail fast when a configured category does not exist in the mailbox."""
    if not config.validate_categories:
        log.info("category validation disabled")
        return
    names = graph.list_master_categories()
    if names is None:
        return
    folded = {n.casefold() for n in names}
    wanted = [config.done_category, *config.open_categories]
    if config.waiting_category and not config.open_categories:
        # otherwise it is one of the open categories already
        wanted.append(config.waiting_category)
    missing = [w for w in wanted if w.casefold() not in folded]
    if missing:
        raise ConfigError(
            "These Outlook categories do not exist in your mailbox: "
            + ", ".join(repr(m) for m in missing)
            + "\nExisting categories: "
            + ", ".join(sorted(names))
            + "\nCreate them in Outlook first, or fix OPEN_CATEGORY / DONE_CATEGORY"
            + " / WAITING_CATEGORY."
        )
    log.info("categories validated open=%r done=%r", config.open_categories, config.done_category)


def collect_candidates(
    graph: GraphClient,
    config: Config,
    since_iso: str | None,
    counters: rp.Counters,
    log: logging.Logger,
) -> list[dict[str, Any]]:
    """List Inbox mails and keep the ones that are open but not done.

    The open category is pushed into the Graph query when configured, which
    makes the scan both complete (no date window needed) and fast. The
    client-side test below stays the authority: it also rejects mails that
    carry the done category alongside the open one.
    """
    folders = graph.inbox_folders()
    log.info("folders=%d since=%s", len(folders), since_iso or "unlimited")

    candidates: list[dict[str, Any]] = []
    for folder_id, folder_path in folders:
        folder_count = 0
        for message in graph.list_messages(folder_id, since_iso, config.open_categories):
            folder_count += 1
            counters.scanned += 1
            if is_candidate(
                message.get("categories") or [],
                config.open_categories,
                config.done_category,
            ):
                message["_folder"] = folder_path
                candidates.append(message)
        log.debug("folder path=%s messages=%d", folder_path, folder_count)

    counters.candidates = len(candidates)
    log.info("candidates=%d of scanned=%d", counters.candidates, counters.scanned)
    return candidates


def resolve_request_ids(
    graph: GraphClient,
    config: Config,
    state: State,
    candidates: list[dict[str, Any]],
    use_cache: bool,
    counters: rp.Counters,
    log: logging.Logger,
) -> None:
    """Attach ``_request_ids`` and ``_source`` to every candidate."""
    patterns = extraction_patterns(config)

    for message in candidates:
        mail_id = str(message.get("id", ""))
        key = cache_key(message)

        if use_cache:
            cached = state.cached_ids(key)
            if cached is not None:
                message["_request_ids"], message["_source"] = cached[0], f"cache:{cached[1]}"
                log.debug("cache hit mail=%s ids=%s", key, cached[0])
                continue

        try:
            detail = graph.get_message_detail(mail_id)
        except (HttpError, GraphAuthError) as exc:
            counters.errors += 1
            message["_request_ids"] = ()
            message["_source"] = "error"
            message["_error"] = str(exc)
            log.error("detail fetch failed mail=%s error=%s", mail_id, exc)
            continue

        body = (detail.get("body") or {}).get("content", "") or ""
        extraction = extract_request_ids(
            subject=str(detail.get("subject") or message.get("subject") or ""),
            body=body,
            headers=detail.get("internetMessageHeaders") or (),
            **patterns,
        )
        message["_request_ids"] = extraction.request_ids
        message["_source"] = extraction.source
        state.remember_ids(key, extraction.request_ids, extraction.source)

    for message in candidates:
        if message.get("_request_ids"):
            counters.id_found += 1
        elif message.get("_source") != "error":
            counters.no_id += 1


VERDICTS = {
    REASON_DONE: "request done -> the open category would be swapped for the done one",
    REASON_WAITING: (
        "request waiting on someone else -> the open category would be swapped "
        "for the waiting one"
    ),
    REASON_MENTION: (
        "request waiting, but someone mentioned you -> mail is left alone, it "
        "may need an answer"
    ),
    REASON_STILL_OPEN: "request still open -> mail is left alone",
    REASON_UNRESOLVED: "request not readable (403/404/unknown) -> mail is left alone",
    REASON_NO_REQUEST_ID: "no request id found -> mail is never modified",
    REASON_AFTER_COMPLETION: (
        "mail arrived after the request was closed -> mail is left alone"
    ),
}


#: verdicts that leave the mail alone: (Counters field, report action)
LEFT_ALONE = {
    REASON_UNRESOLVED: ("unresolved", rp.ACTION_UNRESOLVED),
    REASON_STILL_OPEN: ("still_open", rp.ACTION_STILL_OPEN),
    # someone mentioned you on a waiting request: a question to you
    REASON_MENTION: ("mentions", rp.ACTION_MENTION),
    # a note on a finished ticket may still be waiting for an answer
    REASON_AFTER_COMPLETION: ("after_completion", rp.ACTION_AFTER_CLOSE),
}


def proposed_categories(
    decision: Decision, current: Sequence[str], config: Config
) -> tuple[str, ...] | None:
    """The category list a modifiable mail should carry, None for the rest."""
    if decision.reason == REASON_DONE:
        return next_categories(current, config.open_categories, config.done_category)
    if decision.reason == REASON_WAITING:
        return waiting_categories(
            current, config.open_categories, config.waiting_category
        )
    return None


def explain(
    graph: GraphClient,
    xurrent: XurrentClient,
    config: Config,
    term: str,
    log: logging.Logger,
) -> int:
    """Say why one mail was touched, or why it was not. Writes nothing."""
    patterns = extraction_patterns(config)
    mention_pattern = compile_mention_pattern(config.mention_regex)
    hits = graph.search_messages(term)
    if not hits:
        print(f"\nNo mail found for {term!r}.")
        return EXIT_OK

    print(f"\n{len(hits)} mail(s) matching {term!r}")
    for message in hits:
        mail_id = str(message.get("id", ""))
        categories = list(message.get("categories") or [])
        received = parse_iso(str(message.get("receivedDateTime") or ""))

        detail = graph.get_message_detail(mail_id)
        extraction = extract_request_ids(
            subject=str(detail.get("subject") or message.get("subject") or ""),
            body=(detail.get("body") or {}).get("content", "") or "",
            headers=detail.get("internetMessageHeaders") or (),
            **patterns,
        )
        records = xurrent.fetch_requests(extraction.request_ids)
        completions = completion_times(records)
        decision = decide_for(
            config,
            message,
            extraction.request_ids,
            status_map(records),
            completions,
            mention_pattern,
        )
        candidate = is_candidate(categories, config.open_categories, config.done_category)

        print("\n" + "-" * 78)
        print(f"  subject     {message.get('subject')}")
        print(f"  received    {message.get('receivedDateTime')}")
        print(f"  categories  {categories or '(none)'}")
        open_names = " / ".join(config.open_categories) or "(any)"
        print(
            "  in scope    "
            + (
                "yes"
                if candidate
                else f"no - does not carry {open_names} without "
                f"{config.done_category!r}"
            )
        )
        print(
            f"  request ids {', '.join(extraction.request_ids) or '(none)'}"
            f"   via {extraction.source}"
        )
        for request_id in extraction.request_ids:
            record = records.get(request_id)
            if record is None:
                print(f"    {request_id}  not readable with this token")
                continue
            completed_at = completions.get(request_id)
            gap = ""
            if completed_at is not None and received is not None:
                seconds = (received - completed_at).total_seconds()
                gap = f"  mail {_humanise(abs(seconds))} " + (
                    "after the closure" if seconds > 0 else "before the closure"
                )
            print(
                f"    {request_id}  {record.get('status')}"
                f"  completed_at={record.get('completed_at') or '-'}{gap}"
            )
        print(f"  verdict     {VERDICTS.get(decision.reason, decision.reason)}")
        proposed = proposed_categories(decision, categories, config)
        if proposed is not None:
            if categories_changed(categories, proposed):
                print(f"              {categories} -> {list(proposed)}")
            else:
                print("              categories are already correct")
        if message.get("webLink"):
            print(f"  open        {message.get('webLink')}")
    print()
    return EXIT_OK


def apply_decisions(
    graph: GraphClient,
    config: Config,
    state: State,
    candidates: list[dict[str, Any]],
    statuses: dict[str, str],
    completions: dict[str, Any],
    journal: DailyLog,
    *,
    apply: bool,
    limit: int | None,
    counters: rp.Counters,
    report: rp.Report,
    log: logging.Logger,
) -> None:
    changes_applied = 0
    mention_pattern = compile_mention_pattern(config.mention_regex)

    for message in candidates:
        mail_id = str(message.get("id", ""))
        subject = str(message.get("subject") or "")
        request_ids = tuple(message.get("_request_ids") or ())
        current = list(message.get("categories") or [])

        if message.get("_source") == "error":
            report.add(
                rp.Row("-", "-", subject, rp.ACTION_ERROR, str(message.get("_error", "")))
            )
            journal.append(
                format_error(
                    local_now(),
                    request_ids=(),
                    subject=subject,
                    message=f"could not read the mail: {message.get('_error', '')}",
                    web_link=str(message.get("webLink") or ""),
                )
            )
            continue

        decision = decide_for(
            config, message, request_ids, statuses, completions, mention_pattern
        )
        status_text = ",".join(
            sorted({statuses.get(i, "?") for i in request_ids}) or {"-"}
        )
        id_text = ",".join(request_ids) if request_ids else "-"

        if decision.reason == REASON_NO_REQUEST_ID:
            # counters.no_id was already incremented during extraction
            report.add(rp.Row(id_text, "-", subject, rp.ACTION_NO_ID))
            report.add_unmatched(
                str(message.get("receivedDateTime") or ""),
                subject,
                str(message.get("webLink") or ""),
            )
            continue

        if decision.reason in LEFT_ALONE:
            counter, action = LEFT_ALONE[decision.reason]
            setattr(counters, counter, getattr(counters, counter) + 1)
            report.add(rp.Row(id_text, status_text, subject, action))
            continue

        # decision.reason is REASON_DONE or REASON_WAITING
        if decision.reason == REASON_WAITING:
            counters.waiting += 1
        else:
            counters.done += 1
        proposed = proposed_categories(decision, current, config) or ()
        if not categories_changed(current, proposed):
            counters.unchanged += 1
            report.add(rp.Row(id_text, status_text, subject, rp.ACTION_UNCHANGED))
            continue

        if not apply:
            counters.would_apply += 1
            report.add(
                rp.Row(
                    id_text,
                    status_text,
                    subject,
                    rp.ACTION_WOULD_APPLY,
                    f"{current} -> {list(proposed)}",
                )
            )
            # recorded as well, so a scheduled dry-run leaves a trace of what
            # it would have done
            journal.append(
                format_change(
                    local_now(),
                    applied=False,
                    request_ids=request_ids,
                    status=status_text,
                    before=current,
                    after=proposed,
                    subject=subject,
                )
            )
            continue

        if limit is not None and changes_applied >= limit:
            counters.would_apply += 1
            report.add(rp.Row(id_text, status_text, subject, rp.ACTION_WOULD_APPLY, "limit"))
            continue

        try:
            graph.patch_categories(mail_id, proposed)
        except (HttpError, GraphAuthError) as exc:
            counters.errors += 1
            log.error("patch failed mail=%s error=%s", mail_id, exc)
            report.add(rp.Row(id_text, status_text, subject, rp.ACTION_ERROR, str(exc)))
            journal.append(
                format_error(
                    local_now(),
                    request_ids=request_ids,
                    subject=subject,
                    message=str(exc),
                    web_link=str(message.get("webLink") or ""),
                )
            )
            continue

        changes_applied += 1
        counters.applied += 1
        journal.append(
            format_change(
                local_now(),
                applied=True,
                request_ids=request_ids,
                status=status_text,
                before=current,
                after=proposed,
                subject=subject,
            )
        )
        state.record_change(
            mail_id=mail_id,
            internet_message_id=str(message.get("internetMessageId") or ""),
            subject=subject,
            request_ids=request_ids,
            statuses={i: statuses.get(i, "?") for i in request_ids},
            before=current,
            after=proposed,
        )
        log.info(
            "categories updated mail=%s folder=%s ids=%s before=%s after=%s",
            mail_id,
            message.get("_folder", "?"),
            id_text,
            current,
            list(proposed),
        )
        report.add(rp.Row(id_text, status_text, subject, rp.ACTION_APPLIED))


# ---------------------------------------------------------------------------
# run
# ---------------------------------------------------------------------------


def run(config: Config, args: argparse.Namespace, log: logging.Logger) -> int:
    started = time.monotonic()
    counters = rp.Counters()
    report = rp.Report(counters=counters)
    state = State.load(config.state_file)

    # Xurrent first: it is a single cheap GET and needs no browser, so a wrong
    # token, account or region is reported before a sign-in window opens.
    xurrent = XurrentClient(config, log)
    xurrent.ping()
    log.info(
        "xurrent reachable account=%s root=%s",
        config.xurrent_account,
        config.xurrent_api_root,
    )

    graph = GraphClient(config, log)
    username = graph.sign_in(interactive=not args.no_login)
    log.info("signed in as %s", username)
    validate_categories(graph, config, log)

    if args.check:
        print(
            "\nconfiguration OK\n"
            f"  mailbox ............ {username}\n"
            f"  scope .............. Inbox + subfolders\n"
            f"  lookback ........... "
            f"{f'{config.lookback_days} days' if config.lookback_days else 'no limit'}\n"
            f"  open categories .... "
            f"{', '.join(config.open_categories) or '(none - any mail without done)'}\n"
            f"  done category ...... {config.done_category}\n"
            f"  waiting category ... "
            + (
                f"{config.waiting_category} "
                f"({', '.join(sorted(config.waiting_statuses))})"
                if config.waiting_category
                else "(off)"
            )
            + "\n"
            f"  xurrent root ....... {config.xurrent_api_root}\n"
            f"  xurrent account .... {config.xurrent_account}\n"
            f"  done statuses ...... {', '.join(sorted(config.done_statuses))}\n"
            f"  multi-id policy .... {config.multi_request_policy}\n"
            f"  last run ........... {state.last_successful_run or 'never'}"
        )
        return EXIT_OK

    if args.explain:
        return explain(graph, xurrent, config, args.explain, log)

    days = args.days if args.days is not None else config.lookback_days
    since_iso: str | None = None
    if days > 0:
        since = datetime.now(timezone.utc) - timedelta(days=days)
        since_iso = since.strftime("%Y-%m-%dT%H:%M:%SZ")

    # saved even when the run dies half way: mails may already have been
    # changed in Outlook, and the audit trail has to say so
    try:
        candidates = collect_candidates(graph, config, since_iso, counters, log)
        resolve_request_ids(
            graph, config, state, candidates, not args.no_cache, counters, log
        )

        all_ids = sorted({i for m in candidates for i in (m.get("_request_ids") or ())})
        counters.requests_queried = len(all_ids)
        records = xurrent.fetch_requests(all_ids)
        counters.requests_readable = len(records)

        journal = DailyLog(config.daily_log_dir)
        apply_decisions(
            graph,
            config,
            state,
            candidates,
            status_map(records),
            completion_times(records),
            journal,
            apply=args.apply,
            limit=args.limit,
            counters=counters,
            report=report,
            log=log,
        )

        duration = time.monotonic() - started
        print()
        print(rp.render_table(report.rows))
        print()
        print(rp.render_unmatched(report.unmatched))
        print(rp.render_summary(counters, apply=args.apply, duration=duration))
        log.info("run finished %s", counters.as_line())
        journal.append(
            format_run(
                local_now(), apply=args.apply, counters=counters.as_line(), duration=duration
            )
        )

        if since_iso is None:  # only a full scan knows which mails are gone
            state.prune_cache(cache_key(m) for m in candidates)
        exit_code = EXIT_PARTIAL if counters.errors else EXIT_OK
        if exit_code == EXIT_OK:
            state.last_successful_run = utc_now_iso()
        return exit_code
    finally:
        state.save()


def main(argv: Sequence[str] | None = None) -> int:
    _reconfigure_stdout()
    args = build_arg_parser().parse_args(argv)

    try:
        load_env_file()
        config = Config.from_env()
    except ConfigError as exc:
        print(f"CONFIG ERROR: {exc}", file=sys.stderr)
        return EXIT_CONFIG

    log = setup_logging(config, args.verbose)
    log.info(
        "start apply=%s days=%s account=%s",
        args.apply,
        args.days if args.days is not None else config.lookback_days or "unlimited",
        config.xurrent_account,
    )

    try:
        return run(config, args, log)
    except ConfigError as exc:
        log.error("configuration error: %s", exc)
        print(f"CONFIG ERROR: {exc}", file=sys.stderr)
        return EXIT_CONFIG
    except (GraphAuthError, XurrentAuthError) as exc:
        log.error("authentication error: %s", exc)
        print(f"AUTH ERROR: {exc}", file=sys.stderr)
        return EXIT_CONFIG
    except HttpError as exc:
        log.error("unrecoverable http error: %s", exc)
        print(f"HTTP ERROR: {exc}", file=sys.stderr)
        return EXIT_PARTIAL
    except KeyboardInterrupt:  # pragma: no cover - interactive
        log.warning("interrupted by user")
        return EXIT_PARTIAL
    except Exception:  # last line of defence: pythonw.exe has no stderr, so without
        log.exception("unexpected error")  # this a crash would leave no trace at all
        print("UNEXPECTED ERROR - see the traceback in logs/sync.log", file=sys.stderr)
        return EXIT_PARTIAL
