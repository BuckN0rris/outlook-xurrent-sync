"""Console report: result table, summary counters, unmatched-mail list."""

from __future__ import annotations

from dataclasses import dataclass, field

ACTION_APPLIED = "APPLIED"
ACTION_WOULD_APPLY = "WOULD-APPLY"
ACTION_UNCHANGED = "UNCHANGED"
ACTION_STILL_OPEN = "STILL-OPEN"
ACTION_MENTION = "MENTION"
ACTION_AFTER_CLOSE = "AFTER-CLOSE"
ACTION_UNRESOLVED = "UNRESOLVED"
ACTION_NO_ID = "NO-ID"
ACTION_ERROR = "ERROR"


@dataclass
class Counters:
    """Everything the run log reports."""

    scanned: int = 0
    candidates: int = 0
    id_found: int = 0
    no_id: int = 0
    done: int = 0
    waiting: int = 0
    mentions: int = 0
    still_open: int = 0
    after_completion: int = 0
    unresolved: int = 0
    applied: int = 0
    would_apply: int = 0
    unchanged: int = 0
    errors: int = 0
    requests_queried: int = 0
    requests_readable: int = 0

    def as_line(self) -> str:
        return (
            f"scanned={self.scanned} candidates={self.candidates} "
            f"id_found={self.id_found} no_id={self.no_id} done={self.done} "
            f"waiting={self.waiting} mentions={self.mentions} "
            f"still_open={self.still_open} after_completion={self.after_completion} "
            f"unresolved={self.unresolved} "
            f"applied={self.applied} would_apply={self.would_apply} "
            f"unchanged={self.unchanged} errors={self.errors}"
        )


@dataclass
class Row:
    request_ids: str
    status: str
    subject: str
    action: str
    detail: str = ""


@dataclass
class Report:
    rows: list[Row] = field(default_factory=list)
    unmatched: list[tuple[str, str, str]] = field(default_factory=list)
    counters: Counters = field(default_factory=Counters)

    def add(self, row: Row) -> None:
        self.rows.append(row)

    def add_unmatched(self, received: str, subject: str, web_link: str) -> None:
        self.unmatched.append((received, subject, web_link))


def _clip(value: str, width: int) -> str:
    value = " ".join((value or "").split())
    if len(value) <= width:
        return value.ljust(width)
    return value[: width - 3] + "..."


def render_table(rows: list[Row]) -> str:
    if not rows:
        return "(no mails matched the scope)"

    headers = ("REQUEST", "STATUS", "ACTION", "SUBJECT")
    widths = (24, 20, 11, 52)
    order = {
        ACTION_APPLIED: 0,
        ACTION_WOULD_APPLY: 1,
        ACTION_ERROR: 2,
        ACTION_UNRESOLVED: 3,
        ACTION_UNCHANGED: 4,
        ACTION_STILL_OPEN: 5,
        ACTION_MENTION: 6,
        ACTION_AFTER_CLOSE: 7,
        ACTION_NO_ID: 8,
    }
    ordered = sorted(rows, key=lambda r: (order.get(r.action, 9), r.request_ids))

    lines = [
        "  ".join(_clip(h, w) for h, w in zip(headers, widths)).rstrip(),
        "  ".join("-" * w for w in widths),
    ]
    for row in ordered:
        lines.append(
            "  ".join(
                (
                    _clip(row.request_ids, widths[0]),
                    _clip(row.status, widths[1]),
                    _clip(row.action, widths[2]),
                    _clip(row.subject, widths[3]),
                )
            ).rstrip()
        )
    return "\n".join(lines)


def render_unmatched(unmatched: list[tuple[str, str, str]]) -> str:
    if not unmatched:
        return "No mails without a request id."
    lines = [
        f"Mails with no request id ({len(unmatched)}) - review these by hand:",
    ]
    for received, subject, web_link in unmatched:
        lines.append(f"  {received[:19]}  {' '.join((subject or '(no subject)').split())}")
        if web_link:
            lines.append(f"      {web_link}")
    return "\n".join(lines)


def render_summary(counters: Counters, *, apply: bool, duration: float) -> str:
    mode = "APPLY" if apply else "DRY-RUN"
    return (
        f"\nmode={mode}  duration={duration:.1f}s\n"
        f"  mails scanned .............. {counters.scanned}\n"
        f"  candidates (open, not done)  {counters.candidates}\n"
        f"  request id found ........... {counters.id_found}\n"
        f"  no request id .............. {counters.no_id}\n"
        f"  requests queried / readable  {counters.requests_queried} / "
        f"{counters.requests_readable}\n"
        f"  done in Xurrent ............ {counters.done}\n"
        f"  waiting in Xurrent ......... {counters.waiting}\n"
        f"  mention, kept open ......... {counters.mentions}\n"
        f"  still open ................. {counters.still_open}\n"
        f"  note after closure ......... {counters.after_completion}\n"
        f"  unresolved (403/404) ....... {counters.unresolved}\n"
        f"  categories applied ......... {counters.applied}\n"
        f"  categories would apply ..... {counters.would_apply}\n"
        f"  already correct ............ {counters.unchanged}\n"
        f"  errors ..................... {counters.errors}"
    )
