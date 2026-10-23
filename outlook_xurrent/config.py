"""Configuration: fixed constants, .env loading, validation.

Only settings that genuinely differ between people or environments are
configurable. Operational knobs (page sizes, timeouts, retries) are constants,
because there is no good reason to tune them per run.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Final, Iterable

# ---------------------------------------------------------------------------
# fixed constants
# ---------------------------------------------------------------------------

GRAPH_RESOURCE: Final[str] = "https://graph.microsoft.com"
GRAPH_BASE: Final[str] = "https://graph.microsoft.com/v1.0"

#: Pre-registered Microsoft first-party public client ("Microsoft Graph Command
#: Line Tools"). Works without an own app registration, as long as the tenant
#: allows user consent. Override with GRAPH_CLIENT_ID when it does not.
DEFAULT_GRAPH_CLIENT_ID: Final[str] = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

GRAPH_PAGE_SIZE: Final[int] = 100  # messages per Graph page (max 1000)
XURRENT_BATCH_SIZE: Final[int] = 50  # ids per ?id=1,2,3 lookup (per_page max 100)
HTTP_TIMEOUT_SECONDS: Final[int] = 60
HTTP_MAX_RETRIES: Final[int] = 5

#: Request status enum as documented at developer.xurrent.com/v1/requests/
XURRENT_STATUSES: Final[frozenset[str]] = frozenset(
    {
        "declined",
        "on_backlog",
        "assigned",
        "accepted",
        "in_progress",
        "waiting_for",
        "waiting_for_customer",
        "reservation_pending",
        "workflow_pending",
        "project_pending",
        "completed",
    }
)

#: Subject and body-text pattern. Anchored on a literal '#' on purpose: the
#: notification body contains a customer code ("654321-BEISPIELKUNDE")
#: right next to the request id, and a bare \d{6,} would match that instead.
DEFAULT_SUBJECT_REGEX: Final[str] = r"#\s*(\d{4,})"

#: Header fallback: digits directly behind a "request" token, e.g. a Message-ID
#: like <request.80000001.12@example-it.xurrent.com>.
DEFAULT_HEADER_REGEX: Final[str] = r"requests?[._/-](\d{4,})"

#: Subject of the notification Xurrent sends when someone @-mentions you in a
#: note, e.g. "👉You were mentioned in Request #80000001 - ...". Such a mail
#: asks you something, so it never moves to the waiting category.
DEFAULT_MENTION_REGEX: Final[str] = r"\byou were mentioned\b"


class ConfigError(Exception):
    """Configuration is unusable. Mapped to exit code 2."""


# ---------------------------------------------------------------------------
# .env loading
# ---------------------------------------------------------------------------


def load_env_file(path: Path | None = None) -> Path | None:
    """Load .env into os.environ without overriding real environment vars."""
    try:
        from dotenv import load_dotenv
    except ImportError as exc:  # pragma: no cover - dependency check
        raise ConfigError(
            "python-dotenv is not installed. Run: pip install -r requirements.txt"
        ) from exc

    candidate = path or Path(__file__).resolve().parent.parent / ".env"
    if candidate.is_file():
        load_dotenv(candidate, override=False)
        return candidate
    return None


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _str(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def _bool(name: str, default: bool) -> bool:
    raw = _str(name)
    if not raw:
        return default
    if raw.lower() in {"1", "true", "yes", "y", "on"}:
        return True
    if raw.lower() in {"0", "false", "no", "n", "off"}:
        return False
    raise ConfigError(f"{name} must be true or false, got {raw!r}")


def _int(name: str, default: int, minimum: int, maximum: int) -> int:
    raw = _str(name)
    if not raw:
        return default
    try:
        value = int(raw)
    except ValueError as exc:
        raise ConfigError(f"{name} must be an integer, got {raw!r}") from exc
    if not minimum <= value <= maximum:
        raise ConfigError(f"{name} must be between {minimum} and {maximum}, got {value}")
    return value


def _csv(name: str, default: Iterable[str] = ()) -> tuple[str, ...]:
    raw = _str(name)
    if not raw:
        return tuple(default)
    return tuple(part.strip() for part in raw.split(",") if part.strip())


def _statuses(name: str, default: Iterable[str]) -> frozenset[str]:
    """Read a list of Xurrent statuses and reject anything not in the enum."""
    statuses = frozenset(s.lower() for s in _csv(name, default))
    unknown = sorted(statuses - XURRENT_STATUSES)
    if unknown:
        raise ConfigError(
            f"{name} contains unknown Xurrent status(es): "
            + ", ".join(unknown)
            + ". Valid values: "
            + ", ".join(sorted(XURRENT_STATUSES))
        )
    return statuses


def _qualify_scope(scope: str) -> str:
    """Turn 'Mail.ReadWrite' into a fully qualified Graph scope URI."""
    return scope if "://" in scope else f"{GRAPH_RESOURCE}/{scope}"


# ---------------------------------------------------------------------------
# config object
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Config:
    """Validated runtime configuration."""

    # Microsoft Graph
    graph_client_id: str
    graph_tenant_id: str
    graph_scopes: tuple[str, ...]
    graph_token_cache: Path

    # Outlook
    open_categories: tuple[str, ...]
    done_category: str
    waiting_category: str
    lookback_days: int
    validate_categories: bool

    # Xurrent
    xurrent_api_root: str
    xurrent_account: str
    xurrent_token: str = field(repr=False)  # never let the PAT end up in a log or traceback
    xurrent_instance_hosts: tuple[str, ...]
    done_statuses: frozenset[str]
    waiting_statuses: frozenset[str]
    multi_request_policy: str
    after_close_grace_seconds: int

    # ID extraction
    subject_regex: str
    header_regex: str
    mention_regex: str

    # local files
    state_file: Path
    log_dir: Path
    daily_log_dir: Path

    @classmethod
    def from_env(cls) -> "Config":
        root = Path(__file__).resolve().parent.parent

        def _path(name: str, default: str) -> Path:
            candidate = Path(_str(name, default)).expanduser()
            return candidate if candidate.is_absolute() else root / candidate

        scopes = tuple(
            _qualify_scope(s)
            for s in _csv("GRAPH_SCOPES", ("Mail.ReadWrite", "MailboxSettings.Read"))
        )

        policy = _str("MULTI_REQUEST_POLICY", "all").lower()
        if policy not in {"all", "any"}:
            raise ConfigError("MULTI_REQUEST_POLICY must be 'all' or 'any'")

        done_statuses = _statuses("DONE_STATUSES", ("completed",))

        done_category = _str("DONE_CATEGORY")
        if not done_category:
            raise ConfigError("DONE_CATEGORY is required (exact Outlook category name)")
        open_categories = _csv("OPEN_CATEGORY")
        if any(c.casefold() == done_category.casefold() for c in open_categories):
            raise ConfigError(
                "DONE_CATEGORY must not also appear in OPEN_CATEGORY"
            )

        waiting_category = _str("WAITING_CATEGORY")
        waiting_statuses = _statuses(
            "WAITING_STATUSES", ("waiting_for", "waiting_for_customer")
        )
        if waiting_category:
            if waiting_category.casefold() == done_category.casefold():
                raise ConfigError("WAITING_CATEGORY must differ from DONE_CATEGORY")
            # a waiting mail has to stay a candidate, otherwise it would never be
            # picked up again once its request is completed
            if open_categories and waiting_category.casefold() not in {
                c.casefold() for c in open_categories
            }:
                raise ConfigError(
                    "WAITING_CATEGORY must also be listed in OPEN_CATEGORY, otherwise "
                    "a mail moved there is never set to done when its request closes"
                )
            overlap = sorted(waiting_statuses & done_statuses)
            if overlap:
                raise ConfigError(
                    "WAITING_STATUSES and DONE_STATUSES overlap: " + ", ".join(overlap)
                )
        else:
            waiting_statuses = frozenset()

        mention_regex = _str("MENTION_REGEX", DEFAULT_MENTION_REGEX)
        try:
            re.compile(mention_regex)
        except re.error as exc:
            raise ConfigError(
                f"MENTION_REGEX is not a valid regular expression: {exc}"
            ) from exc

        if not _str("XURRENT_TOKEN"):
            raise ConfigError(
                "XURRENT_TOKEN is required. Create a Personal Access Token in Xurrent "
                "(avatar > My Profile > Personal Access Tokens) with Allow / Request:Read."
            )
        if not _str("XURRENT_ACCOUNT"):
            raise ConfigError(
                "XURRENT_ACCOUNT is required (the value sent as X-Xurrent-Account header)"
            )

        api_root = _str("XURRENT_API_ROOT", "https://api.xurrent.com/v1").rstrip("/")
        if not api_root.startswith("https://"):
            raise ConfigError("XURRENT_API_ROOT must start with https://")

        return cls(
            graph_client_id=_str("GRAPH_CLIENT_ID", DEFAULT_GRAPH_CLIENT_ID),
            graph_tenant_id=_str("GRAPH_TENANT_ID", "organizations"),
            graph_scopes=scopes,
            graph_token_cache=_path("GRAPH_TOKEN_CACHE", ".state/msal_cache.json"),
            open_categories=open_categories,
            done_category=done_category,
            waiting_category=waiting_category,
            # 0 = no date limit. The category filter keeps the scan cheap, and an
            # unhandled mail stays open for as long as it takes, so a short
            # window would silently hide the older backlog.
            lookback_days=_int("LOOKBACK_DAYS", 0, 0, 3650),
            validate_categories=_bool("VALIDATE_CATEGORIES", True),
            xurrent_api_root=api_root,
            xurrent_account=_str("XURRENT_ACCOUNT"),
            xurrent_token=_str("XURRENT_TOKEN"),
            xurrent_instance_hosts=_csv("XURRENT_INSTANCE_HOST"),
            done_statuses=done_statuses,
            waiting_statuses=waiting_statuses,
            multi_request_policy=policy,
            # A mail arriving within this window after completed_at still counts
            # as part of the closure. The default sits in a gap measured against
            # a real mailbox: of 30 mails that arrived after their request was
            # closed, 18 came within 96 minutes (the "completed" notification
            # after 12-75 s, closing notes within minutes) and the next one only
            # 28.5 hours later. Two hours therefore separates the closing act
            # from genuine follow-ups without sitting close to either side.
            after_close_grace_seconds=_int(
                "AFTER_CLOSE_GRACE_SECONDS", 7200, 0, 2_592_000
            ),
            subject_regex=_str("SUBJECT_REGEX", DEFAULT_SUBJECT_REGEX),
            header_regex=_str("HEADER_REGEX", DEFAULT_HEADER_REGEX),
            mention_regex=mention_regex,
            state_file=_path("STATE_FILE", ".state/state.json"),
            log_dir=_path("LOG_DIR", "logs"),
            daily_log_dir=_path("DAILY_LOG_DIR", "logs/daily"),
        )
