"""Pure request-ID extraction. No I/O, no network, no logging.

Extraction order (first hit wins), as specified:

1. ``subject``     - subject regex, anchored on a literal '#'
2. ``body-link``   - request links in the raw body: https://<host>/requests/<id>
3. ``body-hash``   - '#<id>' in the HTML-stripped body text
4. ``header``      - internet message headers

Why the '#' anchor matters: a real notification body reads

    The following request from 654321-BEISPIELKUNDE SUEDTIROL was
    assigned to you:
    Request #80000001 - WG: Zugangsdaten Emailadresse

A naive ``\\d{6,}`` would match the customer number 654321 instead of the
request. Every pattern here therefore requires either a '#' or a '/requests/'
path segment in front of the digits.
"""

from __future__ import annotations

import html
import re
from dataclasses import dataclass
from typing import Iterable, Mapping, Sequence

SOURCE_SUBJECT = "subject"
SOURCE_BODY_LINK = "body-link"
SOURCE_BODY_HASH = "body-hash"
SOURCE_HEADER = "header"
SOURCE_NONE = "none"

#: Headers that may carry the request id. Names are compared case-insensitively.
_ID_BEARING_HEADERS = ("message-id", "references", "in-reply-to")

_TAG_RE = re.compile(r"<(script|style)\b[^>]*>.*?</\1>", re.IGNORECASE | re.DOTALL)
# [^<>] instead of [^>]: a body full of '<' without '>' would otherwise be quadratic
_ANY_TAG_RE = re.compile(r"<[^<>]+>")


@dataclass(frozen=True)
class Extraction:
    """Result of parsing one mail."""

    request_ids: tuple[str, ...]
    source: str

    @property
    def found(self) -> bool:
        return bool(self.request_ids)


def strip_html(body: str) -> str:
    """Reduce an HTML body to plain text.

    Tags are replaced by a space, so attribute values (styles, hex colours,
    tracking parameters) can never leak into the text that is scanned for
    '#<id>'.
    """
    if not body:
        return ""
    text = _TAG_RE.sub(" ", body)
    text = _ANY_TAG_RE.sub(" ", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def compile_subject_pattern(pattern: str) -> re.Pattern[str]:
    return re.compile(pattern, re.IGNORECASE)


def compile_header_pattern(pattern: str) -> re.Pattern[str]:
    return re.compile(pattern, re.IGNORECASE)


def compile_mention_pattern(pattern: str) -> re.Pattern[str] | None:
    """Subject pattern of an @-mention notification. Empty turns it off."""
    return re.compile(pattern, re.IGNORECASE) if pattern else None


def is_mention(subject: str, pattern: re.Pattern[str] | None) -> bool:
    """True when the subject is an @-mention notification."""
    return bool(pattern and pattern.search(subject or ""))


def compile_link_pattern(instance_hosts: str | Sequence[str] = "") -> re.Pattern[str]:
    """Build the request-link pattern.

    ``instance_hosts`` is one host or several (comma-separated string or a
    sequence). The pattern is pinned to them, so a forwarded link to a
    different Xurrent account cannot be picked up. With none given it accepts
    any ``*.xurrent.com`` host.

    Two path shapes are accepted, because the agent UI and the customer
    self-service portal link the same request differently:

        https://<instance>/requests/80000008
        https://<portal>/self-service/requests/80000008
    """
    raw = instance_hosts.split(",") if isinstance(instance_hosts, str) else list(instance_hosts)
    hosts = [
        re.escape(h.strip().rstrip("/").split("://", 1)[-1]) for h in raw if h and h.strip()
    ]
    # without a host list: xurrent.com / xurrent.qa and their subdomains, but not
    # look-alikes such as evilxurrent.com
    host_part = (
        "(?:" + "|".join(hosts) + ")" if hosts else r"(?:[A-Za-z0-9\-]+\.)*xurrent\.(?:com|qa)"
    )
    return re.compile(
        rf"https?://{host_part}/(?:self-service/)?requests/(\d+)", re.IGNORECASE
    )


def _dedupe(values: Iterable[str]) -> tuple[str, ...]:
    seen: dict[str, None] = {}
    for value in values:
        normalised = value.lstrip("0") or value
        seen.setdefault(normalised, None)
    return tuple(seen)


def extract_from_subject(subject: str, pattern: re.Pattern[str]) -> tuple[str, ...]:
    if not subject:
        return ()
    return _dedupe(pattern.findall(subject))


def extract_from_body_links(body: str, pattern: re.Pattern[str]) -> tuple[str, ...]:
    """Scan the *raw* body, so ids inside an <a href="..."> are found too."""
    if not body:
        return ()
    return _dedupe(pattern.findall(body))


def extract_from_body_text(body: str, pattern: re.Pattern[str]) -> tuple[str, ...]:
    if not body:
        return ()
    return _dedupe(pattern.findall(strip_html(body)))


def extract_from_headers(
    headers: Sequence[Mapping[str, str]], pattern: re.Pattern[str]
) -> tuple[str, ...]:
    """Scan internetMessageHeaders.

    Two sources are considered: any ``x-4me-*`` / ``x-xurrent-*`` header (digits
    anywhere in the value) and the RFC5322 id headers (pattern-matched).
    """
    found: list[str] = []
    for header in headers or ():
        name = str(header.get("name", "")).lower()
        value = str(header.get("value", ""))
        if not value:
            continue
        if name.startswith("x-4me-") or name.startswith("x-xurrent-"):
            found.extend(re.findall(r"(\d{4,})", value))
        elif name in _ID_BEARING_HEADERS:
            found.extend(pattern.findall(value))
    return _dedupe(found)


def extract_request_ids(
    *,
    subject: str,
    body: str,
    headers: Sequence[Mapping[str, str]] = (),
    subject_pattern: re.Pattern[str],
    link_pattern: re.Pattern[str],
    hash_pattern: re.Pattern[str],
    header_pattern: re.Pattern[str],
) -> Extraction:
    """Run all extraction steps in order and return the first hit."""
    for ids, source in (
        (extract_from_subject(subject, subject_pattern), SOURCE_SUBJECT),
        (extract_from_body_links(body, link_pattern), SOURCE_BODY_LINK),
        (extract_from_body_text(body, hash_pattern), SOURCE_BODY_HASH),
        (extract_from_headers(headers, header_pattern), SOURCE_HEADER),
    ):
        if ids:
            return Extraction(request_ids=ids, source=source)
    return Extraction(request_ids=(), source=SOURCE_NONE)
