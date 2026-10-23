"""Unit tests for request-ID extraction. Fixtures only, no network."""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any

import pytest

from outlook_xurrent.config import (
    DEFAULT_HEADER_REGEX,
    DEFAULT_MENTION_REGEX,
    DEFAULT_SUBJECT_REGEX,
)
from outlook_xurrent.extract import (
    SOURCE_BODY_LINK,
    SOURCE_HEADER,
    SOURCE_NONE,
    SOURCE_SUBJECT,
    compile_header_pattern,
    compile_link_pattern,
    compile_mention_pattern,
    compile_subject_pattern,
    extract_from_body_links,
    extract_request_ids,
    is_mention,
    strip_html,
)

FIXTURES = Path(__file__).parent / "fixtures"
INSTANCE_HOST = "example-it.xurrent.com"


def load(name: str) -> dict[str, Any]:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def extract(mail: dict[str, Any], instance_host: str = INSTANCE_HOST):
    return extract_request_ids(
        subject=mail.get("subject", ""),
        body=(mail.get("body") or {}).get("content", ""),
        headers=mail.get("internetMessageHeaders", ()),
        subject_pattern=compile_subject_pattern(DEFAULT_SUBJECT_REGEX),
        link_pattern=compile_link_pattern(instance_host),
        hash_pattern=compile_subject_pattern(DEFAULT_SUBJECT_REGEX),
        header_pattern=compile_header_pattern(DEFAULT_HEADER_REGEX),
    )


# --- the real notification mail -------------------------------------------


def test_real_notification_is_resolved_via_body_link() -> None:
    result = extract(load("mail_assigned.json"))
    assert result.request_ids == ("80000001",)
    assert result.source == SOURCE_BODY_LINK


def test_customer_number_is_never_taken_for_a_request_id() -> None:
    """654321-BEISPIELKUNDE sits right next to the real id in the body."""
    result = extract(load("mail_assigned.json"))
    assert "654321" not in result.request_ids


def test_hex_colour_and_url_parameters_are_stripped_before_text_scan() -> None:
    body = (load("mail_assigned.json")["body"])["content"]
    text = strip_html(body)
    assert "123456" not in text  # from <style>.brand{color:#123456}</style>
    assert "384000" not in text  # from <img src="...?w=384000">
    assert "Request #80000001" in text


# --- precedence ------------------------------------------------------------


def test_subject_wins_over_body() -> None:
    mail = load("mail_subject_id.json")
    result = extract(mail)
    assert result.source == SOURCE_SUBJECT
    assert result.request_ids == ("80000003",)


def test_header_is_used_when_subject_and_body_are_silent() -> None:
    result = extract(load("mail_header_only.json"))
    assert result.source == SOURCE_HEADER
    assert result.request_ids == ("70000001",)


def test_plain_colleague_mail_yields_nothing() -> None:
    result = extract(load("mail_plain_no_id.json"))
    assert result.request_ids == ()
    assert result.source == SOURCE_NONE


# --- multiple ids ----------------------------------------------------------


def test_multiple_links_are_deduped_in_order() -> None:
    result = extract(load("mail_multi_request.json"))
    assert result.request_ids == ("80000001", "80000002")
    assert result.source == SOURCE_BODY_LINK


# --- host pinning ----------------------------------------------------------


def test_link_pattern_ignores_a_foreign_instance() -> None:
    body = '<a href="https://other-tenant.xurrent.com/requests/4242">x</a>'
    pinned = compile_link_pattern(INSTANCE_HOST)
    assert extract_from_body_links(body, pinned) == ()


def test_link_pattern_without_host_accepts_any_xurrent_instance() -> None:
    body = '<a href="https://other-tenant.xurrent.com/requests/4242">x</a>'
    generic = compile_link_pattern("")
    assert extract_from_body_links(body, generic) == ("4242",)


def test_instance_host_may_be_configured_with_scheme() -> None:
    body = '<a href="https://example-it.xurrent.com/requests/7">x</a>'
    pattern = compile_link_pattern("https://example-it.xurrent.com/")
    assert extract_from_body_links(body, pattern) == ("7",)


# --- helpers ---------------------------------------------------------------


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("", ""),
        ("<p>a</p><p>b</p>", "a b"),
        ("<p>a&nbsp;&amp;&nbsp;b</p>", "a & b"),
        ("<script>var x=999999;</script>text", "text"),
    ],
)
def test_strip_html(raw: str, expected: str) -> None:
    assert strip_html(raw) == expected


def test_no_body_and_no_headers_is_not_an_error() -> None:
    result = extract({"subject": "Re: Urlaubsantrag", "body": {}, "internetMessageHeaders": []})
    assert result.request_ids == ()
    assert result.source == SOURCE_NONE


# --- OData literal escaping for the server-side category filter ------------


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("Nicht erledigt", "Nicht erledigt"),
        ("O'Brien", "O''Brien"),
        ("a'b'c", "a''b''c"),
        ("", ""),
    ],
)
def test_odata_literal(value: str, expected: str) -> None:
    from outlook_xurrent.graph import odata_literal

    assert odata_literal(value) == expected


# --- customer self-service portal links ------------------------------------
#
# The agent UI and the customer portal link the same request differently:
#   https://<instance>/requests/80000008
#   https://<portal>/self-service/requests/80000008


PORTAL_HOST = "example-customers.xurrent.com"
BOTH_HOSTS = f"{INSTANCE_HOST},{PORTAL_HOST}"


def test_self_service_path_is_recognised() -> None:
    body = f'<a href="https://{PORTAL_HOST}/self-service/requests/80000001">x</a>'
    assert extract_from_body_links(body, compile_link_pattern(BOTH_HOSTS)) == ("80000001",)


def test_portal_host_must_be_configured() -> None:
    """Pinning stays strict: an unlisted host is still ignored."""
    body = f'<a href="https://{PORTAL_HOST}/self-service/requests/80000001">x</a>'
    assert extract_from_body_links(body, compile_link_pattern(INSTANCE_HOST)) == ()


def test_self_service_path_on_the_agent_host() -> None:
    body = f'<a href="https://{INSTANCE_HOST}/self-service/requests/7">x</a>'
    assert extract_from_body_links(body, compile_link_pattern(BOTH_HOSTS)) == ("7",)


def test_both_link_shapes_in_one_body_dedupe_to_one_id() -> None:
    body = (
        f'<a href="https://{INSTANCE_HOST}/requests/80000001">agent</a>'
        f'<a href="https://{PORTAL_HOST}/self-service/requests/80000001">portal</a>'
    )
    assert extract_from_body_links(body, compile_link_pattern(BOTH_HOSTS)) == ("80000001",)


def test_host_list_accepts_a_sequence_too() -> None:
    body = f'<a href="https://{PORTAL_HOST}/self-service/requests/42">x</a>'
    pattern = compile_link_pattern([INSTANCE_HOST, PORTAL_HOST])
    assert extract_from_body_links(body, pattern) == ("42",)


def test_foreign_host_is_still_rejected_with_a_host_list() -> None:
    body = '<a href="https://other-tenant.xurrent.com/requests/4242">x</a>'
    assert extract_from_body_links(body, compile_link_pattern(BOTH_HOSTS)) == ()


def test_unrelated_path_is_not_matched() -> None:
    body = f'<a href="https://{INSTANCE_HOST}/workflows/80000001">x</a>'
    assert extract_from_body_links(body, compile_link_pattern(BOTH_HOSTS)) == ()


# --- mention notifications ----------------------------------------------------


@pytest.mark.parametrize(
    ("subject", "expected"),
    [
        ("👉You were mentioned in Request #80000014 - Account kompromittiert", True),
        ("you were MENTIONED in request #1234", True),
        ("Note added to request: Benutzerin Karin Huber - Phishing Link", False),
        ("📫 New assigned Request #80000015 - Berechtigungen auf Ordner", False),
        ("AW: you were mentionedx", False),
        ("", False),
    ],
)
def test_is_mention(subject: str, expected: bool) -> None:
    pattern = compile_mention_pattern(DEFAULT_MENTION_REGEX)
    assert is_mention(subject, pattern) is expected


def test_empty_mention_pattern_turns_detection_off() -> None:
    assert compile_mention_pattern("") is None
    assert is_mention("👉You were mentioned in Request #1", None) is False


# --- hardening ---------------------------------------------------------------


@pytest.mark.parametrize(
    "host", ["evilxurrent.com", "xurrent.com.evil.io", "evil.io/xurrent.com", "xurrent.org"]
)
def test_link_pattern_without_host_rejects_look_alike_domains(host: str) -> None:
    body = f'<a href="https://{host}/requests/4242">x</a>'
    assert extract_from_body_links(body, compile_link_pattern("")) == ()


def test_link_pattern_without_host_still_accepts_the_bare_domain() -> None:
    body = '<a href="https://xurrent.com/requests/4242">x</a>'
    assert extract_from_body_links(body, compile_link_pattern("")) == ("4242",)


def test_strip_html_stays_fast_on_a_body_of_open_brackets() -> None:
    started = time.monotonic()
    strip_html("<" * 200_000)
    assert time.monotonic() - started < 2
