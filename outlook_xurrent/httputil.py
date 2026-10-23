"""Shared HTTP plumbing: retries, rate limiting, error types."""

from __future__ import annotations

import logging
import time
from typing import Any, Final

import requests

MAX_SLEEP_SECONDS: Final[float] = 120.0
RETRY_STATUS: Final[frozenset[int]] = frozenset({429, 500, 502, 503, 504})


class HttpError(Exception):
    """A non-retryable HTTP error."""

    def __init__(self, status: int, url: str, body: str) -> None:
        super().__init__(f"HTTP {status} for {url}: {body[:400]}")
        self.status = status
        self.url = url
        self.body = body


class AuthError(HttpError):
    """401/403 - credentials, consent or scopes are wrong. Exit code 2."""


def parse_retry_after(response: requests.Response) -> float | None:
    """Read the retry-after header (both APIs send whole seconds)."""
    raw = response.headers.get("retry-after") or response.headers.get("Retry-After")
    if not raw:
        return None
    try:
        return max(0.0, float(raw.strip()))
    except ValueError:
        return None


def request_with_retry(
    session: requests.Session,
    method: str,
    url: str,
    *,
    timeout: int,
    max_retries: int,
    logger: logging.Logger,
    **kwargs: Any,
) -> requests.Response:
    """Perform a request, honouring 429/Retry-After and retrying 5xx.

    Raises AuthError on 401/403 and HttpError on any other 4xx.
    """
    last_error: Exception | None = None

    for attempt in range(1, max_retries + 1):
        try:
            response = session.request(method, url, timeout=timeout, **kwargs)
        except requests.RequestException as exc:
            last_error = exc
            sleep_for = min(2.0 ** (attempt - 1), MAX_SLEEP_SECONDS)
            logger.warning(
                "http transport error attempt=%d/%d sleep=%.1fs url=%s error=%s",
                attempt,
                max_retries,
                sleep_for,
                url,
                exc,
            )
            if attempt == max_retries:
                break
            time.sleep(sleep_for)
            continue

        if response.status_code in RETRY_STATUS:
            retry_after = parse_retry_after(response)
            sleep_for = min(
                retry_after if retry_after is not None else 2.0 ** (attempt - 1),
                MAX_SLEEP_SECONDS,
            )
            logger.warning(
                "http %d attempt=%d/%d sleep=%.1fs url=%s",
                response.status_code,
                attempt,
                max_retries,
                sleep_for,
                url,
            )
            if attempt == max_retries:
                raise HttpError(response.status_code, url, response.text)
            time.sleep(sleep_for)
            continue

        if response.status_code in (401, 403):
            raise AuthError(response.status_code, url, response.text)
        if response.status_code >= 400:
            raise HttpError(response.status_code, url, response.text)
        return response

    raise HttpError(0, url, f"request failed after {max_retries} attempts: {last_error}")
