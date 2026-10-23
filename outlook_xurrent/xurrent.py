"""Xurrent REST client - strictly read-only.

Only GET is possible: ``_call`` refuses anything else, so no code path can
create, update or comment on a record. Batch lookups use the documented "in"
filter (``?id=1,2,3``) plus ``fields`` to keep payloads small.

Requests that the token may not read simply do not come back in the batch
response; those ids are reported as unresolved rather than guessed at.
"""

from __future__ import annotations

import logging
from typing import Any, Iterable, Iterator, Sequence

import requests

from .config import (
    HTTP_MAX_RETRIES,
    HTTP_TIMEOUT_SECONDS,
    XURRENT_BATCH_SIZE,
    Config,
)
from .httputil import AuthError, request_with_retry

LOG = logging.getLogger(__name__)

REQUEST_FIELDS = "id,subject,status,completed_at,team,member"


class XurrentAuthError(Exception):
    """Token, scope, account header or region is wrong. Exit code 2."""


def chunked(values: Sequence[str], size: int) -> Iterator[list[str]]:
    for start in range(0, len(values), size):
        yield list(values[start : start + size])


class XurrentClient:
    def __init__(self, config: Config, logger: logging.Logger | None = None) -> None:
        self._config = config
        self._log = logger or LOG
        self._session = requests.Session()
        self._session.headers.update(
            {
                "Authorization": f"Bearer {config.xurrent_token}",
                "X-Xurrent-Account": config.xurrent_account,
                "Accept": "application/json",
            }
        )

    # -- low level ---------------------------------------------------------

    def _call(self, url: str, params: dict[str, Any] | None = None) -> requests.Response:
        """GET only. This client never writes to Xurrent."""
        try:
            return request_with_retry(
                self._session,
                "GET",
                url,
                timeout=HTTP_TIMEOUT_SECONDS,
                max_retries=HTTP_MAX_RETRIES,
                logger=self._log,
                params=params,
            )
        except AuthError as exc:
            if exc.status == 401:
                raise XurrentAuthError(
                    "Xurrent returned 401 Unauthorized. XURRENT_TOKEN is invalid, expired "
                    "or revoked. Generate a new Personal Access Token."
                ) from exc
            raise XurrentAuthError(
                "Xurrent returned 403 Forbidden. Check that the Personal Access Token has "
                f"Allow / Request:Read and that XURRENT_ACCOUNT={self._config.xurrent_account!r} "
                "matches the account condition of the token."
            ) from exc

    def _get_collection(
        self, path: str, params: dict[str, Any]
    ) -> Iterator[dict[str, Any]]:
        """GET a collection and follow the Link header (rel="next") cursor."""
        url: str | None = f"{self._config.xurrent_api_root}{path}"
        next_params: dict[str, Any] | None = params
        while url:
            response = self._call(url, next_params)
            payload = response.json()
            if isinstance(payload, dict):  # single record or error shape
                payload = [payload]
            for item in payload:
                yield item
            url = response.links.get("next", {}).get("url")
            next_params = None  # the cursor url already carries every parameter

    # -- reads -------------------------------------------------------------

    def ping(self) -> None:
        """Cheapest possible read, used to validate token, account and region."""
        self._call(
            f"{self._config.xurrent_api_root}/requests",
            {"per_page": 1, "fields": "id"},
        )

    def fetch_requests(self, request_ids: Iterable[str]) -> dict[str, dict[str, Any]]:
        """Return {request_id: record} for every id that could be read."""
        ids = sorted({str(i) for i in request_ids if str(i).strip()}, key=_sort_key)
        if not ids:
            return {}

        found: dict[str, dict[str, Any]] = {}
        for batch in chunked(ids, XURRENT_BATCH_SIZE):
            self._log.info("xurrent batch size=%d", len(batch))
            for record in self._get_collection(
                "/requests",
                {
                    "id": ",".join(batch),
                    "fields": REQUEST_FIELDS,
                    "per_page": 100,
                },
            ):
                record_id = record.get("id")
                if record_id is None:
                    continue
                found[str(record_id)] = record

        missing = [i for i in ids if i not in found]
        if missing:
            self._log.warning(
                "xurrent requests not readable (404/403/unknown) count=%d ids=%s",
                len(missing),
                ",".join(missing[:20]) + ("..." if len(missing) > 20 else ""),
            )
        return found


def _sort_key(value: str) -> tuple[int, str]:
    return (0, value.zfill(12)) if value.isdigit() else (1, value)


def status_map(records: dict[str, dict[str, Any]]) -> dict[str, str]:
    """Reduce records to {id: status}."""
    return {
        request_id: str(record.get("status", "")).lower()
        for request_id, record in records.items()
    }
