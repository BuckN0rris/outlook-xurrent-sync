"""Local state: extraction cache, audit trail, last successful run.

One JSON file, written atomically. The cache is keyed by internetMessageId
(stable for the lifetime of the mail, unlike the Graph message id, which
changes when a mail is moved between folders).
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Sequence

LOG = logging.getLogger(__name__)

STATE_VERSION = 1
MAX_AUDIT_ENTRIES = 5000


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


@dataclass
class State:
    path: Path
    last_successful_run: str | None = None
    request_id_cache: dict[str, dict[str, Any]] = field(default_factory=dict)
    audit: list[dict[str, Any]] = field(default_factory=list)

    # -- persistence -------------------------------------------------------

    @classmethod
    def load(cls, path: Path) -> "State":
        if not path.is_file():
            return cls(path=path)
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            version = int(raw.get("version", 0))
        except (OSError, ValueError, AttributeError) as exc:
            LOG.warning("state file unreadable, starting fresh: %s", exc)
            return cls._quarantine(path)
        if version != STATE_VERSION:
            LOG.warning("state file has version %s, starting fresh", version)
            return cls._quarantine(path)
        return cls(
            path=path,
            last_successful_run=raw.get("last_successful_run"),
            request_id_cache=dict(raw.get("request_id_cache", {})),
            audit=list(raw.get("audit", [])),
        )

    @classmethod
    def _quarantine(cls, path: Path) -> "State":
        """Set a state file we cannot use aside, so the next save() keeps the audit trail."""
        try:
            path.replace(path.with_suffix(path.suffix + ".bad"))
        except OSError as exc:
            LOG.warning("could not set the state file aside: %s", exc)
        return cls(path=path)

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": STATE_VERSION,
            "last_successful_run": self.last_successful_run,
            "request_id_cache": self.request_id_cache,
            "audit": self.audit[-MAX_AUDIT_ENTRIES:],
        }
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        tmp.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        tmp.replace(self.path)

    # -- extraction cache --------------------------------------------------

    def cached_ids(self, key: str) -> tuple[tuple[str, ...], str] | None:
        entry = self.request_id_cache.get(key)
        if not entry:
            return None
        ids = tuple(str(i) for i in entry.get("request_ids", []))
        return ids, str(entry.get("source", "cache"))

    def remember_ids(self, key: str, request_ids: Sequence[str], source: str) -> None:
        self.request_id_cache[key] = {
            "request_ids": list(request_ids),
            "source": source,
            "cached_at": utc_now_iso(),
        }

    def prune_cache(self, keep: Iterable[str]) -> None:
        """Drop cache entries of mails that are no longer candidates (done mails)."""
        keep = set(keep)
        self.request_id_cache = {k: v for k, v in self.request_id_cache.items() if k in keep}

    # -- audit trail -------------------------------------------------------

    def record_change(
        self,
        *,
        mail_id: str,
        internet_message_id: str,
        subject: str,
        request_ids: Sequence[str],
        statuses: dict[str, str],
        before: Sequence[str],
        after: Sequence[str],
    ) -> None:
        self.audit.append(
            {
                "ts": utc_now_iso(),
                "mail_id": mail_id,
                "internet_message_id": internet_message_id,
                "subject": subject,
                "request_ids": list(request_ids),
                "statuses": statuses,
                "categories_before": list(before),
                "categories_after": list(after),
            }
        )
