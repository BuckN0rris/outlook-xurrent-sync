"""Microsoft Graph client: delegated access to the signed-in user's mailbox.

Only four kinds of calls are made:
  * read the Inbox folder tree
  * read messages in those folders
  * read the mailbox category list (optional, needs MailboxSettings.Read)
  * PATCH the ``categories`` array of a single message

Nothing else is ever sent. No move, no delete, no flag, no reply, no forward.
"""

from __future__ import annotations

import logging
from typing import Any, Iterator, Sequence
from urllib.parse import quote

import msal
import requests

from .config import (
    GRAPH_BASE,
    GRAPH_PAGE_SIZE,
    HTTP_MAX_RETRIES,
    HTTP_TIMEOUT_SECONDS,
    Config,
)
from .httputil import AuthError, HttpError, request_with_retry

LOG = logging.getLogger(__name__)

MESSAGE_LIST_SELECT = "id,subject,receivedDateTime,categories,webLink,internetMessageId"
MESSAGE_DETAIL_SELECT = "id,subject,body,internetMessageHeaders"


class GraphAuthError(Exception):
    """Sign-in failed or the token is unusable. Exit code 2."""


def odata_literal(value: str) -> str:
    """Escape a string for an OData literal: single quotes are doubled."""
    return value.replace("'", "''")


class GraphClient:
    """Thin, explicit Graph wrapper. One session, one token, shared retries."""

    def __init__(self, config: Config, logger: logging.Logger | None = None) -> None:
        self._config = config
        self._log = logger or LOG
        self._session = requests.Session()
        self._token: str | None = None
        self._cache = msal.SerializableTokenCache()

    # -- authentication ----------------------------------------------------

    def _load_cache(self) -> None:
        path = self._config.graph_token_cache
        if path.is_file():
            try:
                self._cache.deserialize(path.read_text(encoding="utf-8"))
            except (OSError, ValueError) as exc:
                self._log.warning("token cache unreadable, ignoring it: %s", exc)

    def _save_cache(self) -> None:
        if not self._cache.has_state_changed:
            return
        path = self._config.graph_token_cache
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(self._cache.serialize(), encoding="utf-8")
        tmp.replace(path)

    def sign_in(self, interactive: bool = True) -> str:
        """Acquire a token - silently when the on-disk cache still holds one.

        With ``interactive=False`` a missing or expired cached token is an error
        instead of a browser window, which is what an unattended run needs: a
        sign-in window nobody sees would block the scheduled task for good.

        Returns the signed-in user name, taken from the MSAL account, so no
        extra Graph permission is needed just to know who we are.
        """
        self._load_cache()
        app = msal.PublicClientApplication(
            self._config.graph_client_id,
            authority=f"https://login.microsoftonline.com/{self._config.graph_tenant_id}",
            token_cache=self._cache,
        )
        scopes = list(self._config.graph_scopes)

        result: dict[str, Any] | None = None
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            if result:
                self._log.info("graph token from cache")
        if not result and not interactive:
            raise GraphAuthError(
                "No valid cached Graph token and --no-login is set. Run "
                "'python sync.py --check' once in a terminal to sign in again."
            )
        if not result:
            self._log.info("opening browser for interactive sign-in")
            result = app.acquire_token_interactive(scopes=scopes)

        self._save_cache()

        if not result or "access_token" not in result:
            description = str((result or {}).get("error_description", "no token returned"))
            raise GraphAuthError(
                "Microsoft Graph sign-in failed: "
                + description.splitlines()[0]
                + "\nIf the tenant blocks user consent, have an admin approve the "
                "delegated scopes, or set GRAPH_CLIENT_ID to your own app registration."
            )

        self._token = result["access_token"]
        self._session.headers.update(
            {"Authorization": f"Bearer {self._token}", "Accept": "application/json"}
        )
        accounts = app.get_accounts()
        return str(accounts[0].get("username", "unknown")) if accounts else "unknown"

    # -- low level ---------------------------------------------------------

    def _call(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        if self._token is None:
            raise GraphAuthError("sign_in() must be called before any Graph request")
        try:
            return request_with_retry(
                self._session,
                method,
                url,
                timeout=HTTP_TIMEOUT_SECONDS,
                max_retries=HTTP_MAX_RETRIES,
                logger=self._log,
                **kwargs,
            )
        except AuthError as exc:
            if exc.status == 401:
                raise GraphAuthError(
                    "Graph returned 401 Unauthorized. The token expired or was revoked. "
                    f"Delete {self._config.graph_token_cache} and run again to sign in."
                ) from exc
            raise GraphAuthError(
                "Graph returned 403 Forbidden. A delegated permission is missing. "
                "Required: Mail.ReadWrite (plus MailboxSettings.Read for category "
                "validation)."
            ) from exc

    def get_json(self, url: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        return self._call("GET", url, params=params).json()

    def iter_paged(
        self, url: str, params: dict[str, Any] | None = None
    ) -> Iterator[dict[str, Any]]:
        """Follow @odata.nextLink until the collection is exhausted."""
        next_url: str | None = url
        next_params = params
        while next_url:
            payload = self.get_json(next_url, next_params)
            for item in payload.get("value", []):
                yield item
            next_url = payload.get("@odata.nextLink")
            next_params = None  # nextLink already carries every parameter

    # -- reads -------------------------------------------------------------

    def list_master_categories(self) -> list[str] | None:
        """Return the mailbox category names, or None when not permitted."""
        try:
            payload = self.get_json(f"{GRAPH_BASE}/me/outlook/masterCategories")
        except (GraphAuthError, HttpError) as exc:
            self._log.warning(
                "cannot read masterCategories, skipping category validation: %s", exc
            )
            return None
        return [str(item.get("displayName", "")) for item in payload.get("value", [])]

    def inbox_folders(self) -> list[tuple[str, str]]:
        """Return [(folder_id, display_path)] for the Inbox and its subfolders."""
        inbox = self.get_json(
            f"{GRAPH_BASE}/me/mailFolders/inbox",
            {"$select": "id,displayName,childFolderCount"},
        )
        folders: list[tuple[str, str]] = []
        self._walk(inbox, "", folders)
        return folders

    def _walk(
        self, folder: dict[str, Any], prefix: str, out: list[tuple[str, str]]
    ) -> None:
        folder_id = str(folder.get("id", ""))
        if not folder_id:
            return
        name = str(folder.get("displayName", "?"))
        path = f"{prefix}/{name}" if prefix else name
        out.append((folder_id, path))
        if int(folder.get("childFolderCount") or 0) <= 0:
            return
        url = f"{GRAPH_BASE}/me/mailFolders/{quote(folder_id, safe='')}/childFolders"
        children = self.iter_paged(
            url, {"$select": "id,displayName,childFolderCount", "$top": 100}
        )
        for child in list(children):
            self._walk(child, path, out)

    def list_messages(
        self,
        folder_id: str,
        since_iso: str | None = None,
        categories: Sequence[str] = (),
    ) -> list[dict[str, Any]]:
        """List messages in one folder, optionally narrowed server-side.

        Filtering on the categories collection is not promised by the Graph
        documentation, but it works - including several categories joined with
        ``or`` - and is markedly faster than listing a whole folder. It is used
        when available and falls back to listing everything (the caller filters
        client-side either way).
        """
        if categories:
            try:
                return self._list(folder_id, since_iso, categories)
            except HttpError as exc:
                self._log.warning(
                    "server-side category filter rejected (HTTP %s), "
                    "listing the folder instead",
                    exc.status,
                )
        return self._list(folder_id, since_iso, ())

    def _list(
        self, folder_id: str, since_iso: str | None, categories: Sequence[str]
    ) -> list[dict[str, Any]]:
        conditions: list[str] = []
        if categories:
            clause = " or ".join(
                f"categories/any(c:c eq '{odata_literal(name)}')" for name in categories
            )
            conditions.append(f"({clause})" if len(categories) > 1 else clause)
        if since_iso:
            conditions.append(f"receivedDateTime ge {since_iso}")

        params: dict[str, Any] = {
            "$select": MESSAGE_LIST_SELECT,
            "$top": GRAPH_PAGE_SIZE,
        }
        if conditions:
            params["$filter"] = " and ".join(conditions)

        url = f"{GRAPH_BASE}/me/mailFolders/{quote(folder_id, safe='')}/messages"
        return list(self.iter_paged(url, params))

    def search_messages(self, term: str, limit: int = 20) -> list[dict[str, Any]]:
        """Search the whole mailbox for a term, used by --explain.

        $search cannot be combined with $filter or $orderby, so this is a plain
        relevance-ordered search across folders.
        """
        rows = self.iter_paged(
            f"{GRAPH_BASE}/me/messages",
            {
                "$search": '"' + term.replace('"', " ") + '"',  # a quote would end the phrase
                "$select": MESSAGE_LIST_SELECT,
                "$top": min(limit, 100),
            },
        )
        found: list[dict[str, Any]] = []
        for row in rows:
            found.append(row)
            if len(found) >= limit:
                break
        return found

    def get_message_detail(self, message_id: str) -> dict[str, Any]:
        """Fetch body and headers for one message.

        internetMessageHeaders is only returned by a single-message GET with
        $select, which is why this is a separate call, made for candidates only.
        """
        url = f"{GRAPH_BASE}/me/messages/{quote(message_id, safe='')}"
        return self.get_json(url, {"$select": MESSAGE_DETAIL_SELECT})

    # -- the only write ----------------------------------------------------

    def patch_categories(self, message_id: str, categories: Sequence[str]) -> None:
        """Write the categories array. The only write this program performs."""
        url = f"{GRAPH_BASE}/me/messages/{quote(message_id, safe='')}"
        self._call(
            "PATCH",
            url,
            json={"categories": list(categories)},
            headers={"Content-Type": "application/json"},
        )
