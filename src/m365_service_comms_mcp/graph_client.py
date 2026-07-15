"""Microsoft Graph HTTP client for the Service Communications API.

A thin async wrapper around :class:`httpx.AsyncClient` covering the endpoints
used by the v0.2 tools:

- ``GET /admin/serviceAnnouncement/healthOverviews``
- ``GET /admin/serviceAnnouncement/healthOverviews/{id}``
- ``GET /admin/serviceAnnouncement/issues``
- ``GET /admin/serviceAnnouncement/issues/{id}``
- ``GET /admin/serviceAnnouncement/messages``
- ``GET /admin/serviceAnnouncement/messages/{id}``
- ``POST /admin/serviceAnnouncement/messages/markRead`` / ``markUnread``
- ``POST /admin/serviceAnnouncement/messages/archive`` / ``unarchive``
- ``POST /admin/serviceAnnouncement/messages/favorite`` / ``unfavorite``

Cross-cutting concerns:

- **Auth**: each request adds a fresh bearer token via :class:`GraphAuthProvider`.
- **Retries**: HTTP 429 and retryable 5xx are retried up to ``max_attempts``
  times. The wait between attempts uses the response's ``Retry-After`` header
  when present, otherwise an exponential backoff. The viewpoint POSTs are
  idempotent (mark-read-twice = same as once), so the same retry policy
  applies to GETs and POSTs alike.
- **Errors**: non-success responses raise :class:`~m365_service_comms_mcp.errors.GraphError`
  with the Graph error envelope parsed out for the LLM. Viewpoint POSTs
  that return HTTP 200 with ``{"value": false}`` are also surfaced as
  :class:`GraphError` so a logical failure isn't silently reported as success.
"""

from __future__ import annotations

from typing import Any

import httpx
from tenacity import (
    AsyncRetrying,
    RetryCallState,
    RetryError,
    retry_if_exception_type,
    stop_after_attempt,
    wait_random_exponential,
)

from . import __version__
from .auth import GraphAuthProvider
from .config import GRAPH_BASE_URL
from .errors import GraphError, parse_graph_error

_USER_AGENT = f"m365-service-comms-mcp/{__version__}"
_DEFAULT_TIMEOUT = httpx.Timeout(30.0, connect=10.0)
_RETRY_STATUS = frozenset({429, 500, 502, 503, 504})
_MAX_RETRY_AFTER_SECONDS = 120.0


class _Retryable(Exception):
    """Sentinel raised inside the retry loop for transient failures.

    Carries the parsed :class:`GraphError` and the suggested wait time (parsed
    from the ``Retry-After`` response header, when present).
    """

    def __init__(self, graph_error: GraphError, retry_after: float | None) -> None:
        super().__init__(str(graph_error))
        self.graph_error = graph_error
        self.retry_after = retry_after


class GraphClient:
    """Minimal Graph API client scoped to the Service Communications endpoints."""

    def __init__(
        self,
        *,
        auth: GraphAuthProvider,
        base_url: str = GRAPH_BASE_URL,
        http_client: httpx.AsyncClient | None = None,
        max_attempts: int = 4,
    ) -> None:
        self._auth = auth
        self._base_url = base_url.rstrip("/")
        self._max_attempts = max(1, max_attempts)
        self._owns_client = http_client is None
        self._client = http_client or httpx.AsyncClient(
            timeout=_DEFAULT_TIMEOUT,
            headers={"User-Agent": _USER_AGENT, "Accept": "application/json"},
        )

    async def aclose(self) -> None:
        if self._owns_client:
            await self._client.aclose()

    async def __aenter__(self) -> GraphClient:
        return self

    async def __aexit__(self, *exc_info: object) -> None:
        await self.aclose()

    async def list_health_overviews(self, *, top: int = 25) -> dict[str, Any]:
        """List service health overviews for the tenant.

        Maps to ``GET /admin/serviceAnnouncement/healthOverviews``.
        """

        return await self._get(
            "/admin/serviceAnnouncement/healthOverviews",
            params={"$top": str(top)},
        )

    async def get_health_overview(
        self,
        service_id: str,
        *,
        expand_issues: bool = True,
    ) -> dict[str, Any]:
        """Get the health overview for a single service, with related issues.

        Maps to ``GET /admin/serviceAnnouncement/healthOverviews/{id}``.

        ``service_id`` is the ``id`` field from a healthOverview record
        (e.g. ``Exchange``, ``microsoftteams``), URL-encoded into the path so
        unusual characters in future Graph IDs don't break the request.
        """

        if not service_id or not service_id.strip():
            raise ValueError("service_id must be a non-empty string")

        from urllib.parse import quote

        encoded = quote(service_id.strip(), safe="")
        params: dict[str, str] = {}
        if expand_issues:
            params["$expand"] = "issues"
        return await self._get(
            f"/admin/serviceAnnouncement/healthOverviews/{encoded}",
            params=params or None,
        )

    async def list_issues(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]:
        """List service health issues for the tenant.

        Maps to ``GET /admin/serviceAnnouncement/issues``. ``filter_`` is passed
        verbatim to the Graph ``$filter`` query option; callers must build valid
        OData filter expressions (including escaping single quotes in string
        literals).
        """

        params: dict[str, str] = {"$top": str(top)}
        if filter_:
            params["$filter"] = filter_
        if orderby:
            params["$orderby"] = orderby
        return await self._get("/admin/serviceAnnouncement/issues", params=params)

    async def get_issue(self, issue_id: str) -> dict[str, Any]:
        """Fetch a single service health issue by ID.

        Maps to ``GET /admin/serviceAnnouncement/issues/{id}``.
        """

        if not issue_id or not issue_id.strip():
            raise ValueError("issue_id must be a non-empty string")

        return await self._get(f"/admin/serviceAnnouncement/issues/{issue_id.strip()}")

    async def list_messages(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]:
        """List Message Center posts for the tenant.

        Maps to ``GET /admin/serviceAnnouncement/messages``. ``filter_`` is
        passed verbatim to the Graph ``$filter`` query option; tools must build
        valid OData filter expressions.
        """

        params: dict[str, str] = {"$top": str(top)}
        if filter_:
            params["$filter"] = filter_
        if orderby:
            params["$orderby"] = orderby
        return await self._get("/admin/serviceAnnouncement/messages", params=params)

    async def get_message(self, message_id: str) -> dict[str, Any]:
        """Fetch a single Message Center post by ID.

        Maps to ``GET /admin/serviceAnnouncement/messages/{id}``.
        """

        if not message_id or not message_id.strip():
            raise ValueError("message_id must be a non-empty string")

        return await self._get(f"/admin/serviceAnnouncement/messages/{message_id.strip()}")

    async def set_messages_read(
        self,
        message_ids: list[str],
        *,
        read: bool,
    ) -> dict[str, Any]:
        """Mark messages as read (``read=True``) or unread (``read=False``).

        Maps to ``POST /admin/serviceAnnouncement/messages/markRead`` and the
        ``markUnread`` companion. Returns the raw Graph response so callers can
        inspect any future fields; the underlying ``_post`` already raised
        :class:`GraphError` when Graph reported a logical failure.
        """

        action = "markRead" if read else "markUnread"
        return await self._post(
            f"/admin/serviceAnnouncement/messages/{action}",
            json={"messageIds": list(message_ids)},
        )

    async def set_messages_archive(
        self,
        message_ids: list[str],
        *,
        archived: bool,
    ) -> dict[str, Any]:
        """Archive (``archived=True``) or unarchive a list of messages."""

        action = "archive" if archived else "unarchive"
        return await self._post(
            f"/admin/serviceAnnouncement/messages/{action}",
            json={"messageIds": list(message_ids)},
        )

    async def set_messages_favorite(
        self,
        message_ids: list[str],
        *,
        favorite: bool,
    ) -> dict[str, Any]:
        """Favorite (``favorite=True``) or unfavorite a list of messages."""

        action = "favorite" if favorite else "unfavorite"
        return await self._post(
            f"/admin/serviceAnnouncement/messages/{action}",
            json={"messageIds": list(message_ids)},
        )

    async def _get(self, path: str, *, params: dict[str, str] | None = None) -> dict[str, Any]:
        return await self._request("GET", path, params=params, json=None)

    async def _post(self, path: str, *, json: dict[str, Any]) -> dict[str, Any]:
        return await self._request("POST", path, params=None, json=json)

    async def _request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, str] | None,
        json: dict[str, Any] | None,
    ) -> dict[str, Any]:
        url = f"{self._base_url}{path}"

        async def _attempt() -> dict[str, Any]:
            token = self._auth.get_token().token
            response = await self._client.request(
                method,
                url,
                params=params,
                json=json,
                headers={"Authorization": f"Bearer {token}"},
            )
            return _interpret_response(response)

        try:
            async for attempt in AsyncRetrying(
                stop=stop_after_attempt(self._max_attempts),
                wait=_wait_strategy,
                retry=retry_if_exception_type((_Retryable, httpx.TransportError)),
                reraise=True,
            ):
                with attempt:
                    return await _attempt()
        except _Retryable as retryable:
            raise retryable.graph_error from None
        except RetryError as exc:
            cause = exc.last_attempt.exception()
            if isinstance(cause, _Retryable):
                raise cause.graph_error from None
            if cause is not None:
                raise cause from None
            raise

        # Unreachable, but keeps type checkers happy.
        return await _attempt()


def _wait_strategy(retry_state: RetryCallState) -> float:
    """Honour ``Retry-After`` from the response when present.

    Falls back to bounded exponential backoff (``tenacity``'s
    ``wait_random_exponential(multiplier=1, max=30)``) when the failure was a
    transport error or the response did not include ``Retry-After``.
    """

    exc = retry_state.outcome.exception() if retry_state.outcome else None
    if isinstance(exc, _Retryable) and exc.retry_after is not None:
        return min(exc.retry_after, _MAX_RETRY_AFTER_SECONDS)

    return wait_random_exponential(multiplier=1, max=30)(retry_state)


def _parse_retry_after(response: httpx.Response) -> float | None:
    """Parse the ``Retry-After`` header (RFC 9110): either a number of seconds
    or an HTTP-date. Returns ``None`` when absent or unparseable.
    """

    raw = response.headers.get("retry-after")
    if not raw:
        return None
    raw = raw.strip()

    try:
        return max(0.0, float(raw))
    except ValueError:
        pass

    try:
        from email.utils import parsedate_to_datetime

        target = parsedate_to_datetime(raw)
    except (TypeError, ValueError):
        return None

    import datetime as _dt

    now = _dt.datetime.now(tz=target.tzinfo or _dt.UTC)
    return max(0.0, (target - now).total_seconds())


def _interpret_response(response: httpx.Response) -> dict[str, Any]:
    """Translate an :class:`httpx.Response` into a parsed dict or raise."""

    if response.status_code == 200:
        try:
            body = response.json()
        except ValueError as exc:
            raise GraphError(
                status_code=response.status_code,
                code=None,
                message=f"Graph returned non-JSON body: {exc}",
                request_id=response.headers.get("request-id"),
            ) from None

        # Viewpoint POSTs (markRead / archive / favorite / ...) return
        # {"value": true} on success and {"value": false} on logical failure.
        # Surface ``false`` as an error rather than letting it be silently
        # reported as success by callers that only check for HTTP 200.
        if isinstance(body, dict) and body.get("value") is False:
            raise GraphError(
                status_code=response.status_code,
                code="ServiceCommunicationsOperationFailed",
                message=(
                    "Graph returned HTTP 200 but the operation reported failure "
                    "(value=false). Common causes include unknown message IDs or "
                    "missing per-user permissions to mutate the requested state."
                ),
                request_id=response.headers.get("request-id"),
            )

        return body

    body = _safe_json(response)
    request_id = response.headers.get("request-id")

    if response.status_code in _RETRY_STATUS:
        graph_err = parse_graph_error(
            status_code=response.status_code,
            body=body,
            raw_text=response.text,
            request_id=request_id,
        )
        raise _Retryable(graph_err, _parse_retry_after(response))

    raise parse_graph_error(
        status_code=response.status_code,
        body=body,
        raw_text=response.text,
        request_id=request_id,
    )


def _safe_json(response: httpx.Response) -> object:
    try:
        return response.json()
    except ValueError:
        return None
