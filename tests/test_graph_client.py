"""Tests for :mod:`m365_service_comms_mcp.graph_client`."""

from __future__ import annotations

import time

import httpx
import pytest
from azure.core.credentials import AccessToken
from pytest_httpx import HTTPXMock

from m365_service_comms_mcp.auth import GraphAuthProvider
from m365_service_comms_mcp.config import GRAPH_BASE_URL, AuthConfig
from m365_service_comms_mcp.errors import GraphError
from m365_service_comms_mcp.graph_client import GraphClient


class _StubProvider(GraphAuthProvider):
    def __init__(self, token: str = "stub-token") -> None:
        super().__init__(config=AuthConfig(tenant_id="t", client_id="c"))
        self._token = token

    def get_token(self) -> AccessToken:  # type: ignore[override]
        return AccessToken(self._token, int(time.time()) + 3600)


@pytest.fixture
async def client() -> GraphClient:
    return GraphClient(auth=_StubProvider(), max_attempts=2)


async def test_list_health_overviews_returns_parsed_json(
    client: GraphClient,
    httpx_mock: HTTPXMock,
) -> None:
    httpx_mock.add_response(
        url=f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/healthOverviews?%24top=25",
        method="GET",
        json={"value": [{"id": "Exchange", "status": "serviceOperational"}]},
        status_code=200,
    )

    result = await client.list_health_overviews()

    assert result == {"value": [{"id": "Exchange", "status": "serviceOperational"}]}


async def test_list_health_overviews_passes_top(
    client: GraphClient,
    httpx_mock: HTTPXMock,
) -> None:
    httpx_mock.add_response(
        url=f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/healthOverviews?%24top=5",
        method="GET",
        json={"value": []},
    )
    await client.list_health_overviews(top=5)


async def test_list_messages_includes_filter_and_orderby(
    client: GraphClient,
    httpx_mock: HTTPXMock,
) -> None:
    httpx_mock.add_response(
        url=(
            f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/messages"
            "?%24top=10&%24filter=services%2Fany%28s%3As+eq+%27Exchange%27%29"
            "&%24orderby=lastModifiedDateTime+desc"
        ),
        method="GET",
        json={"value": []},
    )

    await client.list_messages(top=10, filter_="services/any(s:s eq 'Exchange')")


async def test_get_message_validates_id(client: GraphClient) -> None:
    with pytest.raises(ValueError):
        await client.get_message("")


async def test_get_message_calls_correct_url(
    client: GraphClient,
    httpx_mock: HTTPXMock,
) -> None:
    httpx_mock.add_response(
        url=f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/messages/MC123",
        method="GET",
        json={"id": "MC123", "title": "Test"},
    )

    result = await client.get_message("MC123")

    assert result["id"] == "MC123"


async def test_403_raises_graph_error_with_envelope(
    client: GraphClient,
    httpx_mock: HTTPXMock,
) -> None:
    httpx_mock.add_response(
        url=f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/healthOverviews?%24top=25",
        method="GET",
        json={
            "error": {
                "code": "Authorization_RequestDenied",
                "message": "Insufficient privileges.",
            }
        },
        status_code=403,
        headers={"request-id": "abc-123"},
    )

    with pytest.raises(GraphError) as info:
        await client.list_health_overviews()

    err = info.value
    assert err.status_code == 403
    assert err.code == "Authorization_RequestDenied"
    assert "Insufficient privileges" in err.message
    assert err.request_id == "abc-123"


async def test_429_is_retried_then_succeeds(
    httpx_mock: HTTPXMock,
) -> None:
    client = GraphClient(auth=_StubProvider(), max_attempts=3)
    httpx_mock.add_response(
        url=f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/healthOverviews?%24top=25",
        method="GET",
        status_code=429,
        json={"error": {"code": "TooManyRequests", "message": "Slow down"}},
    )
    httpx_mock.add_response(
        url=f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/healthOverviews?%24top=25",
        method="GET",
        status_code=200,
        json={"value": []},
    )

    result = await client.list_health_overviews()

    assert result == {"value": []}


async def test_429_after_max_attempts_raises_graph_error(
    httpx_mock: HTTPXMock,
) -> None:
    client = GraphClient(auth=_StubProvider(), max_attempts=2)
    for _ in range(2):
        httpx_mock.add_response(
            url=f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/healthOverviews?%24top=25",
            method="GET",
            status_code=429,
            json={"error": {"code": "TooManyRequests", "message": "Slow down"}},
        )

    with pytest.raises(GraphError) as info:
        await client.list_health_overviews()

    assert info.value.status_code == 429
    assert info.value.code == "TooManyRequests"


async def test_transport_error_is_retried(
    httpx_mock: HTTPXMock,
) -> None:
    client = GraphClient(auth=_StubProvider(), max_attempts=3)
    httpx_mock.add_exception(httpx.ConnectError("DNS lookup failed"))
    httpx_mock.add_response(
        url=f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/healthOverviews?%24top=25",
        method="GET",
        json={"value": []},
    )

    result = await client.list_health_overviews()

    assert result == {"value": []}


async def test_unexpected_status_raises_with_text_fallback(
    client: GraphClient,
    httpx_mock: HTTPXMock,
) -> None:
    httpx_mock.add_response(
        url=f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/healthOverviews?%24top=25",
        method="GET",
        status_code=418,
        text="I'm a teapot",
    )

    with pytest.raises(GraphError) as info:
        await client.list_health_overviews()

    assert info.value.status_code == 418
    assert info.value.code is None
    assert "teapot" in info.value.message
