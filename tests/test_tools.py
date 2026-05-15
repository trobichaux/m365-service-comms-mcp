"""Tests for the v0.1 MCP tools.

Tools are exercised through their FastMCP registration so the test verifies the
end-to-end path the MCP client actually takes (input validation via Pydantic →
tool callable → Graph client). A :class:`RecordingClient` stand-in is used
instead of either the live or demo client so we can assert exact arguments.
"""

from __future__ import annotations

import json
from typing import Any

import pytest
from mcp.server.fastmcp import FastMCP
from pydantic import ValidationError

from m365_service_comms_mcp.demo_client import DemoGraphClient
from m365_service_comms_mcp.tools import register_tools


class _RecordingClient:
    """Captures every call made by the tools so tests can assert behaviour."""

    def __init__(self) -> None:
        self.health_calls: list[dict[str, Any]] = []
        self.messages_calls: list[dict[str, Any]] = []
        self.get_message_calls: list[str] = []

    async def list_health_overviews(self, *, top: int = 25) -> dict[str, Any]:
        self.health_calls.append({"top": top})
        return {"value": [{"id": "Exchange", "status": "serviceOperational"}]}

    async def list_messages(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]:
        self.messages_calls.append({"top": top, "filter_": filter_, "orderby": orderby})
        return {"value": []}

    async def get_message(self, message_id: str) -> dict[str, Any]:
        self.get_message_calls.append(message_id)
        return {"id": message_id, "title": "stub"}

    async def aclose(self) -> None:
        return


def _server_with(client: object) -> FastMCP:
    mcp = FastMCP(name="test-server")
    register_tools(mcp, client)  # type: ignore[arg-type]
    return mcp


async def _call_tool(mcp: FastMCP, name: str, **arguments: Any) -> Any:
    result = await mcp.call_tool(name, arguments)
    # FastMCP returns (content_list, structured_dict) for structured tools.
    if isinstance(result, tuple):
        _content, structured = result
        return structured
    return result


async def test_list_service_health_default_top() -> None:
    client = _RecordingClient()
    mcp = _server_with(client)

    structured = await _call_tool(mcp, "list_service_health")

    assert client.health_calls == [{"top": 25}]
    assert structured == {"value": [{"id": "Exchange", "status": "serviceOperational"}]}


async def test_list_service_health_passes_top() -> None:
    client = _RecordingClient()
    mcp = _server_with(client)

    await _call_tool(mcp, "list_service_health", top=5)

    assert client.health_calls == [{"top": 5}]


@pytest.mark.parametrize("bad_top", [0, -1, 51, 100])
async def test_list_service_health_rejects_out_of_range_top(bad_top: int) -> None:
    client = _RecordingClient()
    mcp = _server_with(client)

    with pytest.raises((ValidationError, Exception)) as info:
        await _call_tool(mcp, "list_service_health", top=bad_top)

    assert client.health_calls == []
    # Either Pydantic ValidationError or FastMCP wrapping; just confirm it complained.
    assert "top" in str(info.value).lower() or "validation" in str(info.value).lower()


async def test_list_messages_no_filters_sends_no_filter() -> None:
    client = _RecordingClient()
    mcp = _server_with(client)

    await _call_tool(mcp, "list_message_center_posts")

    assert client.messages_calls == [
        {"top": 25, "filter_": None, "orderby": "lastModifiedDateTime desc"}
    ]


async def test_list_messages_builds_combined_filter() -> None:
    client = _RecordingClient()
    mcp = _server_with(client)

    await _call_tool(
        mcp,
        "list_message_center_posts",
        top=10,
        category="planForChange",
        severity="high",
    )

    assert client.messages_calls == [
        {
            "top": 10,
            "filter_": "category eq 'planForChange' and severity eq 'high'",
            "orderby": "lastModifiedDateTime desc",
        }
    ]


async def test_list_messages_rejects_unknown_category() -> None:
    client = _RecordingClient()
    mcp = _server_with(client)

    with pytest.raises(Exception):
        await _call_tool(mcp, "list_message_center_posts", category="totally-fake")

    assert client.messages_calls == []


@pytest.mark.parametrize("good_id", ["MC1234", "MC123456", "mc987654"])
async def test_get_message_accepts_valid_ids(good_id: str) -> None:
    client = _RecordingClient()
    mcp = _server_with(client)

    await _call_tool(mcp, "get_message_center_post", message_id=good_id)

    assert client.get_message_calls == [good_id]


@pytest.mark.parametrize("bad_id", ["", "not-an-id", "MC", "12345", "MC1234567890"])
async def test_get_message_rejects_invalid_ids(bad_id: str) -> None:
    client = _RecordingClient()
    mcp = _server_with(client)

    with pytest.raises(Exception):
        await _call_tool(mcp, "get_message_center_post", message_id=bad_id)

    assert client.get_message_calls == []


async def test_demo_client_round_trip() -> None:
    """End-to-end: --demo backend through the real tool registration."""

    mcp = _server_with(DemoGraphClient())

    health = await _call_tool(mcp, "list_service_health", top=2)
    assert len(health["value"]) == 2

    messages = await _call_tool(mcp, "list_message_center_posts", top=1)
    assert len(messages["value"]) == 1

    detail = await _call_tool(mcp, "get_message_center_post", message_id="MC987654")
    assert detail["id"] == "MC987654"
    assert "body" in detail
    # Sanity: structured output is JSON-serializable.
    json.dumps(detail)
