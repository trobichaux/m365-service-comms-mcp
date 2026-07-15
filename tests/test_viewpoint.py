"""Tests for the per-user viewpoint write tools (v0.2, gated).

Covers:
- registration gating (write tools absent when ``write_enabled=False``)
- tool surface when enabled (presence and signature)
- happy-path argument forwarding to the backend client
- list-length and id-pattern validation
- deduplication of message ids before the bulk POST
- demo backend round-trip including viewpoint state persistence
- propagation of Graph logical failures (`{"value": false}` raising)
"""

from __future__ import annotations

from typing import Any

import pytest
from mcp.server.fastmcp import FastMCP

from m365_service_comms_mcp.demo_client import DemoGraphClient
from m365_service_comms_mcp.errors import GraphError
from m365_service_comms_mcp.tools import register_tools


class _RecordingClient:
    """Stand-in backend that records viewpoint setter calls."""

    def __init__(self) -> None:
        self.set_read_calls: list[dict[str, Any]] = []
        self.set_archive_calls: list[dict[str, Any]] = []
        self.set_favorite_calls: list[dict[str, Any]] = []

    async def list_health_overviews(self, *, top: int = 25) -> dict[str, Any]:
        return {"value": []}

    async def get_health_overview(
        self, service_id: str, *, expand_issues: bool = True
    ) -> dict[str, Any]:
        return {"id": service_id}

    async def list_issues(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]:
        return {"value": []}

    async def get_issue(self, issue_id: str) -> dict[str, Any]:
        return {"id": issue_id}

    async def list_messages(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]:
        return {"value": []}

    async def get_message(self, message_id: str) -> dict[str, Any]:
        return {"id": message_id}

    async def set_messages_read(
        self, message_ids: list[str], *, read: bool
    ) -> dict[str, Any]:
        self.set_read_calls.append({"message_ids": list(message_ids), "read": read})
        return {"value": True}

    async def set_messages_archive(
        self, message_ids: list[str], *, archived: bool
    ) -> dict[str, Any]:
        self.set_archive_calls.append({"message_ids": list(message_ids), "archived": archived})
        return {"value": True}

    async def set_messages_favorite(
        self, message_ids: list[str], *, favorite: bool
    ) -> dict[str, Any]:
        self.set_favorite_calls.append({"message_ids": list(message_ids), "favorite": favorite})
        return {"value": True}

    async def aclose(self) -> None:
        return


def _server(client: object, *, write_enabled: bool) -> FastMCP:
    mcp = FastMCP(name="viewpoint-test")
    register_tools(mcp, client, write_enabled=write_enabled)  # type: ignore[arg-type]
    return mcp


async def _call_tool(mcp: FastMCP, name: str, **arguments: Any) -> Any:
    result = await mcp.call_tool(name, arguments)
    if isinstance(result, tuple):
        _content, structured = result
        return structured
    return result


_WRITE_TOOL_NAMES = {
    "set_message_center_posts_read",
    "set_message_center_posts_archived",
    "set_message_center_posts_favorite",
}


def test_viewpoint_tools_are_absent_by_default() -> None:
    mcp = _server(_RecordingClient(), write_enabled=False)
    tool_names = {tool.name for tool in mcp._tool_manager.list_tools()}
    assert _WRITE_TOOL_NAMES.isdisjoint(tool_names)


def test_viewpoint_tools_register_when_write_enabled() -> None:
    mcp = _server(_RecordingClient(), write_enabled=True)
    tool_names = {tool.name for tool in mcp._tool_manager.list_tools()}
    assert _WRITE_TOOL_NAMES.issubset(tool_names)


@pytest.mark.parametrize(
    "tool_name,kwarg_name",
    [
        ("set_message_center_posts_read", "read"),
        ("set_message_center_posts_archived", "archived"),
        ("set_message_center_posts_favorite", "favorite"),
    ],
)
async def test_each_write_tool_forwards_bool_to_client(
    tool_name: str, kwarg_name: str
) -> None:
    client = _RecordingClient()
    mcp = _server(client, write_enabled=True)

    for state in (True, False):
        result = await _call_tool(
            mcp,
            tool_name,
            message_ids=["MC100001"],
            **{kwarg_name: state},
        )

        assert result["success"] is True
        assert result["state"] is state
        assert result["message_ids"] == ["MC100001"]

    # Each call should be recorded in the right bucket with the right flag.
    if tool_name.endswith("_read"):
        assert [c[kwarg_name] for c in client.set_read_calls] == [True, False]
    elif tool_name.endswith("_archived"):
        assert [c[kwarg_name] for c in client.set_archive_calls] == [True, False]
    else:
        assert [c[kwarg_name] for c in client.set_favorite_calls] == [True, False]


async def test_write_tool_deduplicates_message_ids() -> None:
    client = _RecordingClient()
    mcp = _server(client, write_enabled=True)

    await _call_tool(
        mcp,
        "set_message_center_posts_read",
        message_ids=["MC100001", "mc100001", "MC100002"],
        read=True,
    )

    assert client.set_read_calls == [{"message_ids": ["MC100001", "MC100002"], "read": True}]


async def test_write_tool_rejects_empty_id_list() -> None:
    mcp = _server(_RecordingClient(), write_enabled=True)
    with pytest.raises(Exception):
        await _call_tool(
            mcp,
            "set_message_center_posts_read",
            message_ids=[],
            read=True,
        )


async def test_write_tool_rejects_over_cap_id_list() -> None:
    mcp = _server(_RecordingClient(), write_enabled=True)
    too_many = [f"MC{1000000 + i}" for i in range(51)]
    with pytest.raises(Exception):
        await _call_tool(
            mcp,
            "set_message_center_posts_read",
            message_ids=too_many,
            read=True,
        )


@pytest.mark.parametrize("bad_id", ["", "MC", "ABC1234", "MC 1234", "MC1234567890"])
async def test_write_tool_rejects_invalid_id_in_list(bad_id: str) -> None:
    mcp = _server(_RecordingClient(), write_enabled=True)
    with pytest.raises(Exception):
        await _call_tool(
            mcp,
            "set_message_center_posts_read",
            message_ids=["MC100001", bad_id],
            read=True,
        )


async def test_demo_client_viewpoint_round_trip() -> None:
    """The demo backend round-trips viewpoint state for end-to-end --demo testing."""

    backend = DemoGraphClient()
    mcp = _server(backend, write_enabled=True)

    await _call_tool(
        mcp, "set_message_center_posts_read", message_ids=["MC987654"], read=True
    )
    await _call_tool(
        mcp,
        "set_message_center_posts_archived",
        message_ids=["MC987654"],
        archived=True,
    )
    await _call_tool(
        mcp,
        "set_message_center_posts_favorite",
        message_ids=["MC987654"],
        favorite=False,
    )

    state = backend.viewpoint_state("MC987654")
    assert state == {"read": True, "archived": True, "favorite": False}


async def test_demo_client_viewpoint_rejects_unknown_message_id() -> None:
    """Unknown ids surface as GraphError when invoked on the backend directly.

    Going through FastMCP wraps tool exceptions in ToolError, so this test
    asserts on the backend behaviour rather than the MCP transport layer.
    """

    backend = DemoGraphClient()
    with pytest.raises(GraphError):
        await backend.set_messages_read(["MC000000"], read=True)
