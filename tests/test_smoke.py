"""Smoke tests that verify the package is importable and self-consistent."""

from __future__ import annotations

import m365_service_comms_mcp
from m365_service_comms_mcp.demo_client import DemoGraphClient
from m365_service_comms_mcp.server import SERVER_NAME, build_server


def test_package_exposes_version() -> None:
    assert m365_service_comms_mcp.__version__
    parts = m365_service_comms_mcp.__version__.split(".")
    assert len(parts) >= 2
    assert all(
        part.isdigit() or part.replace("rc", "").replace("a", "").replace("b", "").isdigit()
        for part in parts
    ), f"unexpected version segments: {parts}"


def test_build_server_returns_named_fastmcp_with_tools_registered() -> None:
    mcp = build_server(DemoGraphClient())
    assert mcp.name == SERVER_NAME
    # FastMCP exposes a tool manager; verify the three v0.1 tools are present.
    tool_names = {tool.name for tool in mcp._tool_manager.list_tools()}
    assert tool_names == {
        "list_service_health",
        "list_message_center_posts",
        "get_message_center_post",
    }
