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


def test_build_server_default_exposes_read_only_tool_set() -> None:
    mcp = build_server(DemoGraphClient())
    assert mcp.name == SERVER_NAME
    tool_names = {tool.name for tool in mcp._tool_manager.list_tools()}
    assert tool_names == {
        "list_service_health",
        "get_service_health",
        "list_service_issues",
        "get_service_issue",
        "list_message_center_posts",
        "get_message_center_post",
    }


def test_build_server_write_enabled_exposes_viewpoint_tools() -> None:
    mcp = build_server(DemoGraphClient(), write_enabled=True)
    tool_names = {tool.name for tool in mcp._tool_manager.list_tools()}
    assert tool_names == {
        "list_service_health",
        "get_service_health",
        "list_service_issues",
        "get_service_issue",
        "list_message_center_posts",
        "get_message_center_post",
        "set_message_center_posts_read",
        "set_message_center_posts_archived",
        "set_message_center_posts_favorite",
    }
