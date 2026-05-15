"""Common protocol implemented by every Graph backend.

Both the live :class:`~m365_service_comms_mcp.graph_client.GraphClient` and the
canned-data :class:`~m365_service_comms_mcp.demo_client.DemoGraphClient` satisfy
this protocol so the MCP tools can be unit-tested and driven in ``--demo`` mode
without changing tool code.
"""

from __future__ import annotations

from typing import Any, Protocol, runtime_checkable


@runtime_checkable
class GraphClientProtocol(Protocol):
    """Subset of Microsoft Graph operations used by the v0.1 MCP tools."""

    async def list_health_overviews(self, *, top: int = 25) -> dict[str, Any]: ...

    async def list_messages(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]: ...

    async def get_message(self, message_id: str) -> dict[str, Any]: ...

    async def aclose(self) -> None: ...
