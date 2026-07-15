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
    """Subset of Microsoft Graph operations used by the v0.2 MCP tools."""

    async def list_health_overviews(self, *, top: int = 25) -> dict[str, Any]:
        """List service health overviews. See ``GraphClient.list_health_overviews``."""

    async def get_health_overview(
        self,
        service_id: str,
        *,
        expand_issues: bool = True,
    ) -> dict[str, Any]:
        """Get one service's health overview. See ``GraphClient.get_health_overview``."""

    async def list_issues(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]:
        """List service health issues. See ``GraphClient.list_issues``."""

    async def get_issue(self, issue_id: str) -> dict[str, Any]:
        """Fetch a single service health issue. See ``GraphClient.get_issue``."""

    async def list_messages(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]:
        """List Message Center posts. See ``GraphClient.list_messages``."""

    async def get_message(self, message_id: str) -> dict[str, Any]:
        """Fetch a single Message Center post by ID. See ``GraphClient.get_message``."""

    async def set_messages_read(
        self,
        message_ids: list[str],
        *,
        read: bool,
    ) -> dict[str, Any]:
        """Mark messages read or unread. See ``GraphClient.set_messages_read``."""

    async def set_messages_archive(
        self,
        message_ids: list[str],
        *,
        archived: bool,
    ) -> dict[str, Any]:
        """Archive or unarchive messages. See ``GraphClient.set_messages_archive``."""

    async def set_messages_favorite(
        self,
        message_ids: list[str],
        *,
        favorite: bool,
    ) -> dict[str, Any]:
        """Favorite or unfavorite messages. See ``GraphClient.set_messages_favorite``."""

    async def aclose(self) -> None:
        """Release any underlying resources (e.g. HTTP client)."""
