"""Tool registration package.

The :func:`register_tools` helper attaches every tool to a
:class:`mcp.server.fastmcp.FastMCP` instance, binding each one to the supplied
:class:`~m365_service_comms_mcp.graph_protocol.GraphClientProtocol` backend
(either the live :class:`~m365_service_comms_mcp.graph_client.GraphClient` or the
:class:`~m365_service_comms_mcp.demo_client.DemoGraphClient`).

The 3 viewpoint write tools are registered **only** when ``write_enabled`` is
True, so the default install is read-only with no extra OAuth scope requested.
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from ..graph_protocol import GraphClientProtocol
from . import health, issues, messages, viewpoint
from ._constants import MAX_TOP
from .viewpoint import MAX_MESSAGE_IDS_PER_WRITE

__all__ = ["MAX_MESSAGE_IDS_PER_WRITE", "MAX_TOP", "register_tools"]


def register_tools(
    mcp: FastMCP,
    client: GraphClientProtocol,
    *,
    write_enabled: bool = False,
) -> None:
    """Attach the v0.2 tool set to ``mcp``, bound to ``client``.

    When ``write_enabled`` is True the per-user viewpoint write tools are also
    registered. When False (the default) only the 6 read tools are exposed and
    no write-scope is implied for the OAuth flow.
    """

    health.register(mcp, client)
    issues.register(mcp, client)
    messages.register(mcp, client)
    if write_enabled:
        viewpoint.register(mcp, client)
