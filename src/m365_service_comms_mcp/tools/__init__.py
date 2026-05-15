"""Tool registration package.

The :func:`register_tools` helper attaches every v0.1 tool to a
:class:`mcp.server.fastmcp.FastMCP` instance, binding each one to the supplied
:class:`~m365_service_comms_mcp.graph_protocol.GraphClientProtocol` backend
(either the live :class:`~m365_service_comms_mcp.graph_client.GraphClient` or the
:class:`~m365_service_comms_mcp.demo_client.DemoGraphClient`).
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from ..graph_protocol import GraphClientProtocol
from . import health, messages

__all__ = ["register_tools"]


def register_tools(mcp: FastMCP, client: GraphClientProtocol) -> None:
    """Attach every v0.1 tool to ``mcp``, bound to ``client``."""

    health.register(mcp, client)
    messages.register(mcp, client)
