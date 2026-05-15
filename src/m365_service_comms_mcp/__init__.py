"""m365-service-comms-mcp — MCP server for Microsoft Graph Service Communications API.

A read-only Model Context Protocol server that exposes M365 service health and Message
Center posts to AI agents (Claude, Copilot, Cursor, etc.) via the Microsoft Graph
``serviceAnnouncement`` resource.
"""

from __future__ import annotations

__version__ = "0.1.3"
__all__ = ["__version__"]
