"""MCP server wiring.

Constructs a :class:`~mcp.server.fastmcp.FastMCP` instance, attaches the v0.2
tools (bound to either the live Graph client or the demo client), and exposes
``run_stdio_server`` for the console entry point to call.
"""

from __future__ import annotations

import logging
import sys

from mcp.server.fastmcp import FastMCP

from . import __version__
from .auth import GraphAuthProvider
from .config import AuthConfig
from .demo_client import DemoGraphClient
from .graph_client import GraphClient
from .graph_protocol import GraphClientProtocol
from .tools import register_tools

logger = logging.getLogger(__name__)

SERVER_NAME = "m365-service-comms"


def build_server(client: GraphClientProtocol, *, write_enabled: bool = False) -> FastMCP:
    """Build a :class:`FastMCP` instance with the v0.2 tool set registered.

    ``client`` is the Graph backend the tools will call. Tests inject a fake
    backend; production callers inject either :class:`GraphClient` (live Graph)
    or :class:`DemoGraphClient` (canned data via ``--demo``).

    When ``write_enabled`` is True, the per-user viewpoint write tools are
    registered alongside the reads. Default is False (read-only), matching the
    v0.1 behaviour for callers who haven't opted in.
    """

    mcp = FastMCP(name=SERVER_NAME)
    register_tools(mcp, client, write_enabled=write_enabled)
    return mcp


def _configure_stderr_logging() -> None:
    """Force every log handler to write to ``stderr``.

    For stdio MCP servers the parent process owns ``stdout`` for JSON-RPC
    framing; any byte that leaks onto ``stdout`` corrupts the transport. We
    install our own handler at ``WARNING`` (matching Python's default level)
    pointed at ``stderr`` so anything that ``logging``-based libraries
    (azure-identity, msal, httpx) might emit ends up in the right stream.
    """

    root = logging.getLogger()
    if not root.handlers:
        handler = logging.StreamHandler(stream=sys.stderr)
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        root.addHandler(handler)
        root.setLevel(logging.WARNING)


async def run_stdio_server(*, demo: bool = False, write_enabled: bool = False) -> None:
    """Start the MCP server on stdio and run until the client disconnects.

    When ``demo=True``, no auth is performed and Graph is never called \u2014 the
    server returns canned sample data so the wire protocol can be exercised on
    a tenant where admin consent is not available. The demo backend supports the
    viewpoint write tools against in-memory state, so ``--demo --enable-write``
    is a valid combination for end-to-end testing.
    """

    _configure_stderr_logging()

    client: GraphClientProtocol
    if demo:
        print(
            f"[{SERVER_NAME}/{__version__}] Starting in --demo mode \u2014 no Graph calls will be made.",
            file=sys.stderr,
            flush=True,
        )
        client = DemoGraphClient()
    else:
        config = AuthConfig.from_env(write_enabled_override=write_enabled)
        auth = GraphAuthProvider(config=config)
        client = GraphClient(auth=auth)

    try:
        mcp = build_server(client, write_enabled=write_enabled)
        await mcp.run_stdio_async()
    finally:
        await client.aclose()


__all__ = ["SERVER_NAME", "__version__", "build_server", "run_stdio_server"]
