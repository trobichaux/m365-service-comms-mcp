"""Console entry point for the m365-svc-comms-mcp command."""

from __future__ import annotations

import argparse
import asyncio

from . import __version__


def main() -> int:
    """Console entry point.

    Returns the process exit code (0 on success, non-zero on failure).
    """

    parser = argparse.ArgumentParser(
        prog="m365-svc-comms-mcp",
        description=(
            "Read-only MCP server for the Microsoft Graph Service Communications API. "
            "Exposes M365 service health and Message Center posts to AI agents."
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    parser.add_argument(
        "--auth-test",
        action="store_true",
        help=(
            "Acquire a token interactively, print the resolved tenant and Graph scopes, "
            "then exit without starting the MCP server."
        ),
    )
    parser.add_argument(
        "--demo",
        action="store_true",
        help=(
            "Run the server with canned sample data instead of calling Microsoft Graph. "
            "Useful for verifying your MCP client end-to-end without an admin-grantable "
            "tenant. Requires no environment variables."
        ),
    )

    args = parser.parse_args()

    if args.auth_test:
        from .auth_test import run_auth_test

        return run_auth_test()

    from .server import run_stdio_server

    try:
        asyncio.run(run_stdio_server(demo=args.demo))
    except KeyboardInterrupt:
        return 130
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
