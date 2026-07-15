"""Console entry point for the m365-svc-comms-mcp command."""

from __future__ import annotations

import argparse
import asyncio
import os
import sys

from . import __version__

# Truthy values accepted for the M365_ENABLE_WRITE env var. Matches the parsing
# rules used by AuthConfig.from_env for M365_AUTH_DEVICE_CODE.
_TRUTHY = {"1", "true", "yes", "on"}


def _env_truthy(value: str | None) -> bool:
    return (value or "").strip().lower() in _TRUTHY


def main() -> int:
    """Console entry point.

    Returns the process exit code (0 on success, non-zero on failure).
    """

    parser = argparse.ArgumentParser(
        prog="m365-service-comms-mcp",
        description=(
            "MCP server for the Microsoft Graph Service Communications API. "
            "Exposes M365 service health, service issues, and Message Center "
            "posts to AI agents. Per-user Message Center viewpoint writes "
            "(mark read / archive / favorite) are available behind the "
            "--enable-write opt-in flag."
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
            "tenant. Requires no environment variables. Compatible with --enable-write — "
            "the demo backend tracks viewpoint state in memory."
        ),
    )
    parser.add_argument(
        "--enable-write",
        action="store_true",
        help=(
            "Enable per-user Message Center viewpoint write tools "
            "(set_message_center_posts_read/archived/favorite). "
            "Requests the ServiceMessageViewpoint.Write delegated scope on first sign-in. "
            "Equivalent to setting M365_ENABLE_WRITE=1."
        ),
    )

    args = parser.parse_args()

    if args.auth_test and args.demo:
        print(
            "error: --auth-test and --demo cannot be combined. "
            "--auth-test always exercises the real Microsoft Graph flow; "
            "--demo replaces the Graph backend with canned data and skips auth.",
            file=sys.stderr,
        )
        return 2

    # Resolve the write gate exactly once: CLI flag wins, env var is the
    # fallback. This avoids any drift between the scopes that AuthConfig
    # requests and the tools that register_tools actually exposes.
    write_enabled = args.enable_write or _env_truthy(os.environ.get("M365_ENABLE_WRITE"))

    if args.auth_test:
        from .auth_test import run_auth_test

        return run_auth_test(write_enabled=write_enabled)

    from .server import run_stdio_server

    try:
        asyncio.run(run_stdio_server(demo=args.demo, write_enabled=write_enabled))
    except KeyboardInterrupt:
        return 130
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
