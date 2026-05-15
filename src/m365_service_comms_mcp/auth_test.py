"""Implementation of the ``--auth-test`` console flag.

Acquires a token using the configured credential, calls a low-cost Microsoft
Graph endpoint (``/admin/serviceAnnouncement/healthOverviews?$top=1``), and
prints a human-readable summary so the customer can confirm:

1. Their environment variables are set correctly.
2. The OAuth flow completes (interactive or device-code).
3. The signed-in user holds an admin role with at least
   :code:`ServiceHealth.Read.All` consented.

Exits ``0`` on success, ``1`` on auth or Graph failure, ``2`` on configuration
errors.
"""

from __future__ import annotations

import sys
from typing import TextIO

import httpx

from .auth import GraphAuthProvider
from .config import GRAPH_BASE_URL, AuthConfig, ConfigError

_GRAPH_HEALTH_PROBE = f"{GRAPH_BASE_URL}/admin/serviceAnnouncement/healthOverviews?$top=1"


def run_auth_test(
    *,
    out: TextIO | None = None,
    err: TextIO | None = None,
    auth_provider: GraphAuthProvider | None = None,
    http_client: httpx.Client | None = None,
) -> int:
    """Execute the auth-test flow.

    All collaborators are injectable so the function can be tested without
    standing up real network or browser interactions.
    """

    out = out or sys.stdout
    err = err or sys.stderr

    try:
        config = AuthConfig.from_env()
    except ConfigError as exc:
        print(f"\u274c  Configuration error: {exc}", file=err)
        return 2

    print(f"Tenant ID : {config.tenant_id}", file=out)
    print(f"Client ID : {config.client_id}", file=out)
    print(
        f"Auth flow : {'device-code' if config.prefer_device_code else 'interactive-browser (default)'}",
        file=out,
    )
    print("Acquiring access token \u2026", file=out)

    provider = auth_provider or GraphAuthProvider(config=config)
    try:
        access_token = provider.get_token()
    except Exception as exc:
        print(f"\u274c  Failed to acquire token: {exc}", file=err)
        return 1

    print("\u2713  Token acquired.", file=out)
    print(f"Probing {_GRAPH_HEALTH_PROBE} \u2026", file=out)

    client = http_client or httpx.Client(timeout=30.0)
    close_client = http_client is None
    try:
        response = client.get(
            _GRAPH_HEALTH_PROBE,
            headers={"Authorization": f"Bearer {access_token.token}"},
        )
    except httpx.HTTPError as exc:
        print(f"\u274c  Graph request failed: {exc}", file=err)
        return 1
    finally:
        if close_client:
            client.close()

    if response.status_code == 200:
        body = _safe_json(response)
        count = len(body.get("value", [])) if isinstance(body, dict) else 0
        print(
            f"\u2713  Graph responded HTTP 200 (returned {count} healthOverview record(s)).",
            file=out,
        )
        print(
            "Auth test passed. ServiceHealth.Read.All is granted and admin consent is in place.",
            file=out,
        )
        return 0

    if response.status_code in (401, 403):
        body = _safe_json(response)
        graph_message = _extract_graph_error(body)
        print(
            f"\u274c  Graph rejected the token (HTTP {response.status_code}): {graph_message}",
            file=err,
        )
        print(
            "    Most common causes:\n"
            "      \u2022 Admin consent not granted for ServiceHealth.Read.All / ServiceMessage.Read.All\n"
            "      \u2022 Signed-in user lacks Service Support / Helpdesk / Global Reader / Global Admin role\n"
            "    See README.md \u2192 Troubleshooting.",
            file=err,
        )
        return 1

    print(
        f"\u274c  Unexpected response from Graph (HTTP {response.status_code}): "
        f"{response.text[:500]}",
        file=err,
    )
    return 1


def _safe_json(response: httpx.Response) -> object:
    try:
        return response.json()
    except ValueError:
        return None


def _extract_graph_error(body: object) -> str:
    if isinstance(body, dict):
        err = body.get("error")
        if isinstance(err, dict):
            code = err.get("code") or "<no-code>"
            message = err.get("message") or "<no-message>"
            return f"{code} \u2014 {message}"
    return "<no error envelope>"
