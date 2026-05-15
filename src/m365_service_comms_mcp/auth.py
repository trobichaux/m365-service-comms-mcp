"""Delegated-only authentication for Microsoft Graph.

This module wraps ``azure.identity`` credentials to provide a uniform
:class:`GraphAuthProvider` that hands out access tokens for the Microsoft Graph
``serviceAnnouncement`` endpoints.

Two flows are supported:

- :class:`~azure.identity.InteractiveBrowserCredential` for interactive desktop
  use (default).
- :class:`~azure.identity.DeviceCodeCredential` for headless / SSH / CI use
  (selected when ``M365_AUTH_DEVICE_CODE=1`` or when ``DISPLAY``/``WAYLAND_DISPLAY``
  are unset on Linux).

Token caching uses :class:`~azure.identity.TokenCachePersistenceOptions`, which
stores tokens in the OS keyring (Windows DPAPI / macOS Keychain / Linux Secret
Service) with a file fallback when the keyring is unavailable.

**Stdio safety.** When the server is launched as an MCP child process the
parent owns ``stdout`` for JSON-RPC framing. We therefore install a
``prompt_callback`` on the device-code credential that writes to ``stderr``;
the default callback prints to ``stdout`` and would corrupt the transport.

v0.1 deliberately does **not** support client-secret or certificate
(application-permission) flows. Those land in v1.0 — see ``../docs/plan.md``.
"""

from __future__ import annotations

import datetime as _dt
import os
import platform
import sys
from dataclasses import dataclass, field
from typing import Protocol

from azure.core.credentials import AccessToken, TokenCredential
from azure.identity import (
    DeviceCodeCredential,
    InteractiveBrowserCredential,
    TokenCachePersistenceOptions,
)

from .config import TOKEN_CACHE_NAME, AuthConfig


class CredentialFactory(Protocol):
    """Function that builds a :class:`~azure.core.credentials.TokenCredential`.

    Defined as a Protocol so tests can inject a fake credential without monkey
    patching ``azure.identity``.
    """

    def __call__(self, config: AuthConfig) -> TokenCredential:
        """Construct a credential for the given configuration."""


@dataclass
class GraphAuthProvider:
    """Acquires Microsoft Graph access tokens for a configured Entra app.

    Construct one per process. The underlying credential is created lazily on
    first use so that import-time has no side effects.
    """

    config: AuthConfig
    _credential: TokenCredential | None = field(default=None, repr=False)
    _factory: CredentialFactory | None = field(default=None, repr=False)

    def get_token(self) -> AccessToken:
        """Acquire (or refresh) an access token for Microsoft Graph.

        Returns an :class:`~azure.core.credentials.AccessToken` whose ``token``
        attribute is the bearer string and ``expires_on`` is the Unix epoch
        expiry.
        """

        credential = self._ensure_credential()
        return credential.get_token(*self.config.scopes)

    def _ensure_credential(self) -> TokenCredential:
        if self._credential is None:
            factory = self._factory or _default_credential_factory
            self._credential = factory(self.config)
        return self._credential


def _default_credential_factory(config: AuthConfig) -> TokenCredential:
    """Pick the right credential based on config and runtime environment."""

    cache_options = TokenCachePersistenceOptions(name=TOKEN_CACHE_NAME)

    if config.prefer_device_code or _looks_headless():
        return DeviceCodeCredential(
            tenant_id=config.tenant_id,
            client_id=config.client_id,
            prompt_callback=_stderr_device_code_prompt,
            cache_persistence_options=cache_options,
        )

    return InteractiveBrowserCredential(
        tenant_id=config.tenant_id,
        client_id=config.client_id,
        cache_persistence_options=cache_options,
    )


def _stderr_device_code_prompt(
    verification_uri: str,
    user_code: str,
    expires_on: _dt.datetime,
) -> None:
    """Print device-code instructions to **stderr**, never stdout.

    The default ``DeviceCodeCredential`` prompt writes to stdout, which would
    corrupt the JSON-RPC stream when this server runs as a stdio MCP child
    process.
    """

    expires_local = expires_on.astimezone()
    message = (
        "\n[m365-service-comms-mcp] Sign-in required.\n"
        f"  Open: {verification_uri}\n"
        f"  Code: {user_code}\n"
        f"  Expires at: {expires_local.isoformat(timespec='seconds')}\n"
    )
    print(message, file=sys.stderr, flush=True)


def _looks_headless() -> bool:
    """Best-effort detection of headless environments.

    On Linux without ``DISPLAY`` or ``WAYLAND_DISPLAY`` we assume there is no
    browser to launch and fall back to the device-code flow. Windows and macOS
    almost always have a browser available, so we default to interactive there.

    .. note::
       We deliberately do **not** treat "stdin/stdout are pipes" as a headless
       signal. When this server is launched by an MCP client (VS Code, Claude
       Desktop, Cursor) over stdio, neither stream is a TTY \u2014 but the user is
       sitting at a desktop with a perfectly good browser. Customers can still
       force device code via ``M365_AUTH_DEVICE_CODE=1``.
    """

    return platform.system() == "Linux" and not (
        os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY")
    )
