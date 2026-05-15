"""Configuration loaded from environment variables.

All settings are intentionally read once at startup so behaviour is deterministic
across the lifetime of the MCP server process.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_GRAPH_SCOPE = "https://graph.microsoft.com/.default"
TOKEN_CACHE_NAME = "m365-svc-comms-mcp"

# Microsoft's well-known multi-tenant public client used by the Microsoft Graph
# PowerShell SDK. Any caller can use this client without registering their own
# Entra app; the signed-in admin still has to consent to the requested scopes
# (ServiceHealth.Read.All / ServiceMessage.Read.All) on first sign-in.
DEFAULT_PUBLIC_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

# "organizations" lets any work/school account in any Entra tenant sign in
# (consumer Microsoft accounts are excluded, which matches the Service
# Communications API's tenant-only scope).
DEFAULT_TENANT_ID = "organizations"


class ConfigError(RuntimeError):
    """Raised when required configuration is missing or malformed."""


@dataclass(frozen=True)
class AuthConfig:
    """Resolved authentication configuration.

    Attributes:
        tenant_id: Microsoft Entra tenant GUID, ``"organizations"``,
            ``"common"``, or a verified domain (e.g. ``contoso.onmicrosoft.com``).
            Defaults to ``"organizations"`` so the customer can sign in with any
            work/school account without pre-configuring a tenant.
        client_id: Microsoft Entra app-registration client GUID. Defaults to
            the Microsoft Graph PowerShell public client
            (``14d82eec-204b-4c2f-b7e8-296a70dab67e``) so customers do not have
            to register their own app.
        scopes: OAuth scopes to request. Defaults to the Graph ``.default``
            scope, which resolves to whatever permissions admin consent has
            granted to the client app on the customer's tenant.
        prefer_device_code: When ``True``, force the device-code flow regardless
            of whether an interactive browser is available. Useful for headless
            CI / SSH scenarios. Set via ``M365_AUTH_DEVICE_CODE=1``.
    """

    tenant_id: str = DEFAULT_TENANT_ID
    client_id: str = DEFAULT_PUBLIC_CLIENT_ID
    scopes: tuple[str, ...] = (DEFAULT_GRAPH_SCOPE,)
    prefer_device_code: bool = False
    using_default_client: bool = True

    @classmethod
    def from_env(cls, env: dict[str, str] | None = None) -> AuthConfig:
        """Build an :class:`AuthConfig` from environment variables.

        Reads ``M365_TENANT_ID``, ``M365_CLIENT_ID``, and the optional
        ``M365_AUTH_DEVICE_CODE`` flag. Both ``M365_TENANT_ID`` and
        ``M365_CLIENT_ID`` are now optional and fall back to safe defaults
        (the Microsoft Graph PowerShell multi-tenant client) so the server
        runs out of the box without an Entra app registration.
        """

        e = env if env is not None else os.environ

        tenant = (e.get("M365_TENANT_ID") or "").strip() or DEFAULT_TENANT_ID
        client = (e.get("M365_CLIENT_ID") or "").strip() or DEFAULT_PUBLIC_CLIENT_ID

        device_code_flag = (e.get("M365_AUTH_DEVICE_CODE") or "").strip().lower()
        prefer_device_code = device_code_flag in {"1", "true", "yes", "on"}

        return cls(
            tenant_id=tenant,
            client_id=client,
            prefer_device_code=prefer_device_code,
            using_default_client=client == DEFAULT_PUBLIC_CLIENT_ID,
        )
