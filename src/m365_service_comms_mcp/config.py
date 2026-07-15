"""Configuration loaded from environment variables.

All settings are intentionally read once at startup so behaviour is deterministic
across the lifetime of the MCP server process.
"""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Scopes used to call /admin/serviceAnnouncement/*. Requesting these explicitly
# (rather than the broader ``.default``) keeps the first-time consent dialog
# tight: the admin sees a 2-permission request, not the entire scope list of
# whatever client app we're using.
GRAPH_SCOPES: tuple[str, ...] = (
    "https://graph.microsoft.com/ServiceHealth.Read.All",
    "https://graph.microsoft.com/ServiceMessage.Read.All",
)

# Per-user write scope for the serviceUpdateMessage viewpoint actions
# (markRead/markUnread, archive/unarchive, favorite/unfavorite). Added to the
# requested scope list only when write mode is enabled so that the default
# install never triggers a consent dialog change.
SERVICE_MESSAGE_VIEWPOINT_WRITE_SCOPE = (
    "https://graph.microsoft.com/ServiceMessageViewpoint.Write"
)

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
        scopes: Base OAuth scopes to request. Defaults to the explicit Service
            Communications read scopes so the admin-consent dialog only lists
            the two permissions read-only callers actually use. Per-user write
            actions add an extra scope dynamically via :attr:`effective_scopes`.
        prefer_device_code: When ``True``, force the device-code flow regardless
            of whether an interactive browser is available. Useful for headless
            CI / SSH scenarios. Set via ``M365_AUTH_DEVICE_CODE=1``.
        write_enabled: When ``True``, viewpoint write tools are enabled and the
            additional ``ServiceMessageViewpoint.Write`` scope is requested.
            Set via ``M365_ENABLE_WRITE=1`` or the ``--enable-write`` CLI flag.
    """

    tenant_id: str = DEFAULT_TENANT_ID
    client_id: str = DEFAULT_PUBLIC_CLIENT_ID
    scopes: tuple[str, ...] = GRAPH_SCOPES
    prefer_device_code: bool = False
    write_enabled: bool = False

    @property
    def using_default_client(self) -> bool:
        """True iff ``client_id`` is the Microsoft Graph PowerShell well-known client."""

        return self.client_id == DEFAULT_PUBLIC_CLIENT_ID

    @property
    def effective_scopes(self) -> tuple[str, ...]:
        """Scopes actually requested from Entra at token-acquisition time.

        Equals :attr:`scopes` when ``write_enabled`` is False, otherwise
        appends :data:`SERVICE_MESSAGE_VIEWPOINT_WRITE_SCOPE`.
        """

        if self.write_enabled:
            return self.scopes + (SERVICE_MESSAGE_VIEWPOINT_WRITE_SCOPE,)
        return self.scopes

    @classmethod
    def from_env(
        cls,
        env: Mapping[str, str] | None = None,
        *,
        write_enabled_override: bool | None = None,
    ) -> AuthConfig:
        """Build an :class:`AuthConfig` from environment variables.

        Reads ``M365_TENANT_ID``, ``M365_CLIENT_ID``, ``M365_AUTH_DEVICE_CODE``,
        and ``M365_ENABLE_WRITE``. All are optional and fall back to safe
        defaults (the Microsoft Graph PowerShell multi-tenant client, no write
        scope) so the server runs out of the box without an Entra app
        registration.

        ``write_enabled_override`` lets the caller (typically ``__main__``)
        resolve the gate from CLI args + env in one place and pass the final
        decision in, so the CLI flag and the env var never disagree about
        which scopes to request vs. which tools to register.
        """

        e: Mapping[str, str] = env if env is not None else os.environ

        tenant = (e.get("M365_TENANT_ID") or "").strip() or DEFAULT_TENANT_ID
        client = (e.get("M365_CLIENT_ID") or "").strip() or DEFAULT_PUBLIC_CLIENT_ID

        device_code_flag = (e.get("M365_AUTH_DEVICE_CODE") or "").strip().lower()
        prefer_device_code = device_code_flag in {"1", "true", "yes", "on"}

        if write_enabled_override is None:
            write_flag = (e.get("M365_ENABLE_WRITE") or "").strip().lower()
            write_enabled = write_flag in {"1", "true", "yes", "on"}
        else:
            write_enabled = write_enabled_override

        return cls(
            tenant_id=tenant,
            client_id=client,
            prefer_device_code=prefer_device_code,
            write_enabled=write_enabled,
        )
