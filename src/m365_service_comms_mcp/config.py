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


class ConfigError(RuntimeError):
    """Raised when required configuration is missing or malformed."""


@dataclass(frozen=True)
class AuthConfig:
    """Resolved authentication configuration.

    Attributes:
        tenant_id: Microsoft Entra tenant GUID. Required.
        client_id: Microsoft Entra app-registration client GUID. Required.
        scopes: OAuth scopes to request. Defaults to the Graph ``.default`` scope,
            which resolves to whatever permissions the app registration has been
            granted admin consent for.
        prefer_device_code: When ``True``, force the device-code flow regardless
            of whether an interactive browser is available. Useful for headless
            CI / SSH scenarios. Set via ``M365_AUTH_DEVICE_CODE=1``.
    """

    tenant_id: str
    client_id: str
    scopes: tuple[str, ...] = (DEFAULT_GRAPH_SCOPE,)
    prefer_device_code: bool = False

    @classmethod
    def from_env(cls, env: dict[str, str] | None = None) -> AuthConfig:
        """Build an :class:`AuthConfig` from environment variables.

        Reads ``M365_TENANT_ID``, ``M365_CLIENT_ID``, and the optional
        ``M365_AUTH_DEVICE_CODE`` flag. Raises :class:`ConfigError` if either
        required variable is missing or empty.
        """

        e = env if env is not None else os.environ

        tenant = (e.get("M365_TENANT_ID") or "").strip()
        client = (e.get("M365_CLIENT_ID") or "").strip()

        missing = [
            name
            for name, value in (("M365_TENANT_ID", tenant), ("M365_CLIENT_ID", client))
            if not value
        ]
        if missing:
            raise ConfigError(
                "Missing required environment variable(s): "
                + ", ".join(missing)
                + ". See README.md \u2192 Configuration."
            )

        device_code_flag = (e.get("M365_AUTH_DEVICE_CODE") or "").strip().lower()
        prefer_device_code = device_code_flag in {"1", "true", "yes", "on"}

        return cls(
            tenant_id=tenant,
            client_id=client,
            prefer_device_code=prefer_device_code,
        )
