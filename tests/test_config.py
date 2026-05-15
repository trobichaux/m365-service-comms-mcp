"""Tests for :mod:`m365_service_comms_mcp.config`."""

from __future__ import annotations

import pytest

from m365_service_comms_mcp.config import DEFAULT_GRAPH_SCOPE, AuthConfig, ConfigError


def test_from_env_reads_required_vars() -> None:
    cfg = AuthConfig.from_env({"M365_TENANT_ID": "tenant-guid", "M365_CLIENT_ID": "client-guid"})
    assert cfg.tenant_id == "tenant-guid"
    assert cfg.client_id == "client-guid"
    assert cfg.scopes == (DEFAULT_GRAPH_SCOPE,)
    assert cfg.prefer_device_code is False


def test_from_env_strips_whitespace() -> None:
    cfg = AuthConfig.from_env({"M365_TENANT_ID": "  tenant  ", "M365_CLIENT_ID": "\tclient\n"})
    assert cfg.tenant_id == "tenant"
    assert cfg.client_id == "client"


@pytest.mark.parametrize(
    "env",
    [
        {},
        {"M365_TENANT_ID": "tenant"},
        {"M365_CLIENT_ID": "client"},
        {"M365_TENANT_ID": "", "M365_CLIENT_ID": "client"},
        {"M365_TENANT_ID": "tenant", "M365_CLIENT_ID": "   "},
    ],
)
def test_from_env_raises_when_missing(env: dict[str, str]) -> None:
    with pytest.raises(ConfigError) as info:
        AuthConfig.from_env(env)
    assert "Missing required environment variable" in str(info.value)


@pytest.mark.parametrize("flag", ["1", "true", "TRUE", "Yes", "on"])
def test_from_env_device_code_flag_truthy(flag: str) -> None:
    cfg = AuthConfig.from_env(
        {"M365_TENANT_ID": "t", "M365_CLIENT_ID": "c", "M365_AUTH_DEVICE_CODE": flag}
    )
    assert cfg.prefer_device_code is True


@pytest.mark.parametrize("flag", ["0", "false", "no", "off", ""])
def test_from_env_device_code_flag_falsy(flag: str) -> None:
    cfg = AuthConfig.from_env(
        {"M365_TENANT_ID": "t", "M365_CLIENT_ID": "c", "M365_AUTH_DEVICE_CODE": flag}
    )
    assert cfg.prefer_device_code is False
