"""Tests for :mod:`m365_service_comms_mcp.config`."""

from __future__ import annotations

import pytest

from m365_service_comms_mcp.config import (
    DEFAULT_PUBLIC_CLIENT_ID,
    DEFAULT_TENANT_ID,
    GRAPH_SCOPES,
    AuthConfig,
)


def test_from_env_reads_explicit_vars() -> None:
    cfg = AuthConfig.from_env({"M365_TENANT_ID": "tenant-guid", "M365_CLIENT_ID": "client-guid"})
    assert cfg.tenant_id == "tenant-guid"
    assert cfg.client_id == "client-guid"
    assert cfg.scopes == GRAPH_SCOPES
    assert cfg.prefer_device_code is False
    assert cfg.using_default_client is False


def test_from_env_strips_whitespace() -> None:
    cfg = AuthConfig.from_env({"M365_TENANT_ID": "  tenant  ", "M365_CLIENT_ID": "\tclient\n"})
    assert cfg.tenant_id == "tenant"
    assert cfg.client_id == "client"


def test_from_env_falls_back_to_default_client_when_unset() -> None:
    cfg = AuthConfig.from_env({})
    assert cfg.client_id == DEFAULT_PUBLIC_CLIENT_ID
    assert cfg.tenant_id == DEFAULT_TENANT_ID
    assert cfg.using_default_client is True


def test_from_env_falls_back_to_default_client_when_blank() -> None:
    cfg = AuthConfig.from_env({"M365_TENANT_ID": "  ", "M365_CLIENT_ID": ""})
    assert cfg.client_id == DEFAULT_PUBLIC_CLIENT_ID
    assert cfg.tenant_id == DEFAULT_TENANT_ID
    assert cfg.using_default_client is True


def test_from_env_uses_custom_tenant_with_default_client() -> None:
    cfg = AuthConfig.from_env({"M365_TENANT_ID": "contoso.onmicrosoft.com"})
    assert cfg.tenant_id == "contoso.onmicrosoft.com"
    assert cfg.client_id == DEFAULT_PUBLIC_CLIENT_ID
    assert cfg.using_default_client is True


def test_from_env_explicit_client_disables_using_default_flag() -> None:
    cfg = AuthConfig.from_env({"M365_CLIENT_ID": "11111111-2222-3333-4444-555555555555"})
    assert cfg.client_id == "11111111-2222-3333-4444-555555555555"
    assert cfg.tenant_id == DEFAULT_TENANT_ID
    assert cfg.using_default_client is False


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
