"""Tests for :mod:`m365_service_comms_mcp.config`."""

from __future__ import annotations

import pytest

from m365_service_comms_mcp.config import (
    DEFAULT_PUBLIC_CLIENT_ID,
    DEFAULT_TENANT_ID,
    GRAPH_SCOPES,
    SERVICE_MESSAGE_VIEWPOINT_WRITE_SCOPE,
    AuthConfig,
)


def test_from_env_reads_explicit_vars() -> None:
    cfg = AuthConfig.from_env({"M365_TENANT_ID": "tenant-guid", "M365_CLIENT_ID": "client-guid"})
    assert cfg.tenant_id == "tenant-guid"
    assert cfg.client_id == "client-guid"
    assert cfg.scopes == GRAPH_SCOPES
    assert cfg.prefer_device_code is False
    assert cfg.write_enabled is False
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


@pytest.mark.parametrize("flag", ["1", "true", "TRUE", "Yes", "on"])
def test_from_env_write_flag_truthy(flag: str) -> None:
    cfg = AuthConfig.from_env({"M365_ENABLE_WRITE": flag})
    assert cfg.write_enabled is True


@pytest.mark.parametrize("flag", ["0", "false", "no", "off", "", "  "])
def test_from_env_write_flag_falsy(flag: str) -> None:
    cfg = AuthConfig.from_env({"M365_ENABLE_WRITE": flag})
    assert cfg.write_enabled is False


def test_effective_scopes_includes_write_scope_when_enabled() -> None:
    cfg = AuthConfig(write_enabled=True)
    assert cfg.effective_scopes == GRAPH_SCOPES + (SERVICE_MESSAGE_VIEWPOINT_WRITE_SCOPE,)


def test_effective_scopes_omits_write_scope_when_disabled() -> None:
    cfg = AuthConfig(write_enabled=False)
    assert cfg.effective_scopes == GRAPH_SCOPES
    assert SERVICE_MESSAGE_VIEWPOINT_WRITE_SCOPE not in cfg.effective_scopes


def test_from_env_write_override_true_beats_falsy_env() -> None:
    cfg = AuthConfig.from_env({"M365_ENABLE_WRITE": "0"}, write_enabled_override=True)
    assert cfg.write_enabled is True


def test_from_env_write_override_false_beats_truthy_env() -> None:
    cfg = AuthConfig.from_env({"M365_ENABLE_WRITE": "1"}, write_enabled_override=False)
    assert cfg.write_enabled is False


def test_from_env_write_override_none_reads_env() -> None:
    cfg = AuthConfig.from_env({"M365_ENABLE_WRITE": "1"}, write_enabled_override=None)
    assert cfg.write_enabled is True
