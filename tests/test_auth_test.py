"""Tests for :mod:`m365_service_comms_mcp.auth_test` (the ``--auth-test`` flow)."""

from __future__ import annotations

import io
import time

import pytest
from azure.core.credentials import AccessToken
from pytest_httpx import HTTPXMock

from m365_service_comms_mcp.auth import GraphAuthProvider
from m365_service_comms_mcp.auth_test import _GRAPH_HEALTH_PROBE, run_auth_test
from m365_service_comms_mcp.config import AuthConfig


class _StubProvider(GraphAuthProvider):
    """Provider that returns a fixed token without touching the network."""

    def __init__(self, token: str = "stub-token", *, raises: Exception | None = None) -> None:
        super().__init__(config=AuthConfig(tenant_id="t", client_id="c"))
        self._stub_token = token
        self._raises = raises

    def get_token(self) -> AccessToken:  # type: ignore[override]
        if self._raises is not None:
            raise self._raises
        return AccessToken(self._stub_token, int(time.time()) + 3600)


@pytest.fixture
def env_set(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("M365_TENANT_ID", "tenant-1")
    monkeypatch.setenv("M365_CLIENT_ID", "client-1")
    monkeypatch.delenv("M365_AUTH_DEVICE_CODE", raising=False)


def test_returns_1_when_token_acquisition_fails(env_set: None) -> None:
    err = io.StringIO()
    provider = _StubProvider(raises=RuntimeError("user closed browser"))

    rc = run_auth_test(out=io.StringIO(), err=err, auth_provider=provider)

    assert rc == 1
    assert "Failed to acquire token" in err.getvalue()
    assert "user closed browser" in err.getvalue()


def test_uses_defaults_when_env_unset(
    monkeypatch: pytest.MonkeyPatch,
    httpx_mock: HTTPXMock,
) -> None:
    monkeypatch.delenv("M365_TENANT_ID", raising=False)
    monkeypatch.delenv("M365_CLIENT_ID", raising=False)
    httpx_mock.add_response(
        url=_GRAPH_HEALTH_PROBE,
        method="GET",
        json={"value": []},
    )
    out = io.StringIO()

    rc = run_auth_test(out=out, err=io.StringIO(), auth_provider=_StubProvider())

    assert rc == 0
    text = out.getvalue()
    assert "organizations" in text
    assert "Microsoft Graph PowerShell public client" in text


def test_returns_0_on_successful_graph_call(
    env_set: None,
    httpx_mock: HTTPXMock,
) -> None:
    httpx_mock.add_response(
        url=_GRAPH_HEALTH_PROBE,
        method="GET",
        json={"value": [{"service": "Exchange Online", "status": "serviceOperational"}]},
        status_code=200,
    )
    out = io.StringIO()

    rc = run_auth_test(out=out, err=io.StringIO(), auth_provider=_StubProvider())

    assert rc == 0
    assert "Auth test passed" in out.getvalue()
    assert "returned 1 healthOverview record" in out.getvalue()


def test_returns_1_on_graph_403_with_actionable_message(
    env_set: None,
    httpx_mock: HTTPXMock,
) -> None:
    httpx_mock.add_response(
        url=_GRAPH_HEALTH_PROBE,
        method="GET",
        json={
            "error": {
                "code": "Authorization_RequestDenied",
                "message": "Insufficient privileges to complete the operation.",
            }
        },
        status_code=403,
    )
    err = io.StringIO()

    rc = run_auth_test(out=io.StringIO(), err=err, auth_provider=_StubProvider())

    assert rc == 1
    text = err.getvalue()
    assert "Graph rejected the token" in text
    assert "Authorization_RequestDenied" in text
    assert "Admin consent" in text


def test_returns_1_on_unexpected_status(
    env_set: None,
    httpx_mock: HTTPXMock,
) -> None:
    httpx_mock.add_response(
        url=_GRAPH_HEALTH_PROBE,
        method="GET",
        text="boom",
        status_code=500,
    )
    err = io.StringIO()

    rc = run_auth_test(out=io.StringIO(), err=err, auth_provider=_StubProvider())

    assert rc == 1
    assert "Unexpected response" in err.getvalue()
    assert "HTTP 500" in err.getvalue()
