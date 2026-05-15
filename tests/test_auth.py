"""Tests for :mod:`m365_service_comms_mcp.auth`."""

from __future__ import annotations

import time

from azure.core.credentials import AccessToken, TokenCredential

from m365_service_comms_mcp.auth import GraphAuthProvider
from m365_service_comms_mcp.config import DEFAULT_GRAPH_SCOPE, AuthConfig


class _FakeCredential(TokenCredential):
    def __init__(self, *, token: str = "fake-token", expires_in: int = 3600) -> None:
        self.token = token
        self.expires_in = expires_in
        self.calls: list[tuple[str, ...]] = []

    def get_token(self, *scopes: str, **kwargs: object) -> AccessToken:  # type: ignore[override]
        self.calls.append(scopes)
        return AccessToken(self.token, int(time.time()) + self.expires_in)


def _make_config(**overrides: object) -> AuthConfig:
    base: dict[str, object] = {"tenant_id": "tenant", "client_id": "client"}
    base.update(overrides)
    return AuthConfig(**base)  # type: ignore[arg-type]


def test_get_token_uses_injected_factory() -> None:
    fake = _FakeCredential(token="abc-123")
    provider = GraphAuthProvider(config=_make_config(), _factory=lambda _cfg: fake)

    access = provider.get_token()

    assert access.token == "abc-123"
    assert fake.calls == [(DEFAULT_GRAPH_SCOPE,)]


def test_get_token_caches_credential_across_calls() -> None:
    factory_calls = 0

    def factory(_cfg: AuthConfig) -> TokenCredential:
        nonlocal factory_calls
        factory_calls += 1
        return _FakeCredential()

    provider = GraphAuthProvider(config=_make_config(), _factory=factory)
    provider.get_token()
    provider.get_token()
    provider.get_token()

    assert factory_calls == 1


def test_get_token_passes_configured_scopes() -> None:
    fake = _FakeCredential()
    config = _make_config(
        scopes=(
            "https://graph.microsoft.com/ServiceHealth.Read.All",
            "https://graph.microsoft.com/ServiceMessage.Read.All",
        )
    )
    provider = GraphAuthProvider(config=config, _factory=lambda _cfg: fake)

    provider.get_token()

    assert fake.calls == [
        (
            "https://graph.microsoft.com/ServiceHealth.Read.All",
            "https://graph.microsoft.com/ServiceMessage.Read.All",
        )
    ]
