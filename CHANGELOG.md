# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.3] - 2026-05-15

### Fixed

- **Stdio safety \u2014 device-code prompt no longer corrupts the JSON-RPC stream.**
  The default `azure-identity` `DeviceCodeCredential` printed sign-in
  instructions to **stdout**, which broke MCP clients connected over stdio.
  We now install a custom `prompt_callback` that writes to **stderr**.
- **Stdio safety \u2014 incorrect headless detection.** `_looks_headless()`
  previously fell back to device-code mode whenever `stdin`/`stdout` weren't
  TTYs. That's *always* the case for stdio MCP servers, so users on a
  desktop with a perfectly good browser were unnecessarily routed through
  device code. Removed the TTY check; only Linux without
  `DISPLAY`/`WAYLAND_DISPLAY` now triggers headless fallback. Customers can
  always force device code with `M365_AUTH_DEVICE_CODE=1`.
- **Stdio safety \u2014 logging now pinned to stderr.** `run_stdio_server`
  installs a root-logger `StreamHandler(stream=sys.stderr)` so any output
  from `azure-identity`, `msal`, `httpx`, or `tenacity` is guaranteed to
  miss the JSON-RPC stream.
- **Tighter consent dialog.** We now request `ServiceHealth.Read.All` and
  `ServiceMessage.Read.All` explicitly instead of the broad `.default`
  scope, so the admin-consent prompt only lists the two permissions we
  actually use (instead of the entire scope list of the Microsoft Graph
  PowerShell client).
- **`Retry-After` is now actually honored** on HTTP 429 / 5xx responses
  (the docstring previously claimed this but the wait strategy ignored the
  header).
- **`GraphError.args` is now populated** so logging frameworks that
  introspect `exc.args` see the rendered message (previously empty due to
  `@dataclass` overriding `__init__`).

### Changed

- `GraphClient.__init__` parameter `max_retries` renamed to `max_attempts`
  (it controlled the attempt count, not the retry count). Old name dropped
  outright since this is preview-grade and unreleased outside our test
  install.
- Console `prog` name in `--help` is now `m365-service-comms-mcp` (matching
  the package name) instead of the legacy `m365-svc-comms-mcp` short form.
- `--auth-test --demo` now exits with status 2 and a clear error message
  instead of silently ignoring `--demo`.
- `User-Agent` now includes the actual package version
  (`m365-service-comms-mcp/0.1.3`) rather than a hardcoded `0.1`.

### Removed

- `ConfigError` (no longer raised after 0.1.2 made all env vars optional).
- Unused `DEMO_TENANT_TIMESTAMP` constant.
- The `using_default_client` field on `AuthConfig` is now a `@property` so
  it's always correct regardless of how the dataclass is constructed.

## [0.1.2] - 2026-05-15

### Added

- **Zero-setup default client.** `M365_TENANT_ID` and `M365_CLIENT_ID` are now
  optional. When unset, the server falls back to the well-known **Microsoft
  Graph PowerShell** multi-tenant public client
  (`14d82eec-204b-4c2f-b7e8-296a70dab67e`) and tenant
  ID `organizations`, so customers can run `uvx m365-service-comms-mcp
  --auth-test` and complete sign-in via browser without registering their own
  Entra app. The signed-in user still must hold an admin role and grant
  consent to `ServiceHealth.Read.All` / `ServiceMessage.Read.All` on first
  sign-in.
- `--auth-test` output now indicates when the default public client is in use.

### Changed

- `AuthConfig.from_env` no longer raises on missing env vars; it applies
  defaults instead. To restore the original "fail fast" behaviour, set both
  variables explicitly.

## [0.1.1] - 2026-05-15

### Fixed

- Console script `m365-service-comms-mcp` now matches the package name, so
  `uvx m365-service-comms-mcp` works without `--from` redirection. The original
  `m365-svc-comms-mcp` script remains as an alias for backward compatibility.

## [0.1.0] - 2026-05-15

### Added

- Initial preview release.
- Three MCP tools wrapping the Microsoft Graph Service Communications API:
  - `list_service_health` — current M365 service health overview.
  - `list_message_center_posts` — Message Center posts with optional category
    and severity filters.
  - `get_message_center_post` — full post body and metadata by ID.
- Delegated authentication via `azure-identity`
  (`InteractiveBrowserCredential` + `DeviceCodeCredential` fallback).
- `--auth-test` console flag that probes Graph end-to-end and surfaces
  actionable error messages.
- `--demo` console flag that returns canned sample responses without calling
  Microsoft Graph — enables end-to-end MCP wire-protocol verification on
  tenants where admin consent cannot be granted.
- OS-keyring–backed token cache (Windows DPAPI / macOS Keychain /
  Linux Secret Service) with file-permission–restricted fallback.
- Exponential-backoff retry on 429 / 5xx via `tenacity`.
- 68 unit tests; 80% line coverage overall, 100% on tool code.

[Unreleased]: https://github.com/trobichaux/m365-service-comms-mcp/compare/v0.1.3...HEAD
[0.1.3]: https://github.com/trobichaux/m365-service-comms-mcp/releases/tag/v0.1.3
[0.1.2]: https://github.com/trobichaux/m365-service-comms-mcp/releases/tag/v0.1.2
[0.1.1]: https://github.com/trobichaux/m365-service-comms-mcp/releases/tag/v0.1.1
[0.1.0]: https://github.com/trobichaux/m365-service-comms-mcp/releases/tag/v0.1.0
