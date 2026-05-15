# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/trobichaux/m365-service-comms-mcp/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/trobichaux/m365-service-comms-mcp/releases/tag/v0.1.0
