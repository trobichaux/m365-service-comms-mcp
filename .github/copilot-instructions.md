# Copilot Instructions

A read-only **Model Context Protocol** server (Python, stdio transport) that wraps
the Microsoft Graph **Service Communications API** (`/admin/serviceAnnouncement/*`).
v0.1 ships three tools — `list_service_health`, `list_message_center_posts`,
`get_message_center_post` — with delegated auth only.

## Commands

```sh
pip install -e ".[dev]"

ruff check .
ruff format --check .       # CI uses --check; run `ruff format .` to fix
pytest --cov                # full suite with coverage
pytest tests/test_auth.py::test_get_token_calls_credential -x   # single test
m365-svc-comms-mcp --demo   # E2E with canned data, no tenant required
m365-svc-comms-mcp --auth-test   # exercise real token flow, then exit
```

CI matrix: Python 3.11 / 3.12 / 3.13 (`.github/workflows/ci.yml`). All three
checks above must pass before opening a PR.

## Architecture

The whole server is wired around one seam: **`GraphClientProtocol`**
(`src/m365_service_comms_mcp/graph_protocol.py`). Three implementations exist
and are swapped at the boundary:

| Backend | When used |
|---|---|
| `GraphClient` | Live Graph, used in production |
| `DemoGraphClient` | `--demo` mode — canned data, no auth |
| `_RecordingClient` (in `tests/test_tools.py`) | Unit tests — asserts exact args the tools pass |

`server.build_server(client)` constructs a `FastMCP` and calls
`tools.register_tools(mcp, client)`. Each tool lives in its own file under
`tools/` and exposes a `register(mcp, client)` function that wires the
Pydantic-validated handler.

Entry-point flag matrix (`__main__.py`): `--demo` and `--auth-test` are mutually
exclusive. `--auth-test` short-circuits before the MCP server is built.

**Retry / error handling** lives in `graph_client.py`: 429 + retryable 5xx →
up to `max_attempts=4` with `Retry-After`-honoring backoff, then raise
`GraphError`. Always build `GraphError` via `parse_graph_error()` so the Graph
error envelope (`code`, `message`, `request-id`) is preserved for the LLM.

## Stdio safety (read before adding any output)

The parent MCP client owns `stdout` for JSON-RPC framing — **any stray byte on
stdout corrupts the transport**. This shapes a few non-obvious choices:

- `server._configure_stderr_logging()` forces the root logger to stderr at
  WARNING.
- `auth._stderr_device_code_prompt` replaces the default `DeviceCodeCredential`
  prompt (which writes to stdout).
- `--demo` and `--auth-test` print to `sys.stderr` explicitly, not via
  bare `print()`.

If you add new output anywhere in the runtime path, send it to `stderr`.

## Auth conventions (v0.1 = delegated only)

- Default tenant is `"organizations"` and default client is the **Microsoft
  Graph PowerShell well-known public client** (`14d82eec-204b-4c2f-b7e8-296a70dab67e`)
  so users don't need their own Entra app. Don't change either default
  casually — it's the whole "zero-setup quickstart" story.
- Scopes are listed **explicitly** (`ServiceHealth.Read.All`,
  `ServiceMessage.Read.All`) instead of `.default` to keep the admin-consent
  dialog tight to the two permissions actually used.
- Tokens are cached via `TokenCachePersistenceOptions` (OS keyring with file
  fallback). Don't disable persistence — every restart re-prompting would break
  the MCP UX.
- Application-permission (client-secret / certificate) flows are explicitly
  deferred to v1.0. Don't add them in v0.1.

## Project conventions

- **Pytest treats warnings as errors** (`filterwarnings = ["error", ...]` in
  `pyproject.toml`, with one allow-listed `DeprecationWarning` from `msal`).
  New deprecations from dependencies will fail CI.
- **`MAX_TOP = 50`** (`tools/_constants.py`) caps every Graph `$top` parameter,
  intentionally below Graph's soft limit (500–1000) to protect the LLM's context
  window. Keep this cap when adding tools that page through Graph.
- Tool descriptions explicitly note that responses are
  **"untrusted with respect to the calling LLM"** — preserve that wording when
  adding new tools, since Message Center bodies are admin-authored HTML that
  could carry injection content.
- `ruff` config: line length 100, target `py311`, selected rule set includes
  `B`, `UP`, `SIM`, `RUF`. `E501` is ignored — long lines OK if `ruff format`
  is happy.
- Tests use `pytest-asyncio` in **`auto`** mode — `async def test_*` is enough,
  no `@pytest.mark.asyncio` needed.
- Conventional Commits required (`feat:`, `fix(auth):`, etc.) — see
  `CONTRIBUTING.md`.

## Adding a new tool

Per `CONTRIBUTING.md`, every new tool touches four spots:

1. `src/m365_service_comms_mcp/tools/<name>.py` exposing `register(mcp, client)`.
2. `tests/test_<name>.py` exercising it through `register_tools` + a
   `_RecordingClient` stand-in (mirror the pattern in `tests/test_tools.py`).
3. `src/m365_service_comms_mcp/tools/__init__.py` — call the new `register()`.
4. If it calls Graph endpoints not already on `GraphClientProtocol`, add the
   method there and implement it on **both** `GraphClient` and
   `DemoGraphClient` (otherwise `--demo` will break).
5. Update the `## Tools reference` section in `README.md` and add an entry under
   `[Unreleased]` in `CHANGELOG.md`.

## Release

Tag push (`vX.Y.Z`) → `publish.yml` publishes to PyPI via OIDC trusted publishing
(no long-lived tokens). Version is bumped in `pyproject.toml` **and**
`src/m365_service_comms_mcp/__init__.py` **and** `.mcp/server.json` — keep all
three in lockstep or the registry listing drifts from the published package.
