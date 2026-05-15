# m365-service-comms-mcp

> ⚠️ **Preview (v0.1).** Read-only [Model Context Protocol](https://modelcontextprotocol.io)
> server for the Microsoft Graph **Service Communications API**. Exposes M365
> service health and Message Center posts to AI agents (Claude, GitHub Copilot,
> Cursor, VS Code, Claude Desktop, etc.).

[![CI](https://github.com/trobichaux/m365-service-comms-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/trobichaux/m365-service-comms-mcp/actions/workflows/ci.yml)
[![CodeQL](https://github.com/trobichaux/m365-service-comms-mcp/actions/workflows/codeql.yml/badge.svg)](https://github.com/trobichaux/m365-service-comms-mcp/actions/workflows/codeql.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)

## Status

| Item | State |
|---|---|
| Version | `0.1.0` (preview) |
| Tools | 3 — `list_service_health`, `list_message_center_posts`, `get_message_center_post` |
| Auth | Delegated only (browser sign-in / device code) |
| Transport | stdio |
| Distribution | PyPI via `uvx` |
| Listing | GitHub MCP Registry (planned for v0.1.08) |

Application-permission auth, additional tools, Docker image, and additional
documentation are deferred to **v1.0**.

## Why this exists

There is no published MCP server today that wraps the Graph
[Service Communications API](https://learn.microsoft.com/en-us/graph/api/resources/service-communications-api-overview)
with delegated authentication. The closest alternatives:

- [`Softeria/ms-365-mcp-server`](https://github.com/Softeria/ms-365-mcp-server)
  — delegated auth, but **no service health / message center coverage**.
- [`okapi-ca/ms-365-admin-mcp-server`](https://github.com/okapi-ca/ms-365-admin-mcp-server)
  — covers service health and messages, but **application permissions only**
  (cannot be used by an admin signing in interactively).

This server fills the gap: **delegated auth + service health + message center**,
in Python so M365 admins on any platform can install it with a single
`uvx` command.

## Quickstart

### Try it without a tenant (`--demo` mode)

You can verify the MCP wire protocol works with your AI client *before*
registering an Entra app or granting admin consent. The `--demo` flag returns
canned sample data:

```jsonc
// .vscode/mcp.json (or any MCP-compatible client config)
{
  "servers": {
    "m365-svc-comms-demo": {
      "type": "stdio",
      "command": "uvx",
      "args": ["m365-service-comms-mcp", "--demo"]
    }
  }
}
```

Then ask your agent:

> *List the M365 services and tell me which ones are degraded.*

You should see 3 services returned, with **Microsoft Teams** flagged as
`serviceDegradation` (canned data).

### Run against a real tenant

The full setup is **5 steps**:

1. Pick a tenant where you can grant admin consent (see
   [Recommended testing tenant](#recommended-testing-tenant)).
2. Register a Microsoft Entra app (see
   [Entra app registration](#entra-app-registration)).
3. Install `uv`: `pip install uv` (or follow [official install instructions](https://docs.astral.sh/uv/getting-started/installation/)).
4. Configure your MCP client (see [MCP client configuration](#mcp-client-configuration)).
5. Verify with `uvx m365-service-comms-mcp --auth-test`.

## Recommended testing tenant

If your primary tenant is locked down (typical for large organizations
including Microsoft itself), use a free
[**Microsoft 365 Developer Program**](https://developer.microsoft.com/microsoft-365/dev-program)
sandbox tenant. It comes with:

- An E5 license
- 25 pre-provisioned admin and user accounts
- Full Global Administrator rights for the tenant owner
- Auto-renewing 90-day subscription as long as you stay active

Sign up takes about 10 minutes. Once approved, your tenant is ready to register
an Entra app and grant the Service Communications scopes.

## Entra app registration

You only need to do this once per tenant. The walkthrough below assumes the
[Microsoft Entra admin center](https://entra.microsoft.com) UI as of 2026.

1. Go to **Microsoft Entra admin center → Identity → Applications →
   App registrations → + New registration**.
2. **Name**: `m365-service-comms-mcp` (or whatever you like).
3. **Supported account types**: *Accounts in this organizational directory only
   (single tenant)*.
4. **Redirect URI**: select **Public client/native (mobile & desktop)** and
   enter `http://localhost`.
5. Click **Register**. Copy the **Application (client) ID** and **Directory
   (tenant) ID** from the Overview page — these become `M365_CLIENT_ID` and
   `M365_TENANT_ID` below.
6. Go to **API permissions → + Add a permission → Microsoft Graph → Delegated
   permissions**. Add:
   - `ServiceHealth.Read.All`
   - `ServiceMessage.Read.All`
7. Click **Grant admin consent for &lt;your tenant&gt;** and confirm. Both
   permissions should now show **Granted**.

The signed-in user must hold one of these directory roles:

- Service Support Administrator
- Helpdesk Administrator
- Global Reader
- Global Administrator

## MCP client configuration

Set these two environment variables in whatever MCP client you use:

```
M365_TENANT_ID=<directory-tenant-guid>
M365_CLIENT_ID=<application-client-guid>
```

### VS Code (Copilot Chat)

Create `.vscode/mcp.json` in your workspace, or add to your user `mcp.json`:

```jsonc
{
  "servers": {
    "m365-svc-comms": {
      "type": "stdio",
      "command": "uvx",
      "args": ["m365-service-comms-mcp"],
      "env": {
        "M365_TENANT_ID": "00000000-0000-0000-0000-000000000000",
        "M365_CLIENT_ID": "00000000-0000-0000-0000-000000000000"
      }
    }
  }
}
```

### Claude Desktop

Add to `claude_desktop_config.json` (Settings → Developer → Edit Config):

```jsonc
{
  "mcpServers": {
    "m365-svc-comms": {
      "command": "uvx",
      "args": ["m365-service-comms-mcp"],
      "env": {
        "M365_TENANT_ID": "00000000-0000-0000-0000-000000000000",
        "M365_CLIENT_ID": "00000000-0000-0000-0000-000000000000"
      }
    }
  }
}
```

### Cursor

Add to `~/.cursor/mcp.json`:

```jsonc
{
  "mcpServers": {
    "m365-svc-comms": {
      "command": "uvx",
      "args": ["m365-service-comms-mcp"],
      "env": {
        "M365_TENANT_ID": "00000000-0000-0000-0000-000000000000",
        "M365_CLIENT_ID": "00000000-0000-0000-0000-000000000000"
      }
    }
  }
}
```

### GitHub Copilot CLI

Add to `~/.github/copilot/mcp.json`:

```jsonc
{
  "servers": {
    "m365-svc-comms": {
      "type": "stdio",
      "command": "uvx",
      "args": ["m365-service-comms-mcp"],
      "env": {
        "M365_TENANT_ID": "00000000-0000-0000-0000-000000000000",
        "M365_CLIENT_ID": "00000000-0000-0000-0000-000000000000"
      }
    }
  }
}
```

### Headless / CI (device-code flow)

If your environment can't open a browser (SSH session, container, CI), set
`M365_AUTH_DEVICE_CODE=1`. The server will then print a code and a URL for you
to complete sign-in from a different device.

## Verify your setup

After configuring, run:

```sh
uvx m365-service-comms-mcp --auth-test
```

You should see:

```
Tenant ID : <your-tenant-guid>
Client ID : <your-app-client-guid>
Auth flow : interactive-browser (default)
Acquiring access token …
✓  Token acquired.
Probing https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews?$top=1 …
✓  Graph responded HTTP 200 (returned 1 healthOverview record(s)).
Auth test passed. ServiceHealth.Read.All is granted and admin consent is in place.
```

Then in your AI client, ask:

> *Use the m365-svc-comms server to list the current Microsoft 365 service health.*

You should see real service health data for your tenant.

## Tools reference

### `list_service_health`

List the current health status of every M365 service the tenant subscribes to.

| Input | Type | Default | Notes |
|---|---|---|---|
| `top` | `int` (1..50) | 25 | Maximum number of services to return |

Returns Graph
[`healthOverviews`](https://learn.microsoft.com/en-us/graph/api/serviceannouncement-list-healthoverviews)
records: `id`, `service`, `status` (`serviceOperational` / `serviceDegradation` /
`serviceInterruption` / `extendedRecovery` / `investigating` / etc.).

### `list_message_center_posts`

List Message Center posts ordered by most recently modified.

| Input | Type | Default | Notes |
|---|---|---|---|
| `top` | `int` (1..50) | 25 | Maximum number of posts to return |
| `category` | `planForChange` \| `preventOrFixIssue` \| `stayInformed` \| `unknownFutureValue` | none | Optional category filter |
| `severity` | `normal` \| `high` \| `critical` | none | Optional severity filter |

Returns Graph
[`messages`](https://learn.microsoft.com/en-us/graph/api/serviceannouncement-list-messages)
records (without bodies — use `get_message_center_post` for the full content).

### `get_message_center_post`

Fetch the full Message Center post body and metadata.

| Input | Type | Notes |
|---|---|---|
| `message_id` | `str` matching `^[Mm][Cc][0-9]{4,8}$` | e.g. `MC123456` |

Returns the full Graph
[`serviceUpdateMessage`](https://learn.microsoft.com/en-us/graph/api/resources/serviceupdatemessage)
record including the rendered HTML body.

## Troubleshooting

### `--auth-test` fails with `Authorization_RequestDenied`

You haven't granted **admin consent** for the API permissions in step 7 of
[Entra app registration](#entra-app-registration). Open your app registration
in the Entra admin center, go to **API permissions**, and click **Grant admin
consent for &lt;your tenant&gt;**.

### `--auth-test` fails with `Forbidden` even after admin consent

The signed-in user does not hold an admin role. See the role list at the end of
[Entra app registration](#entra-app-registration). Add the user to one of those
roles via **Microsoft Entra admin center → Identity → Roles & admins**.

### Browser doesn't open / I'm on SSH

Set `M365_AUTH_DEVICE_CODE=1` in the `env` block of your MCP client config.
The server will print a code + URL on first call; complete sign-in on any
device.

### `uvx: command not found`

Install `uv` first:

```sh
pip install uv
# or follow https://docs.astral.sh/uv/getting-started/installation/
```

If your corporate machine blocks `uv`, fall back to `pipx`:

```sh
pipx install m365-service-comms-mcp
```

…and replace `"command": "uvx"` with `"command": "m365-svc-comms-mcp"` and drop
the first `args` entry.

### Token cache problems on Linux

If you see errors about `Secret Service` or `keyring`, install the keyring
backend:

```sh
sudo apt install gnome-keyring  # Debian/Ubuntu
```

The server will fall back to a file-permission–restricted cache automatically
if the keyring is unavailable.

### Graph returns `429 TooManyRequests`

The server retries 429s with exponential backoff (4 attempts by default). If
you keep seeing this, you're likely hammering the API in a loop. Service
Communications has a soft limit of about 10 RPS per tenant.

### Locked-down corporate tenant won't let me grant consent

Use a free [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
sandbox tenant for testing. See [Recommended testing tenant](#recommended-testing-tenant).

If you don't want to set up a sandbox tenant, you can still verify the MCP
wire protocol with `--demo` mode (no tenant required).

## Known limitations (v0.1)

- **Read-only.** No tools for marking messages as read / archiving / favoriting
  yet (the Graph API supports it; v1.0 will).
- **3 tools.** `get_service_health` (per-service deep-dive), `list_service_issues`,
  `get_service_issue`, `get_incident_report`, and `summarize_my_tenant` are
  planned for v1.0.
- **Delegated auth only.** No client-secret or certificate flows. Backend
  monitoring scenarios will need to wait for v1.0.
- **stdio transport only.** No Streamable HTTP for hosted/multi-tenant
  scenarios yet.
- **No Docker image** on `mcr.microsoft.com` — install via PyPI/`uvx` only.

## Local development

```powershell
# from the repo root
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
ruff check .
ruff format --check .
pytest --cov
```

End-to-end smoke test against the demo client (no tenant required):

```sh
.\.venv\Scripts\m365-svc-comms-mcp.exe --demo
# then connect any MCP client to it via stdio
```

## License

[MIT](LICENSE) — chosen for broad compatibility:

- ✅ **Customer use**: commercial use, modification, redistribution, and
  incorporation into proprietary products are all permitted with the license
  notice retained.
- ✅ **Microsoft use**: MIT is on Microsoft's approved OSS-license list and is
  the license used by every Microsoft-authored MCP server in the
  [`microsoft/mcp`](https://github.com/microsoft/mcp) catalog (Azure MCP
  Server, Markitdown MCP, Fabric MCP) and by the comparable community
  Microsoft 365 MCP servers (Softeria/ms-365-mcp-server,
  okapi-ca/ms-365-admin-mcp-server).
- ✅ **GPL-compatible**, so downstream redistribution under copyleft licenses
  is permitted if a consumer chooses.

The license is declared in `pyproject.toml` using
[PEP 639](https://peps.python.org/pep-0639/) SPDX format
(`license = "MIT"`) and the `LICENSE` file is included in the published wheel.

## Security

See [SECURITY.md](SECURITY.md). Do not file security issues in the public
tracker — open a [private security advisory](https://github.com/trobichaux/m365-service-comms-mcp/security/advisories/new)
instead.

## Code of Conduct

See [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md). This project follows the
[Contributor Covenant v2.1](https://www.contributor-covenant.org/version/2/1/code_of_conduct/).
