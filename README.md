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

### Zero-setup quickstart (no app registration required)

You don't need to register your own Entra app — by default, the server uses the
**Microsoft Graph PowerShell** well-known multi-tenant public client
(`14d82eec-204b-4c2f-b7e8-296a70dab67e`) so any admin can sign in via browser
and grant consent on first use.

```jsonc
// .vscode/mcp.json (or any MCP-compatible client config)
{
  "servers": {
    "m365-svc-comms": {
      "type": "stdio",
      "command": "uvx",
      "args": ["m365-service-comms-mcp"]
    }
  }
}
```

Then ask your agent:

> *List the M365 services and tell me which ones are degraded.*

A browser will open the first time, prompting you to sign in to your tenant.
On first sign-in, an admin must grant consent for the
`ServiceHealth.Read.All` and `ServiceMessage.Read.All` Graph permissions —
see [Granting admin consent](#granting-admin-consent) below for what to expect
and how to handle the common roadblocks. Subsequent runs reuse the cached
token — no browser unless the token expires.

## Granting admin consent

The `ServiceHealth.Read.All` and `ServiceMessage.Read.All` Graph scopes are
**admin-only** — Microsoft requires a tenant administrator to consent to them
before any user can call the API. There are four paths, ranked by how
locked-down your tenant is.

### Path 1 — You are the admin (one-click consent during sign-in)

This is the default flow. The first time you (or anyone) runs
`uvx m365-service-comms-mcp --auth-test`:

1. Browser opens to `login.microsoftonline.com`.
2. Sign in with an account that holds **Global Administrator**, **Privileged
   Role Administrator**, or **Cloud Application Administrator** (any of these
   can grant tenant-wide consent).
3. You'll see a "Permissions requested" dialog listing:
   - **Read service health** — `ServiceHealth.Read.All`
   - **Read service messages** — `ServiceMessage.Read.All`
   - Plus the standard sign-in scopes (`openid`, `profile`, `email`, `offline_access`)
4. Check the **"Consent on behalf of your organization"** box at the bottom —
   this is the critical step. Without it, only your own account is consented.
5. Click **Accept**.

That's it. All users in the tenant who hold one of the
[required Service Communications roles](#required-roles-for-the-signed-in-user)
can now run the server.

> **Already consented to Microsoft Graph PowerShell?** Many tenants have
> already granted tenant-wide consent to the Microsoft Graph PowerShell client
> for other tooling. If your tenant's already consented to the same scopes
> for app `14d82eec-204b-4c2f-b7e8-296a70dab67e`, the consent dialog won't
> appear at all and you go straight to a token.

### Path 2 — You are NOT the admin (request consent)

If you sign in and see a message like *"AADSTS65001: The user or
administrator has not consented to use the application"* or a "Need admin
approval" page:

1. Click **"Have an admin account? Sign in with that account"** to switch users
   directly, OR
2. Click **"Request approval"** to send a consent request to your tenant's
   admins through the standard Microsoft Entra request flow.

While you wait, run with `--demo` to keep moving:

```sh
uvx m365-service-comms-mcp --demo
```

### Path 3 — Pre-consent via admin URL (no sign-in needed first)

Admins can grant consent before any user even tries to sign in by visiting a
purpose-built consent URL. Useful for automated tenant onboarding or when an
admin wants to grant consent without using the MCP server themselves.

For the **default** Microsoft Graph PowerShell client (no app registration
needed):

```
https://login.microsoftonline.com/<your-tenant-id>/adminconsent?client_id=14d82eec-204b-4c2f-b7e8-296a70dab67e
```

Replace `<your-tenant-id>` with your tenant's GUID, GUID alias, or
verified domain (e.g. `contoso.onmicrosoft.com`). When the admin opens the URL
and signs in, they'll see the same consent dialog as Path 1; they accept and
the whole tenant is consented in one shot.

For **your own Entra app registration**, swap the `client_id` for your app's
Application (client) ID.

### Path 4 — Microsoft Graph PowerShell (CLI, no browser)

Best for admins who already manage their tenant from PowerShell, or for
automated tenant-onboarding scripts. Requires the
[Microsoft Graph PowerShell SDK](https://learn.microsoft.com/powershell/microsoftgraph/installation):

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Then sign in as a **Global / Privileged Role / Cloud Application Administrator**
and grant the two scopes tenant-wide:

```powershell
# Sign in. Browser pops once; subsequent CLI consent is silent.
Connect-MgGraph -Scopes 'Application.ReadWrite.All','DelegatedPermissionGrant.ReadWrite.All' -TenantId <your-tenant-id>

# Resolve the two service principals we'll wire together.
$client   = Get-MgServicePrincipal -Filter "appId eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'"  # Microsoft Graph PowerShell client
if (-not $client) { $client = New-MgServicePrincipal -AppId '14d82eec-204b-4c2f-b7e8-296a70dab67e' }
$resource = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"  # Microsoft Graph itself

# Grant tenant-wide admin consent for the two delegated scopes.
New-MgOauth2PermissionGrant -BodyParameter @{
    ClientId    = $client.Id
    ConsentType = 'AllPrincipals'   # tenant-wide; use 'Principal' + PrincipalId for single-user
    ResourceId  = $resource.Id
    Scope       = 'ServiceHealth.Read.All ServiceMessage.Read.All'
}
```

For **your own Entra app registration**, swap the first `appId` filter for
your app's Application (client) ID.

To verify the grant landed:

```powershell
Get-MgOauth2PermissionGrant -Filter "clientId eq '$($client.Id)'" |
    Where-Object { $_.Scope -match 'ServiceHealth|ServiceMessage' }
```

To revoke later:

```powershell
$grant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($client.Id)'" |
    Where-Object { $_.Scope -match 'ServiceHealth|ServiceMessage' }
Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id
```

### Verifying consent landed

After consent is granted, you can confirm in the
[Microsoft Entra admin center](https://entra.microsoft.com):

1. **Identity → Applications → Enterprise applications**
2. Search for `Microsoft Graph PowerShell` (or your custom app name)
3. Click the app → **Permissions**
4. You should see `ServiceHealth.Read.All` and `ServiceMessage.Read.All` listed
   under "Admin consent" with a **Granted for &lt;tenant&gt;** status.

### Required roles for the signed-in user

Even after admin consent is granted, the signed-in user calling the API must
hold one of these directory roles (these are **API-side** restrictions
enforced by Microsoft Graph independently of OAuth scopes):

- Service Support Administrator
- Helpdesk Administrator
- Global Reader
- Global Administrator

If your sign-in succeeds but Graph returns `403 Authorization_RequestDenied`,
the consent is fine but the user lacks one of these roles. Add them via
**Microsoft Entra admin center → Identity → Roles & admins**.

### Revoking consent

To remove consent later (e.g. you're done evaluating the server):

1. **Microsoft Entra admin center → Identity → Applications → Enterprise applications**
2. Find the app, click it, then **Properties → Delete**.

This removes the consent grant and revokes any cached tokens for the app on
that tenant.

### Try it without any tenant (`--demo` mode)

Verify the MCP wire protocol works with your AI client *before* signing in:

```jsonc
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

You should see 3 services returned, with **Microsoft Teams** flagged as
`serviceDegradation` (canned data).

### Use your own Entra app registration (optional)

If you want a dedicated audit identity in your tenant (so audit logs show
your app's display name instead of "Microsoft Graph PowerShell"), register
your own Entra app following [Entra app registration](#entra-app-registration)
below, then set `M365_TENANT_ID` and `M365_CLIENT_ID`.

### Verify your setup

```sh
uvx m365-service-comms-mcp --auth-test
```

You should see (using defaults):

```
Tenant ID : organizations
Client ID : 14d82eec-204b-4c2f-b7e8-296a70dab67e
            (using the Microsoft Graph PowerShell public client \u2014 no Entra app registration needed)
Auth flow : interactive-browser (default)
Acquiring access token \u2026
\u2713  Token acquired.
Probing https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews?$top=1 \u2026
\u2713  Graph responded HTTP 200 (returned 1 healthOverview record(s)).
Auth test passed.
```

## Recommended testing tenant

If your primary tenant is locked down (typical for large organizations
including Microsoft itself), use a free
[**Microsoft 365 Developer Program**](https://developer.microsoft.com/microsoft-365/dev-program)
sandbox tenant. It comes with:

- An E5 license
- 25 pre-provisioned admin and user accounts
- Full Global Administrator rights for the tenant owner
- Auto-renewing 90-day subscription as long as you stay active

Sign up takes about 10 minutes.

## Entra app registration (optional, advanced)

You only need this if you want a dedicated audit identity instead of the
default "Microsoft Graph PowerShell" client. The walkthrough below assumes the
[Microsoft Entra admin center](https://entra.microsoft.com) UI as of 2026.

1. Go to **Microsoft Entra admin center \u2192 Identity \u2192 Applications \u2192
   App registrations \u2192 + New registration**.
2. **Name**: `m365-service-comms-mcp` (or whatever you like).
3. **Supported account types**: *Accounts in this organizational directory only
   (single tenant)*.
4. **Redirect URI**: select **Public client/native (mobile & desktop)** and
   enter `http://localhost`.
5. Click **Register**. Copy the **Application (client) ID** and **Directory
   (tenant) ID** from the Overview page \u2014 these become `M365_CLIENT_ID` and
   `M365_TENANT_ID` below.
6. Go to **API permissions \u2192 + Add a permission \u2192 Microsoft Graph \u2192 Delegated
   permissions**. Add:
   - `ServiceHealth.Read.All`
   - `ServiceMessage.Read.All`
7. Click **Grant admin consent for &lt;your tenant&gt;** and confirm. Both
   permissions should now show **Granted**.

The signed-in user must hold one of these directory roles (regardless of
whether you use the default client or your own app):

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

Add to `~/.copilot/mcp-config.json` (the file already exists; replace the empty
`{ "mcpServers": {} }` placeholder):

```jsonc
{
  "mcpServers": {
    "m365-svc-comms": {
      "type": "stdio",
      "command": "uvx",
      "args": ["m365-service-comms-mcp@latest"]
    }
  }
}
```

Restart `copilot` and verify with the `/mcp` slash command.

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

### `--auth-test` fails with `Authorization_RequestDenied` or `AADSTS65001`

You haven't granted **admin consent** yet. See
[Granting admin consent](#granting-admin-consent) for the three paths
(one-click during sign-in, request-from-admin, or pre-consent URL).

If you already tried admin consent and it still fails, double-check that
you ticked the **"Consent on behalf of your organization"** checkbox in the
consent dialog — without it, only your individual user is consented, and the
server will fail for any other user.

### `--auth-test` fails with `Forbidden` even after admin consent

Admin consent is in place but the signed-in user does not hold one of the
[required directory roles](#required-roles-for-the-signed-in-user)
(Service Support Admin / Helpdesk Admin / Global Reader / Global Admin). Add
the user to one of those roles via
**Microsoft Entra admin center → Identity → Roles & admins**.

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
