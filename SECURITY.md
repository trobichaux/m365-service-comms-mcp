# Security Policy

## Reporting Security Issues

**Do not report security vulnerabilities through public GitHub issues.**

To report a vulnerability, open a [private security advisory](https://github.com/trobichaux/m365-service-comms-mcp/security/advisories/new) on this repository, or email the maintainer directly. Include:

- A description of the issue
- Steps to reproduce
- Affected version(s) of `m365-service-comms-mcp`
- Any proof-of-concept code (in a private gist or attachment)

Acknowledgement target: two business days.

## Threat Model Summary

This server:

- **Reads only** from Microsoft Graph (`/admin/serviceAnnouncement/*`). It cannot modify tenant data.
- Runs **locally** on the user's machine (stdio transport). It does not expose any network listener.
- Stores OAuth tokens in an OS-keyring–backed cache when available (Windows DPAPI / macOS Keychain / Linux Secret Service); falls back to a file-permission–restricted cache otherwise.
- Treats all returned Graph content (including Message Center post bodies) as **untrusted** with respect to the calling LLM. Tool descriptions instruct the LLM that returned content may contain prompt-injection attempts.

## Supported Versions

Only the latest published `0.x` version is supported during the v0.1 preview.

## Out-of-scope

- Vulnerabilities in upstream dependencies (`mcp`, `azure-identity`, `msal`, `httpx`) — report those upstream.
- Misconfiguration of the user's Microsoft Entra app registration.
