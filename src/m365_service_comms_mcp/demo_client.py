"""Canned-data Graph client used by ``--demo`` mode.

Why this exists:

The Microsoft Graph Service Communications API only returns useful data when the
caller is signed into an admin-grantable tenant. That is a problem for two
audiences:

1. Microsoft FTEs (and large-organization employees more generally) whose primary
   tenant blocks the required admin consent.
2. New users who want to verify the MCP wire protocol works with their LLM
   client *before* spending the time to register an Entra app and grant
   ``ServiceHealth.Read.All`` / ``ServiceMessage.Read.All``.

The canned responses below are shaped exactly like real Graph responses so the
LLM-facing experience (and our tool unit tests) match production. Service IDs,
``id`` fields, and timestamps are obviously fictional.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from typing import Any

DEMO_TENANT_TIMESTAMP = "2026-05-15T17:00:00Z"


class DemoGraphClient:
    """In-memory Graph stand-in. Returns deterministic sample data."""

    async def aclose(self) -> None:
        # No resources to release.
        return

    async def list_health_overviews(self, *, top: int = 25) -> dict[str, Any]:
        records = _HEALTH_OVERVIEWS[:top]
        return {"value": records}

    async def list_messages(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]:
        del filter_, orderby  # demo mode ignores filter/orderby
        return {"value": _MESSAGES[:top]}

    async def get_message(self, message_id: str) -> dict[str, Any]:
        normalized = message_id.strip().upper()
        for message in _MESSAGES:
            if message["id"].upper() == normalized:
                return {**message, "body": _MESSAGE_BODIES.get(message["id"], {})}
        # Mirror the live client's error envelope shape so calling tools see the
        # same not-found behaviour they would in production.
        from .errors import GraphError

        raise GraphError(
            status_code=404,
            code="ResourceNotFound",
            message=(
                f"Demo dataset has no message with id '{message_id}'. "
                "Try one of: " + ", ".join(m["id"] for m in _MESSAGES)
            ),
        )


def _iso(days_ago: int) -> str:
    return (datetime.now(UTC) - timedelta(days=days_ago)).strftime("%Y-%m-%dT%H:%M:%SZ")


_HEALTH_OVERVIEWS: list[dict[str, Any]] = [
    {
        "id": "Exchange",
        "service": "Exchange Online",
        "status": "serviceOperational",
        "feature": None,
        "featureGroup": None,
    },
    {
        "id": "microsoftteams",
        "service": "Microsoft Teams",
        "status": "serviceDegradation",
        "feature": "Teams components",
        "featureGroup": "Teams components",
    },
    {
        "id": "SharePoint",
        "service": "SharePoint Online",
        "status": "serviceOperational",
        "feature": None,
        "featureGroup": None,
    },
    {
        "id": "OneDriveForBusiness",
        "service": "OneDrive for Business",
        "status": "serviceOperational",
        "feature": None,
        "featureGroup": None,
    },
    {
        "id": "PowerBIcom",
        "service": "Power BI",
        "status": "extendedRecovery",
        "feature": "Service",
        "featureGroup": "Service",
    },
]


_MESSAGES: list[dict[str, Any]] = [
    {
        "id": "MC987654",
        "title": "(Updated) New Outlook calendar sharing experience rolling out",
        "category": "stayInformed",
        "severity": "normal",
        "tags": ["User impact", "New feature"],
        "isMajorChange": False,
        "actionRequiredByDateTime": None,
        "services": ["Exchange Online", "Microsoft 365 for the web"],
        "startDateTime": _iso(days_ago=10),
        "endDateTime": _iso(days_ago=-21),
        "lastModifiedDateTime": _iso(days_ago=2),
    },
    {
        "id": "MC987655",
        "title": "Plan for change: Teams meeting recording retention default change",
        "category": "planForChange",
        "severity": "high",
        "tags": ["Admin impact", "Retention"],
        "isMajorChange": True,
        "actionRequiredByDateTime": _iso(days_ago=-30),
        "services": ["Microsoft Teams"],
        "startDateTime": _iso(days_ago=5),
        "endDateTime": _iso(days_ago=-45),
        "lastModifiedDateTime": _iso(days_ago=1),
    },
    {
        "id": "MC987656",
        "title": "Action required: Update conditional access policies for SharePoint",
        "category": "preventOrFixIssue",
        "severity": "critical",
        "tags": ["Admin impact", "Security"],
        "isMajorChange": True,
        "actionRequiredByDateTime": _iso(days_ago=-7),
        "services": ["SharePoint Online", "Microsoft Entra"],
        "startDateTime": _iso(days_ago=14),
        "endDateTime": _iso(days_ago=-7),
        "lastModifiedDateTime": _iso(days_ago=0),
    },
]


_MESSAGE_BODIES: dict[str, dict[str, Any]] = {
    "MC987654": {
        "contentType": "html",
        "content": (
            "<p>The new Outlook calendar sharing experience is rolling out to all tenants. "
            "<strong>This is a sample message body</strong> served by the m365-service-comms-mcp "
            "demo dataset \u2014 the data is fictional and is provided so you can verify your MCP "
            "client end-to-end without an admin-grantable tenant.</p>"
        ),
    },
    "MC987655": {
        "contentType": "html",
        "content": (
            "<p>Effective in 30 days, Teams meeting recordings will default to a retention "
            "period of 60 days. Tenant administrators can override the default in the Teams admin center.</p>"
        ),
    },
    "MC987656": {
        "contentType": "html",
        "content": (
            "<p>A conditional access change is required to maintain SharePoint access for "
            "users assigned the Conditional Access Persistent Browser Session policy. "
            "Please review and update affected policies before the deadline.</p>"
        ),
    },
}
