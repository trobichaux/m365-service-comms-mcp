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

Viewpoint writes (mark read / archive / favorite) are also supported in demo
mode against an in-memory state dict so users can exercise the full v0.2 tool
surface without a real tenant.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from typing import Any

from .errors import GraphError


class DemoGraphClient:
    """In-memory Graph stand-in. Returns deterministic sample data."""

    def __init__(self) -> None:
        # Per-message viewpoint state, populated lazily on first write. We avoid
        # seeding defaults so test assertions can distinguish "never touched"
        # from "explicitly set to False".
        self._viewpoint: dict[str, dict[str, bool]] = {}

    async def aclose(self) -> None:
        """No resources to release."""

    async def list_health_overviews(self, *, top: int = 25) -> dict[str, Any]:
        return {"value": _HEALTH_OVERVIEWS[:top]}

    async def get_health_overview(
        self,
        service_id: str,
        *,
        expand_issues: bool = True,
    ) -> dict[str, Any]:
        normalized = service_id.strip().lower()
        for overview in _HEALTH_OVERVIEWS:
            if overview["id"].lower() == normalized:
                result = dict(overview)
                if expand_issues:
                    result["issues"] = list(_ISSUES_BY_SERVICE.get(overview["id"], ()))
                return result

        raise GraphError(
            status_code=404,
            code="ResourceNotFound",
            message=(
                f"Demo dataset has no health overview with id '{service_id}'. "
                "Try one of: " + ", ".join(o["id"] for o in _HEALTH_OVERVIEWS)
            ),
        )

    async def list_issues(
        self,
        *,
        top: int = 25,
        filter_: str | None = None,
        orderby: str | None = "lastModifiedDateTime desc",
    ) -> dict[str, Any]:
        del filter_, orderby  # demo mode ignores filter/orderby
        return {"value": _ISSUES[:top]}

    async def get_issue(self, issue_id: str) -> dict[str, Any]:
        normalized = issue_id.strip().upper()
        for issue in _ISSUES:
            if issue["id"].upper() == normalized:
                return dict(issue)

        raise GraphError(
            status_code=404,
            code="ResourceNotFound",
            message=(
                f"Demo dataset has no issue with id '{issue_id}'. "
                "Try one of: " + ", ".join(i["id"] for i in _ISSUES)
            ),
        )

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
        raise GraphError(
            status_code=404,
            code="ResourceNotFound",
            message=(
                f"Demo dataset has no message with id '{message_id}'. "
                "Try one of: " + ", ".join(m["id"] for m in _MESSAGES)
            ),
        )

    async def set_messages_read(
        self,
        message_ids: list[str],
        *,
        read: bool,
    ) -> dict[str, Any]:
        self._apply_viewpoint(message_ids, key="read", value=read)
        return {"value": True}

    async def set_messages_archive(
        self,
        message_ids: list[str],
        *,
        archived: bool,
    ) -> dict[str, Any]:
        self._apply_viewpoint(message_ids, key="archived", value=archived)
        return {"value": True}

    async def set_messages_favorite(
        self,
        message_ids: list[str],
        *,
        favorite: bool,
    ) -> dict[str, Any]:
        self._apply_viewpoint(message_ids, key="favorite", value=favorite)
        return {"value": True}

    def viewpoint_state(self, message_id: str) -> dict[str, bool]:
        """Return the in-memory viewpoint state for ``message_id``.

        Exposed for tests; the real Graph backend has no equivalent helper.
        """

        return dict(self._viewpoint.get(message_id.strip().upper(), {}))

    def _apply_viewpoint(self, message_ids: list[str], *, key: str, value: bool) -> None:
        known_ids = {m["id"].upper() for m in _MESSAGES}
        for raw in message_ids:
            normalized = raw.strip().upper()
            if normalized not in known_ids:
                raise GraphError(
                    status_code=404,
                    code="ResourceNotFound",
                    message=(
                        f"Demo dataset has no message with id '{raw}'. "
                        "Try one of: " + ", ".join(sorted(known_ids))
                    ),
                )
            self._viewpoint.setdefault(normalized, {})[key] = value


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


_ISSUES: list[dict[str, Any]] = [
    {
        "id": "MT112233",
        "title": "Some users may be unable to send messages in Microsoft Teams",
        "service": "Microsoft Teams",
        "feature": "Teams components",
        "featureGroup": "Teams components",
        "classification": "incident",
        "status": "serviceRestored",
        "impactDescription": (
            "Affected users were unable to send 1:1 or group chat messages for the duration "
            "of the incident."
        ),
        "isResolved": True,
        "startDateTime": _iso(days_ago=3),
        "endDateTime": _iso(days_ago=2),
        "lastModifiedDateTime": _iso(days_ago=2),
    },
    {
        "id": "EX112299",
        "title": "Extended Recovery — Some Exchange Online users may experience delayed delivery",
        "service": "Exchange Online",
        "feature": "E-Mail and calendar access",
        "featureGroup": "E-Mail and calendar access",
        "classification": "incident",
        "status": "extendedRecovery",
        "impactDescription": (
            "Mail delivery for a small subset of mailboxes is delayed while backlog processing completes."
        ),
        "isResolved": False,
        "startDateTime": _iso(days_ago=1),
        "endDateTime": None,
        "lastModifiedDateTime": _iso(days_ago=0),
    },
    {
        "id": "PB445566",
        "title": "Power BI service is in extended recovery following capacity rebalancing",
        "service": "Power BI",
        "feature": "Service",
        "featureGroup": "Service",
        "classification": "advisory",
        "status": "extendedRecovery",
        "impactDescription": (
            "Capacity rebalancing has completed; tenants may see intermittent report load delays during the "
            "extended recovery window."
        ),
        "isResolved": False,
        "startDateTime": _iso(days_ago=1),
        "endDateTime": None,
        "lastModifiedDateTime": _iso(days_ago=0),
    },
]


_ISSUES_BY_SERVICE: dict[str, list[dict[str, Any]]] = {
    "microsoftteams": [issue for issue in _ISSUES if issue["service"] == "Microsoft Teams"],
    "Exchange": [issue for issue in _ISSUES if issue["service"] == "Exchange Online"],
    "PowerBIcom": [issue for issue in _ISSUES if issue["service"] == "Power BI"],
}


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
