"""Per-user viewpoint write tools for Message Center posts (v0.2+, gated).

These three tools are registered **only** when the server is started with the
write gate on (``--enable-write`` CLI flag or ``M365_ENABLE_WRITE=1`` env var).
That gating keeps the default install fully read-only — no extra tools, no
extra OAuth scope on the admin-consent dialog — and isolates the
``ServiceMessageViewpoint.Write`` consent change to users who explicitly opt
in.

Each tool wraps a paired Graph viewpoint endpoint with a boolean flag:

- ``set_message_center_posts_read``     → ``markRead`` / ``markUnread``
- ``set_message_center_posts_archived`` → ``archive`` / ``unarchive``
- ``set_message_center_posts_favorite`` → ``favorite`` / ``unfavorite``

All operations are **per-user only** (they update the signed-in user's view
state, not tenant-wide state). The server caps the per-call message id list
size as a safety limit on LLM-driven bulk mutations.
"""

from __future__ import annotations

from typing import Annotated, Any

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..graph_protocol import GraphClientProtocol

# Reuse the existing Message Center id shape (case-insensitive ``MC`` prefix +
# 4-8 digits) for each entry in the list.
_MESSAGE_ID_PATTERN = r"^[Mm][Cc][0-9]{4,8}$"

# Type alias enforcing the per-item pattern. Nested in the ``list[...]`` below
# so Pydantic validates every id, not just the list itself.
_MessageId = Annotated[
    str,
    Field(min_length=3, max_length=12, pattern=_MESSAGE_ID_PATTERN),
]

# Per-call cap on the number of message ids in a single viewpoint write.
# Smaller than MAX_TOP for reads: each call is a per-user mutation, and we want
# to bound the blast radius of an LLM that misinterprets a prompt. The Graph
# endpoint itself does not document a hard limit.
MAX_MESSAGE_IDS_PER_WRITE = 50

_MESSAGE_IDS_FIELD_DESCRIPTION = (
    "List of Message Center post ids to update (e.g. ['MC123456', 'MC234567']). "
    f"At least 1, at most {MAX_MESSAGE_IDS_PER_WRITE} ids per call. Duplicate "
    "ids are collapsed before the request is sent."
)

_READ_DESCRIPTION = (
    "Mark Microsoft 365 Message Center posts as read or unread for the SIGNED-IN "
    "USER ONLY. This does not change the tenant-wide read state — it only updates "
    "the per-user viewpoint that drives the unread badge in the admin center. Pass "
    "read=True to mark posts as read, read=False to mark them as unread."
)

_ARCHIVED_DESCRIPTION = (
    "Archive or unarchive Microsoft 365 Message Center posts for the SIGNED-IN "
    "USER ONLY. This does not delete posts or change them for other admins — it "
    "only updates the per-user viewpoint that controls whether posts appear in "
    "the user's archived view. Pass archived=True to archive, archived=False to "
    "unarchive."
)

_FAVORITE_DESCRIPTION = (
    "Favorite or unfavorite Microsoft 365 Message Center posts for the SIGNED-IN "
    "USER ONLY. This updates the per-user viewpoint that controls the user's "
    "favorites list in the admin center. Pass favorite=True to add to favorites, "
    "favorite=False to remove."
)


def _normalize_ids(message_ids: list[str]) -> list[str]:
    """Deduplicate ``message_ids`` while preserving first-seen order.

    Pydantic has already enforced per-element pattern and list length bounds;
    this helper just collapses repeats before the bulk POST so a buggy LLM
    asking to mark the same id 50 times doesn't waste the per-call budget.
    """

    seen: set[str] = set()
    ordered: list[str] = []
    for raw in message_ids:
        key = raw.strip().upper()
        if key in seen:
            continue
        seen.add(key)
        ordered.append(raw.strip())
    return ordered


def register(mcp: FastMCP, client: GraphClientProtocol) -> None:
    """Register the 3 viewpoint write tools on ``mcp``, bound to ``client``."""

    @mcp.tool(
        name="set_message_center_posts_read",
        title="Mark M365 Message Center posts read/unread (per-user)",
        description=_READ_DESCRIPTION,
    )
    async def set_message_center_posts_read(
        message_ids: Annotated[
            list[_MessageId],
            Field(
                min_length=1,
                max_length=MAX_MESSAGE_IDS_PER_WRITE,
                description=_MESSAGE_IDS_FIELD_DESCRIPTION,
            ),
        ],
        read: Annotated[
            bool,
            Field(
                description="True to mark as read, False to mark as unread.",
            ),
        ],
    ) -> dict[str, Any]:
        normalized = _normalize_ids(message_ids)
        await client.set_messages_read(normalized, read=read)
        return {
            "action": "markRead" if read else "markUnread",
            "state": read,
            "message_ids": normalized,
            "success": True,
        }

    @mcp.tool(
        name="set_message_center_posts_archived",
        title="Archive/unarchive M365 Message Center posts (per-user)",
        description=_ARCHIVED_DESCRIPTION,
    )
    async def set_message_center_posts_archived(
        message_ids: Annotated[
            list[_MessageId],
            Field(
                min_length=1,
                max_length=MAX_MESSAGE_IDS_PER_WRITE,
                description=_MESSAGE_IDS_FIELD_DESCRIPTION,
            ),
        ],
        archived: Annotated[
            bool,
            Field(
                description="True to archive, False to unarchive.",
            ),
        ],
    ) -> dict[str, Any]:
        normalized = _normalize_ids(message_ids)
        await client.set_messages_archive(normalized, archived=archived)
        return {
            "action": "archive" if archived else "unarchive",
            "state": archived,
            "message_ids": normalized,
            "success": True,
        }

    @mcp.tool(
        name="set_message_center_posts_favorite",
        title="Favorite/unfavorite M365 Message Center posts (per-user)",
        description=_FAVORITE_DESCRIPTION,
    )
    async def set_message_center_posts_favorite(
        message_ids: Annotated[
            list[_MessageId],
            Field(
                min_length=1,
                max_length=MAX_MESSAGE_IDS_PER_WRITE,
                description=_MESSAGE_IDS_FIELD_DESCRIPTION,
            ),
        ],
        favorite: Annotated[
            bool,
            Field(
                description="True to add to favorites, False to remove.",
            ),
        ],
    ) -> dict[str, Any]:
        normalized = _normalize_ids(message_ids)
        await client.set_messages_favorite(normalized, favorite=favorite)
        return {
            "action": "favorite" if favorite else "unfavorite",
            "state": favorite,
            "message_ids": normalized,
            "success": True,
        }


__all__ = ["MAX_MESSAGE_IDS_PER_WRITE", "register"]
