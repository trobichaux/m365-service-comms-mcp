"""Shared exception types and Graph error-envelope parsing."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class GraphError(RuntimeError):
    """A Microsoft Graph API call failed.

    Attributes:
        status_code: The HTTP status code returned by Graph.
        code: The ``error.code`` string from the Graph error envelope, or
            ``None`` if no envelope was present.
        message: The ``error.message`` string from the Graph error envelope,
            or the raw response text when no envelope was present.
        request_id: The ``request-id`` response header (if any), useful for
            correlating with Microsoft support cases.
    """

    status_code: int
    code: str | None
    message: str
    request_id: str | None = None

    def __str__(self) -> str:
        parts = [f"HTTP {self.status_code}"]
        if self.code:
            parts.append(self.code)
        parts.append(self.message)
        if self.request_id:
            parts.append(f"request-id={self.request_id}")
        return " | ".join(parts)


def parse_graph_error(
    *,
    status_code: int,
    body: object,
    raw_text: str = "",
    request_id: str | None = None,
) -> GraphError:
    """Build a :class:`GraphError` from a parsed JSON body or raw text fallback."""

    if isinstance(body, dict):
        envelope = body.get("error")
        if isinstance(envelope, dict):
            code = envelope.get("code") if isinstance(envelope.get("code"), str) else None
            message = envelope.get("message") if isinstance(envelope.get("message"), str) else ""
            return GraphError(
                status_code=status_code,
                code=code,
                message=message or "<no message>",
                request_id=request_id,
            )

    return GraphError(
        status_code=status_code,
        code=None,
        message=raw_text[:500] if raw_text else "<no body>",
        request_id=request_id,
    )
