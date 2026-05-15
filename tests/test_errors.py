"""Tests for :mod:`m365_service_comms_mcp.errors`."""

from __future__ import annotations

from m365_service_comms_mcp.errors import GraphError, parse_graph_error


def test_parse_graph_error_with_full_envelope() -> None:
    err = parse_graph_error(
        status_code=403,
        body={
            "error": {
                "code": "Authorization_RequestDenied",
                "message": "Insufficient privileges.",
            }
        },
        request_id="req-1",
    )
    assert err.status_code == 403
    assert err.code == "Authorization_RequestDenied"
    assert err.message == "Insufficient privileges."
    assert err.request_id == "req-1"


def test_parse_graph_error_falls_back_to_raw_text() -> None:
    err = parse_graph_error(
        status_code=500,
        body=None,
        raw_text="upstream gateway error",
        request_id=None,
    )
    assert err.status_code == 500
    assert err.code is None
    assert err.message == "upstream gateway error"


def test_parse_graph_error_truncates_long_text() -> None:
    err = parse_graph_error(status_code=502, body=None, raw_text="x" * 10000)
    assert len(err.message) == 500


def test_parse_graph_error_handles_partial_envelope() -> None:
    err = parse_graph_error(
        status_code=400,
        body={"error": {"code": "BadRequest"}},
    )
    assert err.code == "BadRequest"
    assert err.message == "<no message>"


def test_str_includes_status_and_optional_request_id() -> None:
    err = GraphError(status_code=429, code="Throttled", message="slow", request_id="r-9")
    rendered = str(err)
    assert "HTTP 429" in rendered
    assert "Throttled" in rendered
    assert "slow" in rendered
    assert "request-id=r-9" in rendered


def test_str_omits_request_id_when_absent() -> None:
    err = GraphError(status_code=400, code=None, message="bad")
    rendered = str(err)
    assert "HTTP 400" in rendered
    assert "bad" in rendered
    assert "request-id" not in rendered
