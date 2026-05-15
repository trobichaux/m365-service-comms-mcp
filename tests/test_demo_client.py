"""Tests for :mod:`m365_service_comms_mcp.demo_client`."""

from __future__ import annotations

import pytest

from m365_service_comms_mcp.demo_client import DemoGraphClient
from m365_service_comms_mcp.errors import GraphError


async def test_health_overviews_returns_value_array() -> None:
    client = DemoGraphClient()
    result = await client.list_health_overviews()
    assert "value" in result
    assert isinstance(result["value"], list)
    assert len(result["value"]) >= 3
    for row in result["value"]:
        assert {"id", "service", "status"}.issubset(row.keys())


async def test_health_overviews_respects_top() -> None:
    client = DemoGraphClient()
    result = await client.list_health_overviews(top=2)
    assert len(result["value"]) == 2


async def test_messages_returns_message_records() -> None:
    client = DemoGraphClient()
    result = await client.list_messages()
    assert "value" in result
    assert len(result["value"]) >= 1
    for row in result["value"]:
        assert {"id", "title", "category", "severity", "services"}.issubset(row.keys())


async def test_get_message_returns_full_body() -> None:
    client = DemoGraphClient()
    result = await client.get_message("MC987654")
    assert result["id"] == "MC987654"
    assert "body" in result
    assert "content" in result["body"]
    assert "Outlook" in result["body"]["content"]


async def test_get_message_is_case_insensitive() -> None:
    client = DemoGraphClient()
    result = await client.get_message("mc987654")
    assert result["id"] == "MC987654"


async def test_get_message_unknown_id_raises_graph_error() -> None:
    client = DemoGraphClient()
    with pytest.raises(GraphError) as info:
        await client.get_message("MC999999")
    err = info.value
    assert err.status_code == 404
    assert err.code == "ResourceNotFound"
    assert "MC999999" in err.message


async def test_aclose_is_a_noop() -> None:
    client = DemoGraphClient()
    await client.aclose()
    # Still usable after aclose for the demo client.
    result = await client.list_health_overviews(top=1)
    assert len(result["value"]) == 1
