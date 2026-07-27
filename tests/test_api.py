"""Tests for the Microsoft Graph API client (api.py)."""

from __future__ import annotations

from datetime import datetime, timezone
from unittest.mock import AsyncMock, call

import pytest

from custom_components.microsoft_calendar.api import (
    MicrosoftGraphAuthError,
    MicrosoftGraphClient,
    MicrosoftGraphError,
    MicrosoftGraphRateLimitError,
)
from custom_components.microsoft_calendar.const import GRAPH_BASE_URL

from .conftest import MockResponse


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def client(mock_session):
    return MicrosoftGraphClient(mock_session)


# ---------------------------------------------------------------------------
# async_list_calendars
# ---------------------------------------------------------------------------


async def test_list_calendars_single_page(client, mock_session):
    mock_session.async_request.return_value = MockResponse(
        200, {"value": [{"id": "cal-1", "name": "Work"}]}
    )
    result = await client.async_list_calendars()
    assert len(result) == 1
    assert result[0]["id"] == "cal-1"


async def test_list_calendars_pagination(client, mock_session):
    """All pages are followed and items are accumulated."""
    next_link = f"{GRAPH_BASE_URL}/me/calendars?$skiptoken=page2"
    mock_session.async_request.side_effect = [
        MockResponse(
            200,
            {
                "value": [{"id": "cal-1", "name": "Work"}],
                "@odata.nextLink": next_link,
            },
        ),
        MockResponse(200, {"value": [{"id": "cal-2", "name": "Personal"}]}),
    ]
    result = await client.async_list_calendars()
    assert len(result) == 2
    assert result[1]["id"] == "cal-2"


async def test_list_calendars_pagination_no_extra_params_on_next_page(
    client, mock_session
):
    """Original query params must NOT be re-sent on nextLink pages."""
    next_link = f"{GRAPH_BASE_URL}/me/calendars?$skiptoken=page2"
    mock_session.async_request.side_effect = [
        MockResponse(200, {"value": [{"id": "cal-1"}], "@odata.nextLink": next_link}),
        MockResponse(200, {"value": []}),
    ]
    await client.async_list_calendars()
    # Second call: only positional args — URL must be the nextLink, no kwargs
    second_call = mock_session.async_request.call_args_list[1]
    assert second_call == call("GET", next_link)


async def test_list_calendars_401_raises_auth_error(client, mock_session):
    mock_session.async_request.return_value = MockResponse(401)
    with pytest.raises(MicrosoftGraphAuthError):
        await client.async_list_calendars()


async def test_list_calendars_429_raises_rate_limit_error(client, mock_session):
    mock_session.async_request.return_value = MockResponse(
        429, headers={"Retry-After": "30"}
    )
    with pytest.raises(MicrosoftGraphRateLimitError) as exc_info:
        await client.async_list_calendars()
    assert exc_info.value.retry_after == 30


async def test_list_calendars_429_default_retry_after(client, mock_session):
    """Retry-After defaults to 60 when header is absent."""
    mock_session.async_request.return_value = MockResponse(429)
    with pytest.raises(MicrosoftGraphRateLimitError) as exc_info:
        await client.async_list_calendars()
    assert exc_info.value.retry_after == 60


async def test_list_calendars_500_raises_graph_error(client, mock_session):
    mock_session.async_request.return_value = MockResponse(
        500, text_data="Server error"
    )
    with pytest.raises(MicrosoftGraphError):
        await client.async_list_calendars()


# ---------------------------------------------------------------------------
# async_list_calendar_events
# ---------------------------------------------------------------------------


async def test_list_calendar_events_single_page(client, mock_session):
    mock_session.async_request.return_value = MockResponse(
        200, {"value": [{"id": "event-1", "subject": "Meeting"}]}
    )
    start = datetime(2024, 6, 1, tzinfo=timezone.utc)
    end = datetime(2024, 6, 2, tzinfo=timezone.utc)
    result = await client.async_list_calendar_events("cal-1", start, end)
    assert len(result) == 1
    assert result[0]["subject"] == "Meeting"


async def test_list_calendar_events_datetime_formatted_as_iso(client, mock_session):
    """startDateTime and endDateTime must be ISO 8601 strings with tz offset."""
    mock_session.async_request.return_value = MockResponse(200, {"value": []})
    start = datetime(2024, 6, 1, 0, 0, 0, tzinfo=timezone.utc)
    end = datetime(2024, 6, 2, 0, 0, 0, tzinfo=timezone.utc)
    await client.async_list_calendar_events("cal-1", start, end)

    _, url = mock_session.async_request.call_args.args
    params = mock_session.async_request.call_args.kwargs["params"]
    assert "calendarView" in url
    assert (
        "+00:00" in params["startDateTime"]
        or "Z" in params["startDateTime"]
        or "UTC" in params["startDateTime"]
    )


async def test_list_calendar_events_pagination(client, mock_session):
    next_link = f"{GRAPH_BASE_URL}/me/calendars/cal-1/calendarView?$skiptoken=p2"
    mock_session.async_request.side_effect = [
        MockResponse(
            200,
            {"value": [{"id": "e1"}], "@odata.nextLink": next_link},
        ),
        MockResponse(200, {"value": [{"id": "e2"}, {"id": "e3"}]}),
    ]
    start = datetime(2024, 6, 1, tzinfo=timezone.utc)
    end = datetime(2024, 6, 30, tzinfo=timezone.utc)
    result = await client.async_list_calendar_events("cal-1", start, end)
    assert len(result) == 3


async def test_list_calendar_events_401_raises_auth_error(client, mock_session):
    mock_session.async_request.return_value = MockResponse(401)
    start = datetime(2024, 6, 1, tzinfo=timezone.utc)
    end = datetime(2024, 6, 2, tzinfo=timezone.utc)
    with pytest.raises(MicrosoftGraphAuthError):
        await client.async_list_calendar_events("cal-1", start, end)


async def test_list_calendar_events_429_carries_retry_after(client, mock_session):
    mock_session.async_request.return_value = MockResponse(
        429, headers={"Retry-After": "120"}
    )
    start = datetime(2024, 6, 1, tzinfo=timezone.utc)
    end = datetime(2024, 6, 2, tzinfo=timezone.utc)
    with pytest.raises(MicrosoftGraphRateLimitError) as exc_info:
        await client.async_list_calendar_events("cal-1", start, end)
    assert exc_info.value.retry_after == 120
