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


# ---------------------------------------------------------------------------
# Throttling (HTTP 429) handling
# ---------------------------------------------------------------------------


def test_parse_retry_after_http_date():
    """Retry-After may be an HTTP-date instead of a number of seconds."""
    from email.utils import format_datetime
    from datetime import timedelta

    from custom_components.microsoft_calendar.api import _parse_retry_after

    future = datetime.now(timezone.utc) + timedelta(seconds=90)
    seconds = _parse_retry_after(format_datetime(future, usegmt=True))
    assert 80 <= seconds <= 91


def test_parse_retry_after_garbage_uses_default():
    from custom_components.microsoft_calendar.api import _parse_retry_after

    assert _parse_retry_after("not-a-date") == 60
    assert _parse_retry_after(None) == 60


async def test_429_pauses_further_requests_without_network(client, mock_session):
    """After a 429, calls fail fast locally until Retry-After has passed."""
    mock_session.async_request.return_value = MockResponse(
        429, headers={"Retry-After": "30"}
    )
    with pytest.raises(MicrosoftGraphRateLimitError):
        await client.async_list_calendars()
    assert mock_session.async_request.call_count == 1

    with pytest.raises(MicrosoftGraphRateLimitError) as exc_info:
        await client.async_list_calendars()
    # No second network call while the back-off window is active.
    assert mock_session.async_request.call_count == 1
    assert 0 < exc_info.value.retry_after <= 30


async def test_requests_resume_after_backoff(client, mock_session, monkeypatch):
    from custom_components.microsoft_calendar import api as api_module

    clock = [1000.0]
    monkeypatch.setattr(api_module.time, "monotonic", lambda: clock[0])

    mock_session.async_request.return_value = MockResponse(
        429, headers={"Retry-After": "30"}
    )
    with pytest.raises(MicrosoftGraphRateLimitError):
        await client.async_list_calendars()

    clock[0] += 31
    mock_session.async_request.return_value = MockResponse(
        200, {"value": [{"id": "cal-1"}]}
    )
    result = await client.async_list_calendars()
    assert result[0]["id"] == "cal-1"


async def test_concurrent_requests_are_capped(client, mock_session):
    """No more than _MAX_CONCURRENT_REQUESTS Graph calls run at the same time."""
    import asyncio

    from custom_components.microsoft_calendar.api import _MAX_CONCURRENT_REQUESTS

    in_flight = 0
    peak = 0

    async def slow_request(*args, **kwargs):
        nonlocal in_flight, peak
        in_flight += 1
        peak = max(peak, in_flight)
        await asyncio.sleep(0.01)
        in_flight -= 1
        return MockResponse(200, {"value": []})

    mock_session.async_request.side_effect = slow_request
    await asyncio.gather(*(client.async_list_calendars() for _ in range(6)))
    assert peak == _MAX_CONCURRENT_REQUESTS
