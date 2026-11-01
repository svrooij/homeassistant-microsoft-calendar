"""Tests for calendar helper functions and the CalendarEntity (calendar.py)."""

from __future__ import annotations

from datetime import date, datetime, timezone
from typing import Any
from unittest.mock import AsyncMock, MagicMock, patch
from zoneinfo import ZoneInfo

import pytest

from custom_components.microsoft_calendar.calendar import (
    _graph_event_to_calendar_event,
    _parse_graph_datetime,
)


# ---------------------------------------------------------------------------
# _parse_graph_datetime
# ---------------------------------------------------------------------------


def test_parse_timed_event_utc():
    result = _parse_graph_datetime(
        {"dateTime": "2024-06-01T09:00:00.0000000", "timeZone": "UTC"}
    )
    assert isinstance(result, datetime)
    assert result == datetime(2024, 6, 1, 9, 0, 0, tzinfo=timezone.utc)


def test_parse_timed_event_named_timezone():
    result = _parse_graph_datetime(
        {"dateTime": "2024-06-01T09:00:00", "timeZone": "Europe/Amsterdam"}
    )
    assert isinstance(result, datetime)
    assert result.tzinfo == ZoneInfo("Europe/Amsterdam")
    assert result.hour == 9


def test_parse_timed_event_unknown_tz_falls_back_to_utc():
    result = _parse_graph_datetime(
        {"dateTime": "2024-06-01T09:00:00", "timeZone": "tzone://Microsoft/Utc"}
    )
    assert isinstance(result, datetime)
    assert result.tzinfo == timezone.utc


def test_parse_timed_event_strips_subsecond_precision():
    # Should not raise ValueError even with 7-digit fractional seconds
    result = _parse_graph_datetime(
        {"dateTime": "2024-06-01T09:00:00.1234567", "timeZone": "UTC"}
    )
    assert isinstance(result, datetime)
    assert result.microsecond == 0


def test_parse_allday_event_date_field():
    result = _parse_graph_datetime({"date": "2024-06-15"})
    assert isinstance(result, date)
    assert not isinstance(result, datetime)
    assert result == date(2024, 6, 15)


def test_parse_allday_event_date_takes_priority_over_datetime():
    # When both keys exist, "date" wins
    result = _parse_graph_datetime(
        {"date": "2024-06-15", "dateTime": "2024-06-15T00:00:00", "timeZone": "UTC"}
    )
    assert isinstance(result, date)
    assert not isinstance(result, datetime)


# ---------------------------------------------------------------------------
# _graph_event_to_calendar_event
# ---------------------------------------------------------------------------


def test_timed_event_mapping():
    raw: dict[str, Any] = {
        "id": "event-1",
        "iCalUId": "ical-1",
        "subject": "Team meeting",
        "start": {"dateTime": "2024-06-01T09:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2024-06-01T10:00:00", "timeZone": "UTC"},
        "location": {"displayName": "Room 1"},
        "bodyPreview": "Weekly sync",
        "isAllDay": False,
    }
    ev = _graph_event_to_calendar_event(raw)
    assert ev is not None
    assert ev.summary == "Team meeting"
    assert ev.uid == "event-1"
    assert isinstance(ev.start, datetime)
    assert isinstance(ev.end, datetime)
    assert ev.location == "Room 1"
    assert ev.description == "Weekly sync"


def test_allday_event_mapping_uses_date():
    raw: dict[str, Any] = {
        "id": "event-2",
        "iCalUId": "ical-2",
        "subject": "Holiday",
        "start": {"date": "2024-06-15"},
        "end": {"date": "2024-06-16"},
        "location": None,
        "bodyPreview": "",
        "isAllDay": True,
    }
    ev = _graph_event_to_calendar_event(raw)
    assert ev is not None
    assert isinstance(ev.start, date)
    assert not isinstance(ev.start, datetime)
    assert ev.start == date(2024, 6, 15)
    assert ev.end == date(2024, 6, 16)


def test_allday_event_with_datetime_field_still_uses_date():
    """isAllDay=True must produce date objects even if only dateTime is present."""
    raw: dict[str, Any] = {
        "id": "event-3",
        "subject": "Conference",
        "start": {"dateTime": "2024-06-15T00:00:00.0000000", "timeZone": "UTC"},
        "end": {"dateTime": "2024-06-16T00:00:00.0000000", "timeZone": "UTC"},
        "isAllDay": True,
    }
    ev = _graph_event_to_calendar_event(raw)
    assert ev is not None
    assert isinstance(ev.start, date)
    assert not isinstance(ev.start, datetime)
    assert ev.start == date(2024, 6, 15)


def test_no_subject_uses_placeholder():
    raw: dict[str, Any] = {
        "id": "event-4",
        "subject": None,
        "start": {"dateTime": "2024-06-01T09:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2024-06-01T10:00:00", "timeZone": "UTC"},
        "isAllDay": False,
    }
    ev = _graph_event_to_calendar_event(raw)
    assert ev is not None
    assert ev.summary == "(no title)"


def test_empty_location_display_name_gives_none():
    raw: dict[str, Any] = {
        "id": "event-5",
        "subject": "Meeting",
        "start": {"dateTime": "2024-06-01T09:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2024-06-01T10:00:00", "timeZone": "UTC"},
        "location": {"displayName": ""},
        "isAllDay": False,
    }
    ev = _graph_event_to_calendar_event(raw)
    assert ev is not None
    assert ev.location is None


def test_missing_start_returns_none():
    raw: dict[str, Any] = {
        "id": "event-6",
        "subject": "Broken",
        "end": {"dateTime": "2024-06-01T10:00:00", "timeZone": "UTC"},
        "isAllDay": False,
    }
    ev = _graph_event_to_calendar_event(raw)
    assert ev is None


def test_icaluid_falls_back_to_id():
    raw: dict[str, Any] = {
        "id": "event-7",
        # no iCalUId
        "subject": "Meeting",
        "start": {"dateTime": "2024-06-01T09:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2024-06-01T10:00:00", "timeZone": "UTC"},
        "isAllDay": False,
    }
    ev = _graph_event_to_calendar_event(raw)
    assert ev is not None
    assert ev.uid == "event-7"


def test_hex_color_without_hash_is_prefixed():
    """Calendar entity should prefix '#' when Microsoft omits it."""
    from unittest.mock import MagicMock
    from custom_components.microsoft_calendar.calendar import MicrosoftCalendarEntity

    coordinator = MagicMock()
    coordinator.data = []
    client = MagicMock()
    entry = MagicMock()
    entry.entry_id = "entry-1"

    cal_data = {"id": "cal-1", "name": "Work", "hexColor": "16a765"}
    entity = MicrosoftCalendarEntity(coordinator, client, entry, cal_data)
    assert entity._attr_initial_color == "#16a765"


def test_hex_color_with_hash_is_unchanged():
    from unittest.mock import MagicMock
    from custom_components.microsoft_calendar.calendar import MicrosoftCalendarEntity

    coordinator = MagicMock()
    coordinator.data = []
    client = MagicMock()
    entry = MagicMock()
    entry.entry_id = "entry-1"

    cal_data = {"id": "cal-1", "name": "Work", "hexColor": "#16a765"}
    entity = MicrosoftCalendarEntity(coordinator, client, entry, cal_data)
    assert entity._attr_initial_color == "#16a765"


def test_empty_hex_color_gives_none():
    from unittest.mock import MagicMock
    from custom_components.microsoft_calendar.calendar import MicrosoftCalendarEntity

    coordinator = MagicMock()
    coordinator.data = []
    client = MagicMock()
    entry = MagicMock()
    entry.entry_id = "entry-1"

    cal_data = {"id": "cal-1", "name": "Work", "hexColor": ""}
    entity = MicrosoftCalendarEntity(coordinator, client, entry, cal_data)
    assert entity._attr_initial_color is None


# ---------------------------------------------------------------------------
# Event cache and throttle fallback
# ---------------------------------------------------------------------------


def _make_entity(client):
    from custom_components.microsoft_calendar.calendar import MicrosoftCalendarEntity

    coordinator = MagicMock()
    coordinator.data = []
    entry = MagicMock()
    entry.entry_id = "entry-1"
    entry.data = {}
    return MicrosoftCalendarEntity(
        coordinator, client, entry, {"id": "cal-1", "name": "Work"}
    )


_RAW_EVENT = {
    "id": "event-1",
    "subject": "Standup",
    "start": {"dateTime": "2024-06-10T09:00:00.0000000", "timeZone": "UTC"},
    "end": {"dateTime": "2024-06-10T09:15:00.0000000", "timeZone": "UTC"},
    "isAllDay": False,
}

_MONTH_START = datetime(2024, 6, 1, tzinfo=timezone.utc)
_MONTH_END = datetime(2024, 7, 1, tzinfo=timezone.utc)


async def test_get_events_is_cached():
    """Repeated requests for the same range call Graph only once."""
    client = MagicMock()
    client.async_list_calendar_events = AsyncMock(return_value=[_RAW_EVENT])
    entity = _make_entity(client)

    first = await entity.async_get_events(None, _MONTH_START, _MONTH_END)
    second = await entity.async_get_events(None, _MONTH_START, _MONTH_END)

    assert [ev.summary for ev in first] == ["Standup"]
    assert second == first
    assert client.async_list_calendar_events.await_count == 1


async def test_sub_range_served_from_cached_wider_range():
    """A narrower range inside a fresh cached range needs no Graph call."""
    client = MagicMock()
    client.async_list_calendar_events = AsyncMock(return_value=[_RAW_EVENT])
    entity = _make_entity(client)

    await entity.async_get_events(None, _MONTH_START, _MONTH_END)
    day = await entity.async_get_events(
        None,
        datetime(2024, 6, 10, tzinfo=timezone.utc),
        datetime(2024, 6, 11, tzinfo=timezone.utc),
    )
    assert [ev.summary for ev in day] == ["Standup"]
    assert client.async_list_calendar_events.await_count == 1


async def test_throttled_returns_last_known_events(monkeypatch):
    """When Graph throttles, stale cached events are returned instead of []."""
    from custom_components.microsoft_calendar import calendar as calendar_module
    from custom_components.microsoft_calendar.api import MicrosoftGraphRateLimitError

    clock = [1000.0]
    monkeypatch.setattr(calendar_module.time, "monotonic", lambda: clock[0])

    client = MagicMock()
    client.async_list_calendar_events = AsyncMock(return_value=[_RAW_EVENT])
    entity = _make_entity(client)
    await entity.async_get_events(None, _MONTH_START, _MONTH_END)

    clock[0] += calendar_module._EVENT_CACHE_TTL_SECONDS + 1
    client.async_list_calendar_events.side_effect = MicrosoftGraphRateLimitError(
        "throttled", retry_after=30
    )
    events = await entity.async_get_events(None, _MONTH_START, _MONTH_END)
    assert [ev.summary for ev in events] == ["Standup"]
    assert client.async_list_calendar_events.await_count == 2


async def test_throttled_without_cache_returns_empty():
    from custom_components.microsoft_calendar.api import MicrosoftGraphRateLimitError

    client = MagicMock()
    client.async_list_calendar_events = AsyncMock(
        side_effect=MicrosoftGraphRateLimitError("throttled", retry_after=30)
    )
    entity = _make_entity(client)
    assert await entity.async_get_events(None, _MONTH_START, _MONTH_END) == []
