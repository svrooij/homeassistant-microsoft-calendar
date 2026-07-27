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
    assert ev.uid == "ical-1"
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
