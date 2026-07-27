"""Shared fixtures and helpers for Microsoft Calendar tests."""

from __future__ import annotations

import base64
import json
from typing import Any
from unittest.mock import AsyncMock, MagicMock

import pytest


# ---------------------------------------------------------------------------
# HTTP response mock
# ---------------------------------------------------------------------------


class MockResponse:
    """Minimal aiohttp response stand-in."""

    def __init__(
        self,
        status: int,
        json_data: dict | None = None,
        text_data: str = "",
        headers: dict | None = None,
    ) -> None:
        self.status = status
        self.ok = 200 <= status < 300
        self._json = json_data or {}
        self._text = text_data
        self.headers = headers or {}

    async def json(self) -> Any:
        return self._json

    async def text(self) -> str:
        return self._text


# ---------------------------------------------------------------------------
# OAuth2 session mock
# ---------------------------------------------------------------------------


@pytest.fixture
def mock_session() -> MagicMock:
    """Return a mock OAuth2Session whose async_request can be configured per test."""
    session = MagicMock()
    session.async_request = AsyncMock()
    return session


# ---------------------------------------------------------------------------
# JWT helper
# ---------------------------------------------------------------------------


def make_id_token(claims: dict[str, Any]) -> str:
    """Build a minimal unsigned JWT containing *claims* in the payload."""
    header = base64.urlsafe_b64encode(b'{"alg":"none"}').rstrip(b"=").decode()
    payload = (
        base64.urlsafe_b64encode(json.dumps(claims).encode()).rstrip(b"=").decode()
    )
    return f"{header}.{payload}.fakesig"


# ---------------------------------------------------------------------------
# Sample Graph API payloads
# ---------------------------------------------------------------------------

SAMPLE_CALENDAR = {
    "id": "cal-1",
    "name": "Work",
    "hexColor": "16a765",
    "canEdit": True,
    "isDefaultCalendar": True,
}

SAMPLE_CALENDAR_2 = {
    "id": "cal-2",
    "name": "Personal",
    "hexColor": "#ff0000",
    "canEdit": True,
    "isDefaultCalendar": False,
}

SAMPLE_TIMED_EVENT = {
    "id": "event-1",
    "iCalUId": "ical-uid-1",
    "subject": "Team meeting",
    "start": {"dateTime": "2024-06-01T09:00:00.0000000", "timeZone": "UTC"},
    "end": {"dateTime": "2024-06-01T10:00:00.0000000", "timeZone": "UTC"},
    "location": {"displayName": "Room 1"},
    "bodyPreview": "Weekly sync",
    "isAllDay": False,
    "sensitivity": "normal",
}

SAMPLE_ALLDAY_EVENT = {
    "id": "event-2",
    "iCalUId": "ical-uid-2",
    "subject": "Holiday",
    "start": {"date": "2024-06-15"},
    "end": {"date": "2024-06-16"},
    "location": None,
    "bodyPreview": "",
    "isAllDay": True,
    "sensitivity": "normal",
}
