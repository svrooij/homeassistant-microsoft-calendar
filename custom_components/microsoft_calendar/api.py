"""Microsoft Graph API client (no SDK — raw HTTPS via aiohttp)."""

from __future__ import annotations

import logging
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from math import ceil
from time import monotonic
from typing import Any

from homeassistant.helpers.config_entry_oauth2_flow import OAuth2Session

from .const import DEFAULT_EVENT_CACHE_DURATION, GRAPH_BASE_URL

_LOGGER = logging.getLogger(__name__)

# Fields fetched for each calendar
_CALENDAR_SELECT = "id,name,hexColor,canEdit,isDefaultCalendar"

# Fields fetched for each event in a calendarView request
_EVENT_SELECT = (
    "id,iCalUId,subject,start,end,location,bodyPreview,isAllDay,sensitivity,type"
)


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------


class MicrosoftGraphError(Exception):
    """Base exception for Microsoft Graph API errors."""


class MicrosoftGraphAuthError(MicrosoftGraphError):
    """Raised on HTTP 401 — the access token is invalid or expired.

    The coordinator catches this and triggers a re-auth flow.
    """


class MicrosoftGraphRateLimitError(MicrosoftGraphError):
    """Raised on HTTP 429 — the request was throttled by Microsoft Graph.

    Attributes:
        retry_after: Seconds to wait before retrying (from the Retry-After header).
    """

    def __init__(self, message: str, retry_after: int = 60) -> None:
        super().__init__(message)
        self.retry_after = retry_after


def _retry_after_seconds(headers: Any) -> int:
    """Return the delay requested by Graph's Retry-After header."""
    retry_after = headers.get("Retry-After", "60")
    try:
        return max(0, int(retry_after))
    except (TypeError, ValueError):
        try:
            retry_at = parsedate_to_datetime(retry_after)
            if retry_at.tzinfo is None:
                retry_at = retry_at.replace(tzinfo=timezone.utc)
            return max(0, ceil((retry_at - datetime.now(timezone.utc)).total_seconds()))
        except (TypeError, ValueError, OverflowError):
            return 60


def _as_utc_datetime(value: str | datetime) -> datetime:
    """Parse a Graph date/time value and normalize it to UTC."""
    if isinstance(value, str):
        value = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc)


def _event_interval(event: dict[str, Any]) -> tuple[datetime, datetime] | None:
    """Return an event's start and end as UTC datetimes."""
    try:
        start = event["start"]
        end = event["end"]
        return (
            _as_utc_datetime(start.get("date") or start["dateTime"]),
            _as_utc_datetime(end.get("date") or end["dateTime"]),
        )
    except (KeyError, TypeError, ValueError):
        return None


def _event_sort_key(event: dict[str, Any]) -> datetime:
    """Return an event's start time for ordering cached events."""
    interval = _event_interval(event)
    return interval[0] if interval else datetime.max.replace(tzinfo=timezone.utc)


# ---------------------------------------------------------------------------
# Client
# ---------------------------------------------------------------------------


class MicrosoftGraphClient:
    """Thin async wrapper around the Microsoft Graph REST API.

    Uses an HA OAuth2Session so that access-token refresh is handled
    automatically before every request.
    """

    def __init__(
        self,
        oauth_session: OAuth2Session,
        cache_duration: int = DEFAULT_EVENT_CACHE_DURATION,
    ) -> None:
        self._session = oauth_session
        self._cache_duration = cache_duration
        self._event_cache: dict[
            tuple[str, str, str], tuple[float, list[dict[str, Any]]]
        ] = {}
        self._rate_limited_until = 0.0

    def _check_rate_limit(self) -> None:
        """Prevent requests to Graph while the account-wide cooldown is active."""
        remaining = self._rate_limited_until - monotonic()
        if remaining > 0:
            raise MicrosoftGraphRateLimitError(
                "Rate limited (429); waiting for Microsoft Graph Retry-After",
                retry_after=ceil(remaining),
            )

    def _set_rate_limit(self, retry_after: int) -> None:
        """Block all requests from this client for the requested delay."""
        self._rate_limited_until = max(
            self._rate_limited_until, monotonic() + retry_after
        )

    # ------------------------------------------------------------------
    # Low-level helpers
    # ------------------------------------------------------------------

    async def _request(
        self,
        method: str,
        path: str,
        **kwargs: Any,
    ) -> Any:
        """Make a single authenticated request and return the parsed JSON body.

        Args:
            method: HTTP method (e.g. "GET").
            path: Path relative to GRAPH_BASE_URL (e.g. "/me/calendars").
            **kwargs: Forwarded to the underlying aiohttp request (params, json, …).

        Raises:
            MicrosoftGraphAuthError: HTTP 401.
            MicrosoftGraphRateLimitError: HTTP 429.
            MicrosoftGraphError: Any other non-2xx response.
        """
        self._check_rate_limit()
        url = f"{GRAPH_BASE_URL}{path}"
        resp = await self._session.async_request(method, url, **kwargs)

        if resp.status == 401:
            raise MicrosoftGraphAuthError(f"Unauthorized (401) calling {method} {url}")
        if resp.status == 429:
            retry_after = _retry_after_seconds(resp.headers)
            self._set_rate_limit(retry_after)
            raise MicrosoftGraphRateLimitError(
                f"Rate limited (429) calling {method} {url}",
                retry_after=retry_after,
            )
        if not resp.ok:
            body = await resp.text()
            raise MicrosoftGraphError(
                f"HTTP {resp.status} calling {method} {url}: {body[:200]}"
            )

        return await resp.json() if resp.content_length else None

    async def _get_all_pages(self, path: str, **kwargs: Any) -> list[dict[str, Any]]:
        """Collect all items from a paged Graph response.

        Follows ``@odata.nextLink`` until the final page is reached.
        """
        items: list[dict[str, Any]] = []
        next_url: str | None = f"{GRAPH_BASE_URL}{path}"

        while next_url:
            self._check_rate_limit()
            resp = await self._session.async_request("GET", next_url, **kwargs)

            if resp.status == 401:
                raise MicrosoftGraphAuthError(
                    f"Unauthorized (401) fetching paged results from {next_url}"
                )
            if resp.status == 429:
                retry_after = _retry_after_seconds(resp.headers)
                self._set_rate_limit(retry_after)
                raise MicrosoftGraphRateLimitError(
                    f"Rate limited (429) fetching paged results",
                    retry_after=retry_after,
                )
            if not resp.ok:
                body = await resp.text()
                raise MicrosoftGraphError(
                    f"HTTP {resp.status} fetching paged results: {body[:200]}"
                )

            data: dict[str, Any] = await resp.json()
            items.extend(data.get("value", []))

            # nextLink already contains the full URL including query params
            next_url = data.get("@odata.nextLink")
            # kwargs must not be forwarded on subsequent pages — the nextLink
            # already encodes all the required query parameters.
            kwargs = {}

        return items

    # ------------------------------------------------------------------
    # Calendars
    # ------------------------------------------------------------------

    async def async_list_calendars(self) -> list[dict[str, Any]]:
        """Return all calendars for the signed-in user.

        Each dict contains at least: id, name, hexColor, canEdit, isDefaultCalendar.
        """
        return await self._get_all_pages(
            "/me/calendars",
            params={"$select": _CALENDAR_SELECT, "top": 100},
        )

    # ------------------------------------------------------------------
    # Events (calendar view)
    # ------------------------------------------------------------------

    async def async_list_calendar_events(
        self,
        calendar_id: str,
        start_dt: datetime,
        end_dt: datetime,
    ) -> list[dict[str, Any]]:
        """Return all event instances in *calendar_id* within [start_dt, end_dt).

        Uses the ``calendarView`` endpoint which expands recurring event
        instances automatically — no client-side recurrence expansion needed.

        Args:
            calendar_id: The Graph calendar id.
            start_dt: Lower bound (inclusive), timezone-aware.
            end_dt: Upper bound (exclusive), timezone-aware.

        Returns:
            Flat list of event dicts, each containing at minimum the fields
            defined in ``_EVENT_SELECT``.
        """
        # Graph requires ISO 8601 with UTC offset, e.g. "2024-06-01T00:00:00+00:00"
        start_str = start_dt.isoformat()
        end_str = end_dt.isoformat()
        cache_key = (
            calendar_id,
            _as_utc_datetime(start_dt).isoformat(),
            _as_utc_datetime(end_dt).isoformat(),
        )
        cached = self._event_cache.get(cache_key)
        if cached and cached[0] > monotonic():
            return cached[1]

        _LOGGER.debug("Loading events from calendar %s", calendar_id)
        try:
            events = await self._get_all_pages(
                f"/me/calendars/{calendar_id}/calendarView",
                params={
                    "startDateTime": start_str,
                    "endDateTime": end_str,
                    "top": 100,
                    "$select": _EVENT_SELECT,
                    "$orderby": "start/dateTime",
                },
            )
        except MicrosoftGraphRateLimitError:
            if cached:
                return cached[1]
            raise

        if self._cache_duration:
            self._event_cache[cache_key] = (
                monotonic() + self._cache_duration,
                events,
            )
        return events

    async def async_get_event(self, event_id: str) -> dict[str, Any]:
        """Fetch a single event by its Graph event ID."""
        return await self._request(
            "GET",
            f"/me/events/{event_id}",
            params={"$select": "id,type"},
        )

    async def async_create_event(
        self, calendar_id: str, payload: dict[str, Any]
    ) -> dict[str, Any]:
        """Create a new event in the given calendar."""
        event = await self._request(
            "POST",
            f"/me/calendars/{calendar_id}/events",
            json=payload,
        )
        if self._cache_duration and isinstance(event, dict):
            interval = _event_interval(event)
            if interval:
                event_start, event_end = interval
                for (cached_calendar, start, end), (_expires_at, events) in (
                    self._event_cache.items()
                ):
                    if cached_calendar != calendar_id:
                        continue
                    range_start = _as_utc_datetime(start)
                    range_end = _as_utc_datetime(end)
                    if event_start >= range_end or event_end <= range_start:
                        continue
                    if not any(item.get("id") == event.get("id") for item in events):
                        events.append(event)
                        events.sort(key=_event_sort_key)
        return event

    async def async_update_event(
        self, event_id: str, payload: dict[str, Any]
    ) -> dict[str, Any]:
        """Update an existing event (PATCH — only supplied fields are changed)."""
        return await self._request(
            "PATCH",
            f"/me/events/{event_id}",
            json=payload,
        )

    async def async_delete_event(self, event_id: str) -> None:
        """Permanently delete an event."""
        await self._request("DELETE", f"/me/events/{event_id}")
