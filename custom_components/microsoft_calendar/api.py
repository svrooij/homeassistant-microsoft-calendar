"""Microsoft Graph API client (no SDK — raw HTTPS via aiohttp)."""

from __future__ import annotations

import asyncio
import logging
import time
from datetime import datetime
from email.utils import parsedate_to_datetime
from typing import Any

from homeassistant.helpers.config_entry_oauth2_flow import OAuth2Session

from .const import GRAPH_BASE_URL

_LOGGER = logging.getLogger(__name__)

# Fields fetched for each calendar
_CALENDAR_SELECT = "id,name,hexColor,canEdit,isDefaultCalendar"

# Fields fetched for each event in a calendarView request
_EVENT_SELECT = (
    "id,iCalUId,subject,start,end,location,bodyPreview,isAllDay,sensitivity,type"
)

# Fallback back-off when Graph returns 429 without a usable Retry-After header.
_DEFAULT_RETRY_AFTER = 60

# Outlook allows at most 4 concurrent requests per app per mailbox. Stay under
# it: several calendars refreshing at once otherwise burst past the cap.
_MAX_CONCURRENT_REQUESTS = 2


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

    Also raised locally (without a network call) while a previous 429
    back-off window is still in effect.

    Attributes:
        retry_after: Seconds to wait before retrying (from the Retry-After header).
    """

    def __init__(self, message: str, retry_after: int = _DEFAULT_RETRY_AFTER) -> None:
        super().__init__(message)
        self.retry_after = retry_after


def _parse_retry_after(value: str | None) -> int:
    """Parse a Retry-After header (delta-seconds or HTTP-date) into seconds."""
    if not value:
        return _DEFAULT_RETRY_AFTER
    try:
        return max(1, int(value))
    except ValueError:
        pass
    try:
        retry_at = parsedate_to_datetime(value)
        return max(1, int((retry_at - datetime.now(retry_at.tzinfo)).total_seconds()))
    except (TypeError, ValueError):
        return _DEFAULT_RETRY_AFTER


# ---------------------------------------------------------------------------
# Client
# ---------------------------------------------------------------------------


class MicrosoftGraphClient:
    """Thin async wrapper around the Microsoft Graph REST API.

    Uses an HA OAuth2Session so that access-token refresh is handled
    automatically before every request.

    The client is shared by every calendar entity of a config entry, so the
    429 back-off window and the concurrency cap apply to the whole account,
    which matches how Graph applies its limits.
    """

    def __init__(self, oauth_session: OAuth2Session) -> None:
        self._session = oauth_session
        self._throttled_until: float = 0.0
        self._semaphore = asyncio.Semaphore(_MAX_CONCURRENT_REQUESTS)

    # ------------------------------------------------------------------
    # Throttle handling
    # ------------------------------------------------------------------

    @property
    def throttle_remaining(self) -> int:
        """Seconds left in the current 429 back-off window (0 if none)."""
        remaining = self._throttled_until - time.monotonic()
        return int(remaining) + 1 if remaining > 0 else 0

    def _check_throttle(self) -> None:
        """Raise without calling Graph while a back-off window is active."""
        if remaining := self.throttle_remaining:
            raise MicrosoftGraphRateLimitError(
                f"Throttled by Microsoft Graph; backing off for {remaining}s more",
                retry_after=remaining,
            )

    def _start_throttle(self, retry_after: int) -> None:
        """Record a 429 back-off window."""
        until = time.monotonic() + retry_after
        if until > self._throttled_until:
            self._throttled_until = until
            _LOGGER.warning(
                "Microsoft Graph throttled requests (HTTP 429); "
                "pausing requests for %s seconds",
                retry_after,
            )

    async def _send(self, method: str, url: str, **kwargs: Any) -> Any:
        """Send one request under the concurrency cap and throttle check."""
        async with self._semaphore:
            # Re-check after waiting: a 429 may have arrived meanwhile.
            self._check_throttle()
            resp = await self._session.async_request(method, url, **kwargs)
            if resp.status == 429:
                self._start_throttle(
                    _parse_retry_after(resp.headers.get("Retry-After"))
                )
            return resp

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
            MicrosoftGraphRateLimitError: HTTP 429, or a back-off is in effect.
            MicrosoftGraphError: Any other non-2xx response.
        """
        self._check_throttle()
        url = f"{GRAPH_BASE_URL}{path}"
        resp = await self._send(method, url, **kwargs)

        if resp.status == 401:
            raise MicrosoftGraphAuthError(f"Unauthorized (401) calling {method} {url}")
        if resp.status == 429:
            raise MicrosoftGraphRateLimitError(
                f"Rate limited (429) calling {method} {url}",
                retry_after=_parse_retry_after(resp.headers.get("Retry-After")),
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
            self._check_throttle()
            resp = await self._send("GET", next_url, **kwargs)

            if resp.status == 401:
                raise MicrosoftGraphAuthError(
                    f"Unauthorized (401) fetching paged results from {next_url}"
                )
            if resp.status == 429:
                raise MicrosoftGraphRateLimitError(
                    "Rate limited (429) fetching paged results",
                    retry_after=_parse_retry_after(resp.headers.get("Retry-After")),
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

        _LOGGER.debug("Loading events from calendar %s", calendar_id)
        return await self._get_all_pages(
            f"/me/calendars/{calendar_id}/calendarView",
            params={
                "startDateTime": start_str,
                "endDateTime": end_str,
                "top": 100,
                "$select": _EVENT_SELECT,
                "$orderby": "start/dateTime",
            },
        )

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
        return await self._request(
            "POST",
            f"/me/calendars/{calendar_id}/events",
            json=payload,
        )

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
