"""Calendar platform for Microsoft Calendar."""

from __future__ import annotations

import asyncio
import logging
import time
from datetime import date, datetime, timedelta, timezone
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from homeassistant.components.calendar import (
    CalendarEntity,
    CalendarEntityFeature,
    CalendarEvent,
    EVENT_DESCRIPTION,
    EVENT_END,
    EVENT_LOCATION,
    EVENT_START,
    EVENT_SUMMARY,
)
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant, callback
from homeassistant.exceptions import HomeAssistantError
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity
from homeassistant.util import dt as dt_util

from .api import (
    MicrosoftGraphClient,
    MicrosoftGraphError,
    MicrosoftGraphRateLimitError,
)
from .const import (
    CALENDAR_SCOPE_WRITE,
    CONF_CALENDAR_SCOPE,
    DATA_CLIENT,
    DATA_COORDINATOR,
    DOMAIN,
)
from .coordinator import MicrosoftCalendarCoordinator

_LOGGER = logging.getLogger(__name__)

# How far ahead to look when determining the current/next event for the state.
_NEXT_EVENT_LOOKAHEAD = timedelta(days=1)

# How long fetched events are reused before Graph is asked again. Calendar
# cards on several open dashboards request the same range repeatedly; without
# this every redraw is a Graph call and the account gets throttled (HTTP 429).
_EVENT_CACHE_TTL_SECONDS = 300
# Maximum number of distinct date ranges cached per calendar.
_EVENT_CACHE_MAX_ENTRIES = 16


# ---------------------------------------------------------------------------
# Platform setup
# ---------------------------------------------------------------------------


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up Microsoft Calendar entities from a config entry."""
    data = hass.data[DOMAIN][entry.entry_id]
    coordinator: MicrosoftCalendarCoordinator = data[DATA_COORDINATOR]
    client: MicrosoftGraphClient = data[DATA_CLIENT]

    known_calendar_ids: set[str] = set()

    @callback
    def _add_new_calendars() -> None:
        """Add entities for calendars that appeared since the last refresh."""
        if coordinator.data is None:
            return
        new_entities = []
        for cal in coordinator.data:
            cal_id = cal["id"]
            if cal_id not in known_calendar_ids:
                known_calendar_ids.add(cal_id)
                new_entities.append(
                    MicrosoftCalendarEntity(coordinator, client, entry, cal)
                )
        if new_entities:
            async_add_entities(new_entities)

    # Register the callback so it runs on every future coordinator refresh.
    entry.async_on_unload(coordinator.async_add_listener(_add_new_calendars))

    # Also run immediately for the data already fetched during first_refresh.
    _add_new_calendars()


# ---------------------------------------------------------------------------
# Date/time helpers
# ---------------------------------------------------------------------------


def _parse_graph_datetime(
    dt_obj: dict[str, str],
) -> datetime | date:
    """Convert a Microsoft Graph dateTime/date object to a Python value.

    Graph returns one of two shapes:
    - Timed event:  {"dateTime": "2024-06-01T09:00:00.0000000", "timeZone": "UTC"}
    - All-day event: {"dateTime": "2024-06-01T00:00:00.0000000", "timeZone": "tzone://Microsoft/Utc"}
      *or* the "date" key is used by some Graph versions/configurations.

    We rely on the ``isAllDay`` flag on the parent event (passed as a hint)
    rather than trying to infer it from the time value.
    """
    raw_dt: str | None = dt_obj.get("dateTime")
    raw_date: str | None = dt_obj.get("date")
    tz_name: str | None = dt_obj.get("timeZone", "UTC")

    if raw_date:
        # Pure date — all-day event, no time component.
        return date.fromisoformat(raw_date)

    if raw_dt:
        # Strip sub-second precision that Python's fromisoformat can't handle
        # in all versions (e.g. "2024-06-01T09:00:00.0000000").
        raw_dt = raw_dt.split(".")[0]
        naive_dt = datetime.fromisoformat(raw_dt)

        try:
            tz = ZoneInfo(tz_name) if tz_name else timezone.utc
        except (ZoneInfoNotFoundError, ValueError):
            _LOGGER.debug(
                "Unknown timezone %r from Microsoft Graph, falling back to UTC",
                tz_name,
            )
            tz = timezone.utc

        return naive_dt.replace(tzinfo=tz)

    # Defensive fallback — should not occur with valid Graph responses.
    return date.today()


def _graph_event_to_calendar_event(raw: dict[str, Any]) -> CalendarEvent | None:
    """Map a Microsoft Graph event dict to a HA CalendarEvent.

    Returns None if mandatory fields are missing or unparseable.
    """
    try:
        is_all_day: bool = raw.get("isAllDay", False)

        start_obj: dict[str, str] = raw["start"]
        end_obj: dict[str, str] = raw["end"]

        # For all-day events, always use date (not datetime).
        if is_all_day:
            start_val: date | datetime = date.fromisoformat(
                start_obj.get("date") or start_obj["dateTime"].split("T")[0]
            )
            end_val: date | datetime = date.fromisoformat(
                end_obj.get("date") or end_obj["dateTime"].split("T")[0]
            )
        else:
            start_val = _parse_graph_datetime(start_obj)
            end_val = _parse_graph_datetime(end_obj)

        location_obj: dict[str, str] | None = raw.get("location")
        location: str | None = (
            location_obj.get("displayName") or None if location_obj else None
        )

        return CalendarEvent(
            uid=raw.get("id"),
            summary=raw.get("subject") or "(no title)",
            start=start_val,
            end=end_val,
            location=location,
            description=raw.get("bodyPreview") or None,
        )
    except (KeyError, ValueError) as err:
        _LOGGER.debug("Could not parse Graph event %s: %s", raw.get("id"), err)
        return None


# ---------------------------------------------------------------------------
# Graph payload helpers
# ---------------------------------------------------------------------------

_RECURRING_TYPES = frozenset({"occurrence", "exception", "seriesMaster"})


def _dt_to_graph(value: datetime | date) -> dict[str, str]:
    """Convert a Python datetime/date to a Microsoft Graph dateTime object."""
    if isinstance(value, datetime):
        utc = value.astimezone(timezone.utc)
        return {"dateTime": utc.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"}
    # All-day: date only, no time component
    return {"dateTime": f"{value.isoformat()}T00:00:00", "timeZone": "UTC"}


def _event_kwargs_to_graph(kwargs: dict[str, Any]) -> dict[str, Any]:
    """Convert HA calendar event kwargs/dict to a Microsoft Graph event payload."""
    start: datetime | date = kwargs[EVENT_START]
    end: datetime | date = kwargs[EVENT_END]
    is_all_day = isinstance(start, date) and not isinstance(start, datetime)

    payload: dict[str, Any] = {
        "subject": kwargs[EVENT_SUMMARY],
        "isAllDay": is_all_day,
        "start": _dt_to_graph(start),
        "end": _dt_to_graph(end),
    }
    if description := kwargs.get(EVENT_DESCRIPTION):
        payload["body"] = {"contentType": "text", "content": description}
    if location := kwargs.get(EVENT_LOCATION):
        payload["location"] = {"displayName": location}
    return payload


# ---------------------------------------------------------------------------
# Entity
# ---------------------------------------------------------------------------


class MicrosoftCalendarEntity(
    CoordinatorEntity[MicrosoftCalendarCoordinator], CalendarEntity
):
    """Represents a single Microsoft Calendar as a HA calendar entity."""

    _attr_has_entity_name = True

    def __init__(
        self,
        coordinator: MicrosoftCalendarCoordinator,
        client: MicrosoftGraphClient,
        entry: ConfigEntry,
        calendar_data: dict[str, Any],
    ) -> None:
        super().__init__(coordinator)
        self._client = client
        self._entry = entry
        self._calendar_id: str = calendar_data["id"]
        self._attr_unique_id = f"{entry.entry_id}_{self._calendar_id}"
        self._attr_name: str = calendar_data.get("name", "Microsoft Calendar")
        self._can_edit: bool = bool(calendar_data.get("canEdit", False))

        # Prefix the hex colour with '#' if Microsoft omitted it.
        raw_color: str = calendar_data.get("hexColor") or ""
        if raw_color and not raw_color.startswith("#"):
            raw_color = f"#{raw_color}"
        self._attr_initial_color: str | None = raw_color or None

        # Enable write features only when the user chose ReadWrite scope
        # and Microsoft reports the calendar as editable.
        if (
            self._can_edit
            and entry.data.get(CONF_CALENDAR_SCOPE) == CALENDAR_SCOPE_WRITE
        ):
            self._attr_supported_features = (
                CalendarEntityFeature.CREATE_EVENT
                | CalendarEntityFeature.DELETE_EVENT
                | CalendarEntityFeature.UPDATE_EVENT
            )

        # Current / next upcoming event, refreshed on each coordinator update.
        self._event: CalendarEvent | None = None

        # Short-lived cache of fetched events, keyed by requested range.
        # Value: (fetched_at monotonic seconds, range start, range end, events)
        self._event_cache: dict[
            tuple[str, str], tuple[float, datetime, datetime, list[CalendarEvent]]
        ] = {}
        # Serialises fetches so concurrent identical requests share one call.
        self._fetch_lock = asyncio.Lock()

    # ------------------------------------------------------------------
    # State
    # ------------------------------------------------------------------

    @property
    def event(self) -> CalendarEvent | None:
        """Return the current or next upcoming event."""
        return self._event

    @callback
    def _handle_coordinator_update(self) -> None:
        """Schedule a background refresh of the upcoming event on each poll."""
        # Trigger an async task — we can't await in a callback.
        self.hass.async_create_task(self._async_refresh_next_event())

    # ------------------------------------------------------------------
    # Event fetching with cache and throttle fallback
    # ------------------------------------------------------------------

    def _cached_events_covering(
        self, start: datetime, end: datetime, max_age: float | None = None
    ) -> list[CalendarEvent] | None:
        """Return events for [start, end) from the newest cached wider range.

        With max_age, only cache entries younger than that are considered;
        without it, any age is accepted (used as a fallback while throttled,
        so the calendar keeps showing the last known events).
        """
        now = time.monotonic()
        best: tuple[float, list[CalendarEvent]] | None = None
        for fetched_at, c_start, c_end, events in self._event_cache.values():
            if max_age is not None and now - fetched_at >= max_age:
                continue
            if c_start <= start and c_end >= end and (
                best is None or fetched_at > best[0]
            ):
                best = (fetched_at, events)
        if best is None:
            return None
        return [
            ev
            for ev in best[1]
            if ev.end_datetime_local > start and ev.start_datetime_local < end
        ]

    async def _async_fetch_events(
        self, start: datetime, end: datetime
    ) -> list[CalendarEvent]:
        """Fetch events for a range, reusing fresh cached data when present.

        Raises MicrosoftGraphError if Graph fails and no cached data can
        stand in for the request.
        """
        key = (start.isoformat(), end.isoformat())
        async with self._fetch_lock:
            now = time.monotonic()
            cached = self._event_cache.get(key)
            if cached and now - cached[0] < _EVENT_CACHE_TTL_SECONDS:
                return cached[3]
            # A fresh, wider range already fetched (e.g. the month view)
            # can answer this request without calling Graph.
            if (
                fresh := self._cached_events_covering(
                    start, end, _EVENT_CACHE_TTL_SECONDS
                )
            ) is not None:
                return fresh

            try:
                raw_events = await self._client.async_list_calendar_events(
                    self._calendar_id, start, end
                )
            except MicrosoftGraphRateLimitError:
                if cached:
                    return cached[3]
                if (fallback := self._cached_events_covering(start, end)) is not None:
                    return fallback
                raise

            events = [
                ev
                for raw in raw_events
                if (ev := _graph_event_to_calendar_event(raw)) is not None
            ]
            self._event_cache[key] = (now, start, end, events)
            if len(self._event_cache) > _EVENT_CACHE_MAX_ENTRIES:
                oldest = min(self._event_cache, key=lambda k: self._event_cache[k][0])
                del self._event_cache[oldest]
            return events

    async def _async_after_write(self) -> None:
        """Drop cached events after a change so the next read is fresh."""
        self._event_cache.clear()
        await self._async_refresh_next_event()

    async def _async_refresh_next_event(self) -> None:
        """Find the current or next event and update state.

        The lookup window is aligned to the start of the hour so repeated
        refreshes reuse one cache entry instead of each calling Graph.
        """
        now = dt_util.now()
        start = now.replace(minute=0, second=0, microsecond=0)
        end = start + _NEXT_EVENT_LOOKAHEAD + timedelta(hours=1)
        try:
            events = await self._async_fetch_events(start, end)
        except MicrosoftGraphError as err:
            _LOGGER.debug("Could not refresh next event for %s: %s", self.name, err)
            return

        upcoming = [ev for ev in events if ev.end_datetime_local > now]
        upcoming.sort(key=lambda ev: ev.start_datetime_local)
        self._event = upcoming[0] if upcoming else None
        self.async_write_ha_state()

    # ------------------------------------------------------------------
    # CalendarEntity interface
    # ------------------------------------------------------------------

    async def async_get_events(
        self,
        hass: HomeAssistant,
        start_date: datetime,
        end_date: datetime,
    ) -> list[CalendarEvent]:
        """Return all calendar events within the requested time range."""
        try:
            return await self._async_fetch_events(start_date, end_date)
        except MicrosoftGraphRateLimitError as err:
            _LOGGER.warning(
                "Microsoft Graph is throttling requests and no cached events are "
                "available yet for calendar %s: %s",
                self.name,
                err,
            )
            return []
        except MicrosoftGraphError as err:
            _LOGGER.error("Failed to fetch events for calendar %s: %s", self.name, err)
            return []

    async def async_added_to_hass(self) -> None:
        """Fetch the initial next-event when the entity is first added."""
        await super().async_added_to_hass()
        await self._async_refresh_next_event()

    # ------------------------------------------------------------------
    # Write operations
    # ------------------------------------------------------------------

    async def async_create_event(self, **kwargs: Any) -> None:
        """Create a new single event in this calendar."""
        payload = _event_kwargs_to_graph(kwargs)
        _LOGGER.debug(
            "Adding event %s to calendar %s",
            kwargs.get(EVENT_SUMMARY),
            self._calendar_id,
        )
        await self._client.async_create_event(self._calendar_id, payload)
        await self._async_after_write()

    async def async_update_event(
        self,
        uid: str,
        event: dict[str, Any],
        recurrence_id: str | None = None,
        recurrence_range: str | None = None,
    ) -> None:
        """Update an existing event, blocking modifications to recurring events."""
        payload = _event_kwargs_to_graph(event)
        await self._client.async_update_event(uid, payload)
        await self._async_after_write()

    async def async_delete_event(
        self,
        uid: str,
        recurrence_id: str | None = None,
        recurrence_range: str | None = None,
    ) -> None:
        """Delete an event, blocking deletion of recurring events."""
        await self._client.async_delete_event(uid)
        await self._async_after_write()
