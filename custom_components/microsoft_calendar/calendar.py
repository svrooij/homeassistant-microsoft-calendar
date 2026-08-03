"""Calendar platform for Microsoft Calendar."""

from __future__ import annotations

import logging
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

from .api import MicrosoftGraphClient, MicrosoftGraphError
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

    async def _async_refresh_next_event(self) -> None:
        """Fetch the nearest upcoming event and update state."""
        now = dt_util.now()
        end = now + _NEXT_EVENT_LOOKAHEAD
        try:
            raw_events = await self._client.async_list_calendar_events(
                self._calendar_id, now, end
            )
        except MicrosoftGraphError as err:
            _LOGGER.debug("Could not refresh next event for %s: %s", self.name, err)
            return

        events = [
            ev
            for raw in raw_events
            if (ev := _graph_event_to_calendar_event(raw)) is not None
        ]
        self._event = events[0] if events else None
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
            raw_events = await self._client.async_list_calendar_events(
                self._calendar_id, start_date, end_date
            )
        except MicrosoftGraphError as err:
            _LOGGER.error("Failed to fetch events for calendar %s: %s", self.name, err)
            return []

        return [
            ev
            for raw in raw_events
            if (ev := _graph_event_to_calendar_event(raw)) is not None
        ]

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
        await self._client.async_create_event(self._calendar_id, payload)

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

    async def async_delete_event(
        self,
        uid: str,
        recurrence_id: str | None = None,
        recurrence_range: str | None = None,
    ) -> None:
        """Delete an event, blocking deletion of recurring events."""
        await self._client.async_delete_event(uid)
