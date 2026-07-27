"""DataUpdateCoordinator for Microsoft Calendar."""

from __future__ import annotations

import logging
from typing import Any

from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.exceptions import ConfigEntryAuthFailed
from homeassistant.helpers.update_coordinator import DataUpdateCoordinator, UpdateFailed

from .api import MicrosoftGraphAuthError, MicrosoftGraphClient, MicrosoftGraphError
from .const import DEFAULT_SCAN_INTERVAL, DOMAIN

_LOGGER = logging.getLogger(__name__)

# Type alias for the coordinator's data payload
type CalendarData = list[dict[str, Any]]


class MicrosoftCalendarCoordinator(DataUpdateCoordinator[CalendarData]):
    """Fetch the list of calendars for the signed-in user.

    Each refresh retrieves the full calendar list from Microsoft Graph.
    Individual event fetches are performed on-demand by the CalendarEntity
    (via async_get_events) and are not cached here.
    """

    def __init__(
        self,
        hass: HomeAssistant,
        entry: ConfigEntry,
        client: MicrosoftGraphClient,
    ) -> None:
        self._client = client
        self._entry = entry
        super().__init__(
            hass,
            _LOGGER,
            name=f"{DOMAIN} ({entry.title})",
            update_interval=DEFAULT_SCAN_INTERVAL,
        )

    async def _async_update_data(self) -> CalendarData:
        """Fetch all calendars from Microsoft Graph.

        Raises:
            ConfigEntryAuthFailed: when the token is invalid/expired, which
                causes HA to start a re-auth flow automatically.
            UpdateFailed: for any other API error.
        """
        try:
            return await self._client.async_list_calendars()
        except MicrosoftGraphAuthError as err:
            # Raising ConfigEntryAuthFailed tells HA to disable the entry and
            # prompt the user to re-authenticate — no manual handling needed.
            raise ConfigEntryAuthFailed(
                f"Microsoft credentials expired or revoked: {err}"
            ) from err
        except MicrosoftGraphError as err:
            raise UpdateFailed(
                f"Error communicating with Microsoft Graph: {err}"
            ) from err
