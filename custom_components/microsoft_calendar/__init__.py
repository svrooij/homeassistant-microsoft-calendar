"""Microsoft Calendar integration."""

from __future__ import annotations

from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers import config_entry_oauth2_flow
from homeassistant.helpers.config_entry_oauth2_flow import (
    LocalOAuth2ImplementationWithPkce,
    OAuth2Session,
)
from homeassistant.helpers.typing import ConfigType

from .api import MicrosoftGraphClient
from .const import (
    DATA_CLIENT,
    DATA_COORDINATOR,
    DEFAULT_CLIENT_ID,
    DOMAIN,
    OAUTH2_AUTHORIZE,
    OAUTH2_TOKEN,
)
from .coordinator import MicrosoftCalendarCoordinator

PLATFORMS: list[str] = ["calendar"]


async def async_setup(hass: HomeAssistant, config: ConfigType) -> bool:
    """Register the built-in default OAuth2 implementation.

    Users who prefer their own Microsoft app registration can add credentials
    via Settings → Devices & services → Application Credentials instead.
    """
    if DEFAULT_CLIENT_ID:
        config_entry_oauth2_flow.async_register_implementation(
            hass,
            DOMAIN,
            LocalOAuth2ImplementationWithPkce(
                hass,
                DOMAIN,
                DEFAULT_CLIENT_ID,
                authorize_url=OAUTH2_AUTHORIZE,
                token_url=OAUTH2_TOKEN,
                client_secret="",
            ),
        )
    return True


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Set up Microsoft Calendar from a config entry."""
    implementation = (
        await config_entry_oauth2_flow.async_get_config_entry_implementation(
            hass, entry
        )
    )
    oauth_session = OAuth2Session(hass, entry, implementation)
    client = MicrosoftGraphClient(oauth_session)

    coordinator = MicrosoftCalendarCoordinator(hass, entry, client)
    await coordinator.async_config_entry_first_refresh()

    hass.data.setdefault(DOMAIN, {})[entry.entry_id] = {
        DATA_CLIENT: client,
        DATA_COORDINATOR: coordinator,
    }

    await hass.config_entries.async_forward_entry_setups(entry, PLATFORMS)
    return True


async def async_unload_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Unload a config entry."""
    if await hass.config_entries.async_unload_platforms(entry, PLATFORMS):
        hass.data[DOMAIN].pop(entry.entry_id)
        return True
    return False
