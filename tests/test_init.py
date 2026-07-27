"""Tests for the integration setup/unload lifecycle (__init__.py)."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from custom_components.microsoft_calendar.const import (
    DATA_CLIENT,
    DATA_COORDINATOR,
    DOMAIN,
)


@pytest.fixture
def mock_config_entry():
    entry = MagicMock()
    entry.entry_id = "test-entry-id"
    entry.title = "Ada Lovelace"
    entry.data = {"token": {"access_token": "tok", "refresh_token": "ref"}}
    return entry


async def test_async_setup_entry_stores_data_and_forwards_platforms(
    mock_config_entry,
):
    """After setup, hass.data contains client and coordinator; calendar platform is loaded."""
    hass = MagicMock()
    hass.data = {}
    hass.config_entries = MagicMock()
    hass.config_entries.async_forward_entry_setups = AsyncMock(return_value=True)

    mock_coordinator = MagicMock()
    mock_coordinator.async_config_entry_first_refresh = AsyncMock()

    with (
        patch(
            "custom_components.microsoft_calendar.config_entry_oauth2_flow"
            ".async_get_config_entry_implementation",
            new=AsyncMock(return_value=MagicMock()),
        ),
        patch(
            "custom_components.microsoft_calendar.OAuth2Session",
            return_value=MagicMock(),
        ),
        patch(
            "custom_components.microsoft_calendar.MicrosoftGraphClient",
            return_value=MagicMock(),
        ),
        patch(
            "custom_components.microsoft_calendar.MicrosoftCalendarCoordinator",
            return_value=mock_coordinator,
        ),
    ):
        from custom_components.microsoft_calendar import async_setup_entry

        result = await async_setup_entry(hass, mock_config_entry)

    assert result is True
    assert DATA_CLIENT in hass.data[DOMAIN][mock_config_entry.entry_id]
    assert DATA_COORDINATOR in hass.data[DOMAIN][mock_config_entry.entry_id]
    hass.config_entries.async_forward_entry_setups.assert_awaited_once()


async def test_async_unload_entry_cleans_up_hass_data(mock_config_entry):
    """After unload, the entry's data is removed from hass.data."""
    hass = MagicMock()
    hass.data = {
        DOMAIN: {
            mock_config_entry.entry_id: {
                DATA_CLIENT: MagicMock(),
                DATA_COORDINATOR: MagicMock(),
            }
        }
    }
    hass.config_entries = MagicMock()
    hass.config_entries.async_unload_platforms = AsyncMock(return_value=True)

    from custom_components.microsoft_calendar import async_unload_entry

    result = await async_unload_entry(hass, mock_config_entry)

    assert result is True
    assert mock_config_entry.entry_id not in hass.data[DOMAIN]


async def test_async_unload_entry_does_not_clean_up_if_platform_unload_fails(
    mock_config_entry,
):
    """If platform unload fails, hass.data must not be wiped."""
    hass = MagicMock()
    hass.data = {
        DOMAIN: {
            mock_config_entry.entry_id: {
                DATA_CLIENT: MagicMock(),
                DATA_COORDINATOR: MagicMock(),
            }
        }
    }
    hass.config_entries = MagicMock()
    hass.config_entries.async_unload_platforms = AsyncMock(return_value=False)

    from custom_components.microsoft_calendar import async_unload_entry

    result = await async_unload_entry(hass, mock_config_entry)

    assert result is False
    assert mock_config_entry.entry_id in hass.data[DOMAIN]
