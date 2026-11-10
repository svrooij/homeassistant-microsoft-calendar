"""Tests for the config flow (config_flow.py)."""

from __future__ import annotations

from datetime import datetime
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from homeassistant.helpers.config_entry_oauth2_flow import OAuth2Session

from custom_components.microsoft_calendar.application_credentials import (
    MicrosoftCalendarOAuth2Implementation,
)
from custom_components.microsoft_calendar.config_flow import (
    MicrosoftCalendarFlowHandler,
    _parse_id_token,
)
from custom_components.microsoft_calendar.const import (
    DEFAULT_CLIENT_ID,
    DOMAIN,
    OAUTH2_AUTHORIZE,
    OAUTH2_TOKEN,
)

from .conftest import make_id_token


# ---------------------------------------------------------------------------
# _parse_id_token
# ---------------------------------------------------------------------------


def test_parse_id_token_extracts_all_claims():
    claims = {
        "oid": "user-object-id-123",
        "name": "Ada Lovelace",
        "preferred_username": "ada@example.com",
        "tid": "tenant-id-456",
    }
    result = _parse_id_token(make_id_token(claims))
    assert result["oid"] == "user-object-id-123"
    assert result["name"] == "Ada Lovelace"
    assert result["preferred_username"] == "ada@example.com"


def test_parse_id_token_not_enough_segments_returns_empty():
    result = _parse_id_token("onlyone")
    assert result == {}


def test_parse_id_token_invalid_base64_returns_empty():
    result = _parse_id_token("header.!!!notbase64!!!.sig")
    assert result == {}


def test_parse_id_token_non_json_payload_returns_empty():
    import base64

    bad_payload = base64.urlsafe_b64encode(b"not json").rstrip(b"=").decode()
    result = _parse_id_token(f"header.{bad_payload}.sig")
    assert result == {}


def test_parse_id_token_extra_padding_handled():
    """Tokens with payloads of varying lengths (padding edge cases) must parse."""
    for extra_claims in [{}, {"x": "y"}, {"a": "b", "c": "d", "e": "f"}]:
        claims = {"oid": "u1", **extra_claims}
        result = _parse_id_token(make_id_token(claims))
        assert result.get("oid") == "u1"


def test_parse_id_token_empty_string_returns_empty():
    result = _parse_id_token("")
    assert result == {}


@pytest.mark.parametrize("expires_in", [3599, "3599"])
async def test_token_expiry_is_stored_and_refreshed(expires_in, freezer, caplog):
    """UTC expiry remains valid despite appearing earlier than the local log time."""
    freezer.move_to("2026-11-09T15:19:40.933061+01:00")
    now = datetime.now().timestamp()
    hass = MagicMock()
    implementation = MicrosoftCalendarOAuth2Implementation(
        hass,
        DOMAIN,
        DEFAULT_CLIENT_ID,
        authorize_url=OAUTH2_AUTHORIZE,
        token_url=OAUTH2_TOKEN,
    )
    token = {
        "access_token": "initial-access",
        "refresh_token": "initial-refresh",
        "id_token": make_id_token({"oid": "user-1"}),
        "expires_in": expires_in,
    }
    flow = MicrosoftCalendarFlowHandler()
    flow.hass = hass
    flow.flow_impl = implementation
    flow.external_data = {
        "code": "authorization-code",
        "state": {"redirect_uri": "https://my.home-assistant.io/redirect/oauth"},
    }
    flow.async_set_unique_id = AsyncMock()
    flow._abort_if_unique_id_configured = MagicMock()
    flow.async_create_entry = MagicMock(side_effect=lambda **kwargs: kwargs)
    caplog.set_level("DEBUG", logger="custom_components.microsoft_calendar")

    with patch.object(
        implementation, "_token_request", new=AsyncMock(return_value=token)
    ):
        result = await flow.async_step_creation()

    entry = MagicMock()
    entry.data = result["data"]
    session = OAuth2Session(hass, entry, implementation)
    assert session.token["expires_in"] == 3599
    assert session.token["expires_at"] == now + 3599
    assert session.valid_token
    assert "2026-11-09T15:19:39.933061+00:00 (3599 seconds remaining)" in caplog.text

    with patch.object(implementation, "_token_request", new=AsyncMock()) as request:
        await session.async_ensure_token_valid()
    request.assert_not_awaited()

    freezer.tick(3599)
    assert not session.valid_token

    def update_entry(config_entry, *, data):
        config_entry.data = data

    hass.config_entries.async_update_entry.side_effect = update_entry
    with patch.object(
        implementation,
        "_token_request",
        new=AsyncMock(
            return_value={
                "access_token": "refreshed-access",
                "refresh_token": "rotated-refresh",
                "expires_in": expires_in,
            }
        ),
    ) as request:
        await session.async_ensure_token_valid()

    assert request.call_args.args[0]["refresh_token"] == "initial-refresh"
    assert session.token["access_token"] == "refreshed-access"
    assert session.token["refresh_token"] == "rotated-refresh"
    assert session.token["expires_at"] == now + 2 * 3599
    assert session.valid_token
    assert "Token refreshed; valid until" in caplog.text
