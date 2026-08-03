"""Config flow for Microsoft Calendar."""

from __future__ import annotations

import base64
import json
import logging
from typing import Any

import voluptuous as vol
from homeassistant.config_entries import SOURCE_REAUTH
from homeassistant.helpers import config_entry_oauth2_flow
from homeassistant.helpers.selector import (
    SelectOptionDict,
    SelectSelector,
    SelectSelectorConfig,
    SelectSelectorMode,
)

from .application_credentials import MicrosoftCalendarOAuth2Implementation
from .const import (
    CALENDAR_SCOPE_READ,
    CALENDAR_SCOPE_WRITE,
    CONF_CALENDAR_SCOPE,
    DEFAULT_CLIENT_ID,
    DOMAIN,
    OAUTH2_AUTHORIZE,
    OAUTH2_TOKEN,
    SCOPES_BASE,
)

_LOGGER = logging.getLogger(__name__)


def _parse_id_token(id_token: str) -> dict[str, Any]:
    """Decode the payload of a JWT id_token without signature verification.

    The token was received directly from Microsoft's token endpoint over TLS,
    so we trust the content — we only need to read the claims, not verify them.
    """
    try:
        payload_b64 = id_token.split(".")[1]
        # JWT base64url encoding omits padding; restore it before decoding.
        payload_b64 += "=" * (4 - len(payload_b64) % 4)
        return json.loads(base64.urlsafe_b64decode(payload_b64))
    except Exception as err:  # noqa: BLE001
        _LOGGER.error("Failed to decode id_token payload: %s", err)
        return {}


class MicrosoftCalendarFlowHandler(
    config_entry_oauth2_flow.AbstractOAuth2FlowHandler,
    domain=DOMAIN,
):
    """Handle the OAuth2 config flow for Microsoft Calendar."""

    DOMAIN = DOMAIN

    def __init__(self) -> None:
        """Initialise the flow."""
        super().__init__()
        self._calendar_scope: str = CALENDAR_SCOPE_READ

    @property
    def extra_authorize_data(self) -> dict:
        """Append the user-selected calendar scope to the authorize URL."""
        return {"scope": " ".join(SCOPES_BASE + [self._calendar_scope])}

    def is_matching(self, other_flow: Any) -> bool:
        """Not used — this integration is not discoverable."""
        return False

    @property
    def logger(self) -> logging.Logger:
        """Return logger."""
        return _LOGGER

    async def async_step_user(self, user_input: dict[str, Any] | None = None) -> Any:
        """Start the flow by asking the user to choose a calendar access level."""
        return await self.async_step_scope()

    async def async_step_pick_implementation(
        self, user_input: dict[str, Any] | None = None
    ) -> Any:
        """Register the built-in implementation if needed, then proceed normally.

        async_setup is only called during HA startup (when a config entry already
        exists), NOT when a brand-new config flow is started for the first time.
        Registering here ensures the built-in credential is always available
        regardless of whether async_setup has run yet.
        """
        if DEFAULT_CLIENT_ID:
            config_entry_oauth2_flow.async_register_implementation(
                self.hass,
                DOMAIN,
                MicrosoftCalendarOAuth2Implementation(
                    self.hass,
                    DOMAIN,
                    DEFAULT_CLIENT_ID,
                    authorize_url=OAUTH2_AUTHORIZE,
                    token_url=OAUTH2_TOKEN,
                ),
            )
        return await super().async_step_pick_implementation(user_input)

    async def async_step_scope(self, user_input: dict[str, Any] | None = None) -> Any:
        """Ask the user which calendar permission level they want."""
        if user_input is None:
            return self.async_show_form(
                step_id="scope",
                data_schema=vol.Schema(
                    {
                        vol.Required(
                            CONF_CALENDAR_SCOPE, default=CALENDAR_SCOPE_READ
                        ): SelectSelector(
                            SelectSelectorConfig(
                                options=[
                                    SelectOptionDict(
                                        value=CALENDAR_SCOPE_READ,
                                        label="Read-only (Calendars.ReadBasic)",
                                    ),
                                    SelectOptionDict(
                                        value=CALENDAR_SCOPE_WRITE,
                                        label="Read & write (Calendars.ReadWrite)",
                                    ),
                                ],
                                mode=SelectSelectorMode.LIST,
                            )
                        )
                    }
                ),
            )
        self._calendar_scope = user_input[CONF_CALENDAR_SCOPE]
        return await self.async_step_pick_implementation()

    async def async_oauth_create_entry(self, data: dict[str, Any]) -> Any:
        """Finalise the config entry after a successful OAuth2 flow.

        Extracts the user identity from the id_token that Microsoft returns
        as part of the token response (openid scope), avoiding an extra
        network round-trip to /me.
        """
        id_token: str | None = data.get("token", {}).get("id_token")
        if not id_token:
            _LOGGER.error("No id_token in token response; cannot identify user")
            return self.async_abort(reason="oauth_error")

        claims = _parse_id_token(id_token)

        # oid is the stable object ID for the user in their tenant.
        # sub is the OIDC subject claim, always present and also stable per app.
        # Prefer oid (tenant-wide), fall back to sub.
        oid: str | None = claims.get("oid") or claims.get("sub")
        if not oid:
            _LOGGER.error("id_token missing both 'oid' and 'sub' claims")
            return self.async_abort(reason="oauth_error")

        upn: str = claims.get("preferred_username") or oid
        display_name: str = claims.get("name") or upn

        await self.async_set_unique_id(oid)

        if self.source == SOURCE_REAUTH:
            self._abort_if_unique_id_mismatch(reason="wrong_account")
            return self.async_update_reload_and_abort(
                self._get_reauth_entry(),
                data={**data, CONF_CALENDAR_SCOPE: self._calendar_scope},
            )

        self._abort_if_unique_id_configured()
        return self.async_create_entry(
            title=display_name,
            data={**data, CONF_CALENDAR_SCOPE: self._calendar_scope},
        )

    async def async_step_reauth(self, entry_data: dict[str, Any]) -> Any:
        """Start re-authentication; preserve the existing calendar scope."""
        self._calendar_scope = entry_data.get(CONF_CALENDAR_SCOPE, CALENDAR_SCOPE_READ)
        return await self.async_step_reauth_confirm()

    async def async_step_reauth_confirm(
        self, user_input: dict[str, Any] | None = None
    ) -> Any:
        """Show a confirmation form before restarting the OAuth2 flow."""
        if user_input is None:
            return self.async_show_form(step_id="reauth_confirm")
        return await self.async_step_pick_implementation()
