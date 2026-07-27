"""Config flow for Microsoft Calendar."""

from __future__ import annotations

import base64
import json
import logging
from typing import Any

from homeassistant.config_entries import SOURCE_REAUTH
from homeassistant.helpers import config_entry_oauth2_flow

from .const import DOMAIN

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

    @property
    def logger(self) -> logging.Logger:
        """Return logger."""
        return _LOGGER

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
        # It does not change even when the user's UPN or email changes.
        oid: str | None = claims.get("oid")
        if not oid:
            _LOGGER.error("id_token missing 'oid' claim")
            return self.async_abort(reason="oauth_error")

        upn: str = claims.get("preferred_username") or oid
        display_name: str = claims.get("name") or upn

        await self.async_set_unique_id(oid)

        if self.source == SOURCE_REAUTH:
            self._abort_if_unique_id_mismatch(reason="wrong_account")
            return self.async_update_reload_and_abort(
                self._get_reauth_entry(), data=data
            )

        self._abort_if_unique_id_configured()
        return self.async_create_entry(title=display_name, data=data)

    async def async_step_reauth(self, entry_data: dict[str, Any]) -> Any:
        """Start re-authentication when the token has expired or been revoked."""
        return await self.async_step_reauth_confirm()

    async def async_step_reauth_confirm(
        self, user_input: dict[str, Any] | None = None
    ) -> Any:
        """Show a confirmation form before restarting the OAuth2 flow."""
        if user_input is None:
            return self.async_show_form(step_id="reauth_confirm")
        return await self.async_step_user()
