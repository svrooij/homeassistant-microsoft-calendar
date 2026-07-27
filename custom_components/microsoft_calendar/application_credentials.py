"""Application credentials for Microsoft Calendar (PKCE, no client secret)."""

from __future__ import annotations
from typing import override

from homeassistant.components.application_credentials import ClientCredential
from homeassistant.core import HomeAssistant
from homeassistant.helpers.config_entry_oauth2_flow import (
    MY_AUTH_CALLBACK_PATH,
    AbstractOAuth2Implementation,
    LocalOAuth2ImplementationWithPkce,
    async_get_redirect_uri,
)

from .const import DEFAULT_CLIENT_ID, OAUTH2_AUTHORIZE, OAUTH2_TOKEN, SCOPES


class MicrosoftCalendarOAuth2Implementation(LocalOAuth2ImplementationWithPkce):
    """OAuth2 implementation for Microsoft Calendar with PKCE and scopes."""

    @property
    @override
    def name(self) -> str:
        """Return a human-readable name that distinguishes built-in from custom credentials."""
        if self.client_id == DEFAULT_CLIENT_ID:
            return "Microsoft Calendar (Built-in app)"
        return f"Microsoft Calendar ({self.client_id[:8]}…)"

    @property
    @override
    def extra_authorize_data(self) -> dict:
        """Append required scopes to the authorize URL."""
        return super().extra_authorize_data | {
            "scope": " ".join(SCOPES),
        }


async def async_get_auth_implementation(
    hass: HomeAssistant,
    auth_domain: str,
    credential: ClientCredential,
) -> AbstractOAuth2Implementation:
    """Return PKCE-based OAuth2 implementation (no client secret required)."""
    return MicrosoftCalendarOAuth2Implementation(
        hass,
        auth_domain,
        credential.client_id,
        authorize_url=OAUTH2_AUTHORIZE,
        token_url=OAUTH2_TOKEN,
    )


async def async_get_description_placeholders(hass: HomeAssistant) -> dict[str, str]:
    """Return description placeholders shown in the Application Credentials UI."""
    try:
        redirect_url = async_get_redirect_uri(hass)
    except RuntimeError:
        redirect_url = MY_AUTH_CALLBACK_PATH

    return {
        "azure_url": (
            "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps"
            "/ApplicationsListBlade"
        ),
        "redirect_url": redirect_url,
    }
