"""Constants for the Microsoft Calendar integration."""

from __future__ import annotations

from datetime import timedelta

DOMAIN = "microsoft_calendar"

# OAuth2 endpoints (Microsoft identity platform, multi-tenant)
OAUTH2_AUTHORIZE = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
OAUTH2_TOKEN = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

# Scopes requested during authorization
SCOPES = ["openid", "offline_access", "Calendars.ReadBasic"]

# Microsoft Graph REST API
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# How often to poll for calendar/event changes
DEFAULT_SCAN_INTERVAL = timedelta(minutes=15)

# Keys used in hass.data[DOMAIN][entry_id]
DATA_CLIENT = "client"
DATA_COORDINATOR = "coordinator"

# Built-in (globally registered) public client app.
# This is a public PKCE client — no secret is involved.
# The app registration must have https://my.home-assistant.io/redirect/oauth
# added as a Mobile/Desktop redirect URI in the Microsoft Entra portal.
# Set to an empty string until a real app has been registered.
DEFAULT_CLIENT_ID = "7133658a-31b3-4c66-9e96-89156ecc229e"
