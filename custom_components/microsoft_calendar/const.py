"""Constants for the Microsoft Calendar integration."""

from __future__ import annotations

from datetime import timedelta

DOMAIN = "microsoft_calendar"

# OAuth2 endpoints (Microsoft identity platform, multi-tenant)
OAUTH2_AUTHORIZE = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
OAUTH2_TOKEN = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

# Scopes always requested (identity + refresh token)
SCOPES_BASE = ["openid", "offline_access", "profile"]

# Calendar-specific scope options the user can choose between
CALENDAR_SCOPE_READ = "Calendars.ReadBasic"
CALENDAR_SCOPE_WRITE = "Calendars.ReadWrite"

# Config entry key that stores the chosen calendar scope
CONF_CALENDAR_SCOPE = "calendar_scope"

# Default scopes (base + read-only) – used as fallback
SCOPES = SCOPES_BASE + [CALENDAR_SCOPE_READ]

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
