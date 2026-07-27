# Microsoft Calendar for Home Assistant

[![HACS Custom][hacs-badge]][hacs-url]
[![GitHub Actions][ci-badge]][ci-url]
[![HA min version][ha-badge]][ha-url]

A custom Home Assistant integration that exposes your **Microsoft / Outlook / Microsoft 365 calendars** as native [calendar entities][ha-calendar-docs]. Each calendar becomes its own entity you can use in dashboards, automations, and scripts.

Authentication uses the **OAuth 2.0 Authorization Code + PKCE** flow — no client secret is stored anywhere.

---

## Features

- One calendar entity per Microsoft calendar (work, personal, shared)
- `STATE_ON` when an event is currently active; `STATE_OFF` otherwise
- Full event browsing in the Home Assistant calendar UI
- Recurring events automatically expanded by Microsoft Graph — no client-side logic needed
- Supports both **Microsoft 365 (work/school)** and **personal Microsoft accounts**
- Token refresh handled automatically; re-auth flow triggered when the refresh token expires
- No third-party Python SDK — only standard HTTPS calls to Microsoft Graph

---

## Requirements

| Requirement | Minimum version |
|---|---|
| Home Assistant | **2025.4.0** (when `LocalOAuth2ImplementationWithPkce` was introduced) |
| HACS | 2.x |
| Microsoft account | Personal, work, or school |

---

## Installation

### Via HACS (recommended)

1. In Home Assistant open **HACS → Integrations → ⋮ → Custom repositories**.
2. Add `https://github.com/svrooij/homeassistant-microsoft-calendar` with category **Integration**.
3. Search for **Microsoft Calendar** and click **Download**.
4. Restart Home Assistant.

### Manual

Copy the `custom_components/microsoft_calendar/` folder into your Home Assistant `config/custom_components/` directory and restart.

---

## Setup

There are two ways to authenticate. Pick the one that fits your setup.

### Option A — Built-in app (easiest)

A Microsoft Entra app is pre-registered for this integration with the redirect URI `https://my.home-assistant.io/redirect/oauth`. This works for **any Home Assistant user** who can reach `my.home-assistant.io` — you do not need a paid Nabu Casa subscription.

1. In Home Assistant go to **Settings → Devices & services → Add integration**.
2. Search for **Microsoft Calendar** and select it.
3. Choose **Microsoft Calendar** (the built-in option) when prompted to pick an implementation.
4. Sign in with your Microsoft account and grant the requested permissions.
5. Done — your calendars will appear as entities.

### Option B — Your own app registration

Use this if you want full control over the Microsoft app, need a different redirect URI, or the built-in app doesn't work in your environment.

#### 1 — Create the app registration

1. Go to the [Microsoft Entra admin center](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) and sign in.
2. Click **New registration**.
3. Give it a name (e.g. `Home Assistant Calendar`).
4. Under **Supported account types** select:
   **Accounts in any organizational directory and personal Microsoft accounts**.
5. Click **Register**.

#### 2 — Configure authentication

1. In your new app, go to **Authentication → Add a platform → Mobile and desktop applications**.
2. Add one or both of the following redirect URIs:

   | Redirect URI | When to use |
   |---|---|
   | `https://my.home-assistant.io/redirect/oauth` | Recommended — works for any HA instance reachable via `my.home-assistant.io` |
   | `https://<your-ha-url>/auth/external/callback` | Use this if `my.home-assistant.io` is not available for your instance |

3. Scroll down and enable **Allow public client flows**.
4. Click **Save**.

#### 3 — Add API permissions

1. Go to **API permissions → Add a permission → Microsoft Graph → Delegated permissions**.
2. Add the following permissions:

   | Permission | Purpose |
   |---|---|
   | `openid` | Sign-in and identity |
   | `offline_access` | Refresh tokens (stay signed in) |
   | `Calendars.ReadBasic` | List calendars and read events |

3. Click **Add permissions**. (No admin consent required for these delegated permissions.)

#### 4 — Copy the client ID

On the **Overview** page copy the **Application (client) ID** — you will need it in the next step.

#### 5 — Add credentials in Home Assistant

1. Go to **Settings → Devices & services → ⋮ → Application credentials**.
2. Click **Add application credential** and select **Microsoft Calendar**.
3. Paste the **Application (client) ID** into the **Client ID** field.
4. Enter any placeholder (e.g. `not-used`) in the **Client Secret** field — this integration uses PKCE and the value is ignored.
5. Click **Create**.

#### 6 — Add the integration

1. Go to **Settings → Devices & services → Add integration**.
2. Search for **Microsoft Calendar**.
3. Choose your newly added credential when prompted.
4. Sign in with your Microsoft account.

---

## Permissions explained

| Permission | Why it's needed |
|---|---|
| `openid` | Identifies the signed-in user from the `id_token` (no extra `/me` call needed) |
| `offline_access` | Allows Home Assistant to refresh the access token in the background |
| `Calendars.ReadBasic` | Minimum scope to list all calendars and read event details |

No write permissions are requested. The integration is **read-only**.

---

## Known limitations

- **Read-only** — creating, updating, or deleting events is not supported in v0.1.
- **No push notifications** — the integration polls Microsoft Graph every 15 minutes.
- **Shared calendars** — visible only if they appear in the signed-in user's calendar list.

---

## Development

The repository includes a VS Code dev container with Home Assistant pre-installed.

```bash
# Open in dev container, then:

# Run Home Assistant
python3 -m homeassistant --debug -c config --skip-pip

# Run tests
python3 -m pytest tests/ -v
```

### Releasing

The git tag is the single source of truth for the version — `manifest.json` does **not** need to be updated manually.

```bash
git tag v0.2.0
git push origin v0.2.0
```

The [release workflow](.github/workflows/release.yml) will:
1. Run the test suite (fails fast on any broken test)
2. Inject the version from the tag into `manifest.json`
3. Package `custom_components/microsoft_calendar/` as `microsoft_calendar.zip`
4. Publish a GitHub Release with the zip and auto-generated release notes

HACS downloads the zip (configured via `zip_release: true` in `hacs.json`) and notifies users of the update.

### Project structure

```
custom_components/microsoft_calendar/
├── __init__.py               # Entry setup / teardown
├── manifest.json
├── const.py                  # Constants (URLs, scopes, keys)
├── application_credentials.py# PKCE OAuth2 implementation
├── config_flow.py            # OAuth2 config flow + re-auth
├── api.py                    # Microsoft Graph HTTPS client (no SDK)
├── coordinator.py            # DataUpdateCoordinator
├── calendar.py               # CalendarEntity platform
├── strings.json
└── translations/en.json
tests/
├── test_api.py
├── test_calendar.py
├── test_config_flow.py
└── test_init.py
```

---

[hacs-badge]: https://img.shields.io/badge/HACS-Custom-orange.svg
[hacs-url]: https://github.com/hacs/integration
[ci-badge]: https://github.com/svrooij/homeassistant-microsoft-calendar/actions/workflows/tests.yml/badge.svg
[ci-url]: https://github.com/svrooij/homeassistant-microsoft-calendar/actions/workflows/tests.yml
[ha-badge]: https://img.shields.io/badge/HA-2025.4%2B-blue.svg
[ha-url]: https://www.home-assistant.io/
[ha-calendar-docs]: https://developers.home-assistant.io/docs/core/entity/calendar/
