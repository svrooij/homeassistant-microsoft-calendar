# Microsoft 365 Calendar integration for Home Assistant

[![HACS Default][hacs-badge]][hacs-url]
[![GitHub Actions][ci-badge]][ci-url]
[![HA min version][ha-badge]][ha-url]
![GitHub Release][release-badge]
[![Downloads for latest release][downloads-badge]][release-link] ![GitHub Downloads (all assets, all releases)][download-badge]

A custom Home Assistant integration that exposes your **Microsoft / Outlook / Microsoft 365 calendars** as native [calendar entities][ha-calendar-docs]. Each calendar becomes its own entity you can use in dashboards, automations, and scripts.

Authentication uses the **OAuth 2.0 Authorization Code + PKCE** flow — no client secret is stored anywhere.

There is an **application built-in**, no need to register your own. Off-course you can [register](#option-b--your-own-app-registration) your own application, if you want.

## Installation

This app is available in the HACS default repository, since `2026-08-02`.

### Via HACS (recommended)

1. In Home Assistant open **HACS → Integrations**
2. Search for **Microsoft Calendar** and click **Download**.
3. Restart Home Assistant.

### Via HACS (custom repository)

1. In Home Assistant open **HACS → Integrations → ⋮ → Custom repositories**.
2. Add `https://github.com/svrooij/homeassistant-microsoft-calendar` with category **Integration**.
3. Search for **Microsoft Calendar** and click **Download**.
4. Restart Home Assistant.

### Manual

Download the [latest release](https://github.com/svrooij/homeassistant-microsoft-calendar/releases) and extract the `microsoft_calendar.zip` file in your Home Assistant `config/custom_components/microsoft_calendar` directory and restart.

## Features

- One calendar entity per Microsoft calendar (work, personal, shared)
- `STATE_ON` when an event is currently active; `STATE_OFF` otherwise
- Full event browsing in the Home Assistant calendar UI
- Recurring events automatically expanded by Microsoft Graph — no client-side logic needed
- Supports both **Microsoft 365 (work/school)** and **personal Microsoft accounts**
- Token refresh handled automatically; re-auth flow triggered when the refresh token expires
- No third-party Python SDK — only standard HTTPS calls to Microsoft Graph

## Setup

There are two ways to authenticate. Pick the one that fits your setup.

### Option A — Built-in app (easiest)

A Microsoft Entra app is pre-registered for this integration with the redirect URI `https://my.home-assistant.io/redirect/oauth`. This works for **any Home Assistant user** who can reach `my.home-assistant.io` — you do not need a paid Nabu Casa subscription.

1. In Home Assistant go to **Settings → Devices & services → Add integration**.
2. Search for **Microsoft Calendar** and select it.
3. Select if you want read-only or read-write access to your calendar.
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
2. Add the following permissions (only one of the calendar permissions is required):

   | Permission | Purpose |
   |---|---|
   | `openid` | Sign-in and identity |
   | `offline_access` | Refresh tokens (stay signed in) |
   | `Calendars.ReadBasic` | List calendars and read events |
   | `Calendars.ReadWrite` | List calendars and read events, and modify it |

3. Click **Add permissions**. (No admin consent required for these delegated permissions.)

#### 4 — Copy the client ID

On the **Overview** page copy the **Application (client) ID** — you will need it in the next step.

#### 5 — Add credentials in Home Assistant

1. Go to **Settings → Devices & services → ⋮ → Application credentials**.
2. Click **Add application credential** and select **Microsoft Calendar**.
3. Paste the **Application (client) ID** into the **Client ID** field.
4. Enter any placeholder (e.g. `xxx`) in the **Client Secret** field — this integration uses PKCE and the value is ignored.
5. Click **Create**.

#### 6 — Add the integration

1. Go to **Settings → Devices & services → Add integration**.
2. Search for **Microsoft Calendar**.
3. Choose your newly added credential when prompted.
4. Sign in with your Microsoft account.

### Permissions explained

| Permission | Why it's needed |
|---|---|
| `openid` | Identifies the signed-in user from the `id_token` (no extra `/me` call needed) |
| `offline_access` | Allows Home Assistant to refresh the access token in the background |
| `Calendars.ReadBasic` | Minimum scope to list all calendars and read event details |
| `Calendars.ReadWrite` | Read and edit access to your calendar |

Your choice whether or not you want to give this app Read or Read-Write access to your calendar. All code is [public][repo-link], so you can verify what happens with the code.

## Known limitations

- **No push notifications** — the integration polls Microsoft Graph every 15 minutes.
- **Shared calendars** — visible only if they appear in the signed-in user's calendar list.


[hacs-badge]: https://img.shields.io/badge/HACS-default-blue.svg?style=for-the-badge&logo=homeassistantcommunitystore&logoColor=white
[hacs-url]: https://github.com/hacs/integration
[ci-badge]: https://img.shields.io/github/actions/workflow/status/svrooij/homeassistant-microsoft-calendar/tests.yml?style=for-the-badge&logo=github&logoColor=white
[ci-url]: https://github.com/svrooij/homeassistant-microsoft-calendar/actions/workflows/tests.yml
[ha-badge]: https://img.shields.io/badge/HA-2025.4%2B-blue.svg?style=for-the-badge&logo=homeassistant&logoColor=white
[ha-url]: https://www.home-assistant.io/
[ha-calendar-docs]: https://developers.home-assistant.io/docs/core/entity/calendar/
[release-badge]: https://img.shields.io/github/v/release/svrooij/homeassistant-microsoft-calendar?style=for-the-badge
[downloads-badge]: https://img.shields.io/github/downloads/svrooij/homeassistant-microsoft-calendar/latest/total.svg?style=for-the-badge
[release-link]: https://github.com/svrooij/homeassistant-microsoft-calendar/releases/latest
[download-badge]: https://img.shields.io/github/downloads/svrooij/homeassistant-microsoft-calendar/total?label=downloads%40all&style=for-the-badge
[repo-link]: https://github.com/svrooij/homeassistant-microsoft-calendar/