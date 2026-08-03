# Microsoft Calendar developer guide

## Development

The repository includes a VS Code dev container with Home Assistant pre-installed.

```bash
# Open in dev container, then:

# Run Home Assistant
python3 -m homeassistant --debug -c config --skip-pip

# Run tests
python3 -m pytest tests/ -v
```

## Releasing

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

## Project structure

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
