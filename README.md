# Licenses WebUI

FLEXnet License Status Web UI — small local web app to query lmutil/lmstat and present license usage with an optional system tray and service restart support.

## Features
- Periodic background refresh of lmstat output
- Manual refresh and raw output debug view
- Windows service restart with robust state checking (requires admin)
- System tray integration (pystray + Pillow)
- Internationalization (i18n JSON files: en, fr, de, es)
- Configurable via `config.ini`

## Quickstart (Windows)
1. Download the exe file and run it
2. You can edit or create a `config.ini` (created automatically if missing):
   - SETTINGS.lmutil_path — path to `lmutil.exe`
   - SETTINGS.port — license manager TCP port (default 27008)
   - SETTINGS.web_port — web UI port (default 8080)
   - SETTINGS.refresh_minutes — auto-refresh interval (minutes)
   - SETTINGS.hide_maintenance / hide_list (comma separated) — filtering
   - SETTINGS.enable_restart — enable/disable restart button
   - SERVICE.service_name — Windows service name used for restart

   ```
   The app will attempt initial lmstat fetch and open a browser to http://localhost:8080 (or configured port).

## Endpoints
- `/` — main UI
- `/status` — JSON status (licenses + last_update)
- `/refresh` — POST to force synchronous refresh
- `/restart` — POST to request service restart (admin + enable_restart required)
- `/raw` — raw lmstat output for debugging

## Internationalization
Translation files live in `i18n/` as JSON. Supported locales are loaded from `app.py` (DEFAULT_LOCALE and SUPPORTED_LOCALES). To add a language:
1. Create `i18n/xx.json` (xx = locale code).
2. Include the same keys as `en.json` and translated values.
3. Add locale code to `SUPPORTED_LOCALES` in `app.py` if needed.

Locale negotiation: `?lang=xx` query param → `lang` cookie → `Accept-Language` → default.

## Configuration details
- `config.ini` created automatically if missing with example values.

## Logs & Troubleshooting
- Logs are written to `logs/Licenses_WebUI.log`.
- If the service restart fails, detailed `sc` output is captured in the returned log for diagnosis.
- Ensure `lmutil.exe` path is correct and that the license manager is reachable.

## Development & Packaging notes
- The app requests elevation at startup to enable service restart; run as admin or accept UAC.

## Contributing
- Open issues or PRs.
- Keep translations in `i18n/` and update `SUPPORTED_LOCALES` as needed.

## License
MIT