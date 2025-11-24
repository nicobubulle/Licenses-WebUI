# Licenses WebUI

FLEXnet License Status Web UI — small local web app to query lmutil/lmstat and present license usage with an optional system tray and service restart support.

## Features
- Periodic background refresh of lmstat output
- Manual refresh (runs same parsing + notification checkers as background loop)
- Raw output debug view (`/raw`)
- Windows service restart with robust state checking (requires admin + enabled in config)
- System tray integration (pystray + Pillow)
- Internationalization (i18n JSON files: en, fr, de, es) with query/cookie/header locale negotiation
- Automatic GitHub release check (daily) + optional Teams update notification
- Microsoft Teams notifications (Adaptive Card) for:
   - New version available
   - Duplicate license checkouts (same user@computer multiple times for a feature)
   - Extended usage ("extratime" beyond configurable hours threshold, grouped per user)
   - Sold-out feature transitions (becomes fully used / becomes available again)
- Maintenance filtering: optionally hide and suppress notifications for features containing `maint`
- Additional hide filtering via substring list
- Configurable via `config.ini`

## Quickstart (Windows)
1. Download / build the executable and run it (first start will create `config.ini` if missing).
2. Edit `config.ini` as needed (see below). Restart app after changing values.
3. Browser auto-opens at `http://localhost:<web_port>`.
4. Optional: configure Teams notifications (see `TEAMS_SETUP.md`).

### Core `config.ini` keys (SETTINGS section)
- `lmutil_path`: Full path to `lmutil.exe` (default points to Leica folder).
- `port`: FLEX license manager port (default `27008`).
- `web_port`: Web UI port (default `8080`).
- `refresh_minutes`: Background refresh interval in minutes.
- `default_locale`: Fallback UI locale (`en`, `fr`, `de`, `es`).
- `hide_maintenance`: `yes|no` hide features containing `maint` and suppress related notifications.
- `hide_list`: Comma-separated substrings; any feature containing one is hidden.
- `enable_restart`: Enable Windows service restart button (requires admin elevation on startup).

### SERVICE section
- `service_name`: Display + target for restart functionality.

### TEAMS section (summary)
See `TEAMS_SETUP.md` for full details.
- `enabled`: `yes|no` turns on webhook notifications.
- `webhook`: Incoming webhook URL (keep secret).
- `notify_update`: Notify when a newer GitHub release is found.
- `notify_duplicate_checker`: Duplicate checkout alerts (one per (feature,user,computer)).
- `notify_extratime`: Extended usage alerts (one per (user,computer)).
- `extratime_duration`: Threshold hours (default 72).
- `extratime_exclusion`: Comma-separated features to skip for extratime.
- `notify_soldout`: Sold-out transition alerts.
- `soldout_exclusion`: Comma-separated features to skip for sold-out.

Manual refresh (`POST /refresh`) performs the same parsing and runs duplicate, extratime, and sold-out checkers immediately.

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
`config.ini` is created automatically with example values. Edit and restart the application. Unknown keys are ignored. Percent symbols (%) in webhook URLs are preserved using raw read mode.

Example TEAMS block:
```
[TEAMS]
enabled = yes
webhook = https://outlook.office.com/webhook/....
notify_update = yes
notify_duplicate_checker = yes
notify_extratime = yes
extratime_duration = 72
extratime_exclusion = maint-test,temp-feature
notify_soldout = yes
soldout_exclusion = legacy,trial
```

## Logs & Troubleshooting
- Logs are written to `logs/Licenses_WebUI.log`.
- If the service restart fails, detailed `sc` output is captured in the returned log for diagnosis.
- Ensure `lmutil.exe` path is correct and that the license manager is reachable.

## Development & Packaging notes
- The app requests elevation at startup if `enable_restart = yes`; accept UAC for restart capability.
- Threads: refresh loop, update check loop (daily), systray, optional browser opener.
- Avoid calling request-dependent functions in background threads (already handled).

## Contributing
- Open issues or PRs.
- Keep translations in `i18n/` and update `SUPPORTED_LOCALES` as needed.
 - Include new notification types with clear state tracking to prevent spam.

## Microsoft Teams Integration
For full setup instructions see `TEAMS_SETUP.md`. Adaptive Card payload is sent; messages appear with title, body, and optional "View Details" link for update notifications.

Notification de-duplication rules:
- Update: once per discovered version.
- Duplicate: once per (feature,user,computer) combination.
- Extratime: once per (user,computer) after threshold; aggregates all exceeding features.
- Sold-out: on state transitions (sold out -> available / available -> sold out).

To test quickly:
1. Set `enabled = yes` and supply webhook.
2. Trigger manual refresh or create duplicate sessions.
3. Temporarily lower `extratime_duration` to a small number (e.g., 1) to force extended usage notifications (requires sessions older than threshold).
4. Simulate sold-out by exhausting all licenses for a feature.

## License
MIT