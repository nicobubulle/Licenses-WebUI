# Microsoft Teams Webhook Setup

This guide explains how to enable and configure Microsoft Teams notifications for Licenses WebUI.

## 1. Overview
Licenses WebUI can send Adaptive Card messages to a Teams channel via an incoming webhook. Supported notification types:
- Update available (GitHub release newer than current version)
- Duplicate license checkout (same user@computer multiple times for a feature)
- Extended usage ("extratime" threshold exceeded; aggregates all long-running features per user)
- Sold-out transitions (feature becomes fully used or becomes available again)

All notifications are optional and individually toggleable in `config.ini`.

## 2. Create an Incoming Webhook in Teams
1. Navigate to the desired Teams channel.
2. Click the three-dot menu (⋯) next to the channel name.
3. Select **Workflows** from the menu.
4. Search for and select **Post to a channel when a webhook request is received**.
5. Follow the prompts to configure the workflow and generate a webhook URL.
6. Copy the webhook URL and keep it secret (treat it like a password).

## 3. Secure Handling of Webhook URL
The webhook URL authorizes posting directly into your channel:
- Do not commit it to public repositories.
- Store it in `config.ini` only on trusted machines.
- Rotate (delete + recreate) if you believe it was exposed.

## 4. Enable Teams in `config.ini`
Example section:
```
[TEAMS]
enabled = yes
webhook = https://outlook.office.com/webhook/XXXXXXXXXXXXXXXXXXXXXXXX
notify_update = yes
notify_duplicate_checker = yes
notify_extratime = yes
extratime_duration = 72
extratime_exclusion = maint-test,temp-feature
notify_soldout = yes
soldout_exclusion = legacy,trial
```

Key descriptions:
- `enabled`: Master switch for all Teams notifications.
- `webhook`: The copied URL from Teams. Quotes are stripped automatically if present.
- `notify_update`: Sends one card when a newer GitHub release is detected (per version).
- `notify_duplicate_checker`: Sends a card for each newly observed duplicate (feature,user,computer) combination.
- `notify_extratime`: Sends one card per (user,computer) listing all features exceeding the threshold once they cross it.
- `extratime_duration`: Threshold in hours (default 72). Lower for testing (e.g., 1 hour).
- `extratime_exclusion`: Comma-separated features to skip from extratime evaluation.
- `notify_soldout`: Sends a card when a feature becomes sold out or becomes available again.
- `soldout_exclusion`: Comma-separated features excluded from sold-out transitions.

Global filters affecting notifications:
- `hide_maintenance = yes` suppresses any feature containing `maint` from display AND all notifications.
- `hide_list` substrings also hide features and suppress related notifications.

## 5. Adaptive Card Structure
Each notification posts an Adaptive Card with:
- Title (e.g., "Feature Sold Out", "Extended Usage Detected")
- Body text (includes details; markdown emphasis supported in basic form)
- Optional "View Details" button for update messages linking to the GitHub release.

Card schema version: `1.2` (compatible with standard Teams rendering). If you need richer formatting, extend `send_teams_notification` in `app.py`.

## 6. De-Duplication Rules
- Update: One notification per discovered version.
- Duplicate: Once per (feature,user,computer); repeated duplicates for same tuple are ignored.
- Extratime: Once per (user,computer); includes all currently exceeding features at send time.
- Sold-out: Only on state transitions (available -> sold out, sold out -> available).

## 7. Testing Notifications Quickly
1. Set `enabled = yes` and supply webhook.
2. Start the application.
3. Trigger manual refresh via POST `/refresh` (UI button) after inducing conditions.
4. Duplicate: Open same feature multiple times under same user@computer (simulate with test data or actual checkout).
5. Extratime: Temporarily set `extratime_duration = 1` and ensure an existing session's start timestamp is >1h old (or edit raw output during testing).
6. Sold-out: Consume all available licenses for a feature (simulate or temporarily adjust total/used in lmstat output).
7. Update: Temporarily bump `VERSION` in `app.py` lower than current GitHub release or create a test release.

## 8. Troubleshooting
- Check `logs/Licenses_WebUI.log` for Teams POST status, HTTP codes, and response body.
- Ensure firewall/proxy does not block outbound HTTPS to the webhook domain.
- Verify `webhook` URL has not expired or been deleted.
- If message formatting seems off, confirm Teams still supports Adaptive Card 1.2 in your tenant.

## 9. Limitations
- No retry/backoff currently—failed sends are logged but not retried.
- Notification state resets on application restart (avoids spam during a single run, but may re-alert after restart). Persisting state can be added if needed.
- Extratime parsing relies on best-effort timestamp formats from lmstat; unrecognized formats are skipped.

## 10. Extending
To add a new notification type:
1. Implement a checker function after license parsing in `refresh_loop`.
2. Maintain a state set/dict to track sent notifications.
3. Use `send_teams_notification(title, message, link=None)` for delivery.
4. Add config toggles similar to existing ones.

## 11. Security Reminder
Treat the webhook like a secret. Anyone with it can post messages to your channel. Rotate periodically.

---
For any issues or feature requests please open an issue on GitHub.
