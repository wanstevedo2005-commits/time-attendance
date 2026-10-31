# Time Attendance

This project has a LIFF frontend (`index.html`) and Google Apps Script backend (`รหัส.js`). The frontend posts to `GAS_API_URL`; set it to the Web App URL of the same Apps Script project that contains this backend.

The backend prevents more than one check-in per LINE user per Bangkok calendar day. Employee names are read from the `Employee` sheet first (column A: LINE user ID, column B: name), then fall back to an existing name in `Log`.

## Required Script Properties

In Apps Script **Project Settings → Script properties**, add:

| Property | Value |
| --- | --- |
| `FOLDER_ID` | ID of the Drive folder for attendance photos |
| `TELEGRAM_TOKEN` | Telegram bot token |
| `TELEGRAM_CHAT_IDS` | Comma-separated recipient IDs, e.g. `111,222` |
| `ADMIN_API_KEY` | A new random secret needed by the dashboard log/name APIs |

Optional: `TIMEZONE` (default `Asia/Bangkok`). The existing late rule remains after 08:40.

## Deploy

1. Set the Script Properties above.
2. Push with clasp.
3. Update the existing Web App deployment to a **New version**; retain its URL when that is the URL in `GAS_API_URL`.
4. Run `testTelegramDirect` once. Execution logs show an independent result for every chat ID.
5. Confirm the deployed web app accepts a LIFF check-in, rejects a second check-in on the same day, and writes the employee name from `Employee`.

The dashboard must add `apiKey=<ADMIN_API_KEY>` to its GET requests. Photos are private by default; share the Drive folder only with managers who need access.

Secrets are stored only in Script Properties, never in Git.
