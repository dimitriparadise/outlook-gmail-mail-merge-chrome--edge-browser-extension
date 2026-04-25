# Outlook Mail Merge Helper

This folder contains a small Manifest V3 browser extension that helps create personalized Outlook Web email drafts from a CSV list. The extension runs entirely from the popup UI: users import or paste CSV data, write subject and body templates, preview the generated message, and open one Outlook compose draft at a time.

## What It Does

- Reads recipient data from an uploaded CSV file or pasted CSV text.
- Requires an `Email` or `email` column for usable rows.
- Supports template variables in the form `{{ColumnName}}`.
- Generates a personalized subject and body for each CSV row.
- Opens Outlook Web compose deeplinks for each generated draft in a background tab.
- Saves popup state with `chrome.storage.local`, so progress can continue after closing and reopening the popup.
- Provides a `Copy Body` fallback for manually pasting the generated message body.

## Folder Contents

| File | Purpose |
| --- | --- |
| `manifest.json` | Defines the extension metadata, popup entry point, permissions, and Outlook host permissions. |
| `popup.html` | Defines the popup interface for CSV input, templates, preview, status, and draft controls. |
| `popup.css` | Styles the popup layout, form fields, buttons, status messages, and preview area. |
| `popup.js` | Implements CSV parsing, template replacement, state persistence, preview generation, and Outlook draft opening. |

## How The Extension Works

1. The browser loads `popup.html` as the extension action popup.
2. `popup.js` restores any saved CSV text, templates, generated messages, and progress from `chrome.storage.local`.
3. The user uploads or pastes CSV data.
4. `parseCSV()` converts the CSV text into row objects using the header row as keys.
5. `applyTemplate()` replaces placeholders such as `{{Name}}` or `{{Course}}` with values from each row.
6. `makeMessages()` builds the personalized messages and stores them in memory.
7. `Open Next Draft` increments and saves `currentIndex` before opening the Outlook Web compose URL.
8. The draft opens in a background tab, so the popup can stay open while the user queues additional drafts.

## CSV Format

The CSV must include a header row and at least one data row. It must include either `Email` or `email`.

Example:

```csv
Name,Email,Course
John,john@example.com,ISOM 210
Jane,jane@example.com,ISOM 340
```

Any header can be used as a template variable:

```text
Subject: Reminder for {{Course}}

Body:
Hi {{Name}},

This is a quick reminder about {{Course}}.
```

## Loading The Extension Locally

1. Open a Chromium-based browser such as Chrome or Edge.
2. Go to the browser extensions page.
   - Chrome: `chrome://extensions`
   - Edge: `edge://extensions`
3. Enable developer mode.
4. Choose `Load unpacked`.
5. Select this folder.
6. Click the extension icon to open the popup.

## Permissions

The extension declares:

- `storage`: saves CSV text, templates, generated drafts, and progress locally.
- `tabs`: opens each generated Outlook compose draft in a background tab.
- Outlook host permissions for `https://outlook.office.com/*` and `https://outlook.live.com/*`.

## Implementation Notes

- CSV parsing is implemented locally in `popup.js`. It handles quoted fields and escaped quotes, but it is intentionally lightweight.
- Uploaded file contents are copied into the CSV textarea because browser extension file inputs cannot be restored after the popup closes.
- Outlook compose URLs use `encodeURIComponent()` instead of `URLSearchParams` because Outlook may show `+` characters literally in some compose fields.
- Progress is saved before opening each Outlook tab, because extension popups can close automatically when browser focus changes.
- The extension opens drafts only; it does not automatically send email.

## Current Limitations

- CSV rows without an `Email` or `email` value are ignored.
- There is no attachment support.
- There is no rich-text email editor.
- There is no duplicate-recipient detection.
- Drafts are opened one at a time to keep the user in control of reviewing and sending messages.
