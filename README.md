# Outlook Mail Merge Helper

This folder contains a small Manifest V3 browser extension that helps create personalized Outlook Web email drafts from a CSV list. The extension runs entirely from the popup UI: users import or paste CSV data, write subject and body templates, load the CSV to preview the first generated message, and open Outlook compose drafts one at a time or by selected range.

## What It Does

- Reads recipient data from an uploaded CSV file or pasted CSV text.
- Requires an `Email` or `email` column for usable rows.
- Supports template variables in the form `{{ColumnName}}`.
- Shows variable buttons from the CSV headers so users can insert columns into templates.
- Generates personalized To, CC, BCC, subject, and body values for each CSV row.
- Opens Outlook Web compose deeplinks for each generated draft in a background tab.
- Can open the next draft only or open a selected draft range at once.
- Saves popup state with `chrome.storage.local`, so progress can continue after closing and reopening the popup.

## Folder Contents

| File | Purpose |
| --- | --- |
| `manifest.json` | Defines the extension metadata, popup entry point, permissions, and Outlook host permissions. |
| `popup.html` | Defines the popup interface for CSV input, variable buttons, recipient templates, preview, status, and draft controls. |
| `popup.css` | Styles the popup layout, form fields, buttons, status messages, and preview area. |
| `popup.js` | Implements CSV parsing, template replacement, state persistence, preview generation, and Outlook draft opening. |

## How The Extension Works

1. The browser loads `popup.html` as the extension action popup.
2. `popup.js` restores any saved CSV text, templates, generated messages, and progress from `chrome.storage.local`.
3. The user uploads or pastes CSV data.
4. `parseCSV()` converts the CSV text into row objects using the header row as keys.
5. `Generate Drafts` renders one insert button for each CSV header, such as `{{Name}}`, `{{Course}}`, or `{{DueDate}}`.
6. `applyTemplate()` replaces placeholders with values from each row.
7. `makeMessages()` builds the personalized To, CC, BCC, subject, and body values and stores them in memory.
8. The popup shows the first generated message as a preview.
9. `Open Next Draft` increments and saves `currentIndex` before opening one Outlook Web compose URL.
10. `Open Range` saves progress and opens the selected 1-based draft range in background tabs.
11. Drafts open in background tabs, so the popup can stay open while the user queues additional drafts.

## CSV Format

The CSV must include a header row and at least one data row. It must include either `Email` or `email`.

Example:

```csv
Name,Email,Course,Section,DueDate,CcEmail,BccEmail
John,john@example.com,ISOM 210,A,Friday,ta@example.com,archive@example.com
Jane,jane@example.com,ISOM 340,B,Monday,ta@example.com,archive@example.com
```

Any header can be used as a template variable:

```text
CC: {{CcEmail}}
BCC: {{BccEmail}}
Subject: Reminder for {{Course}}

Body:
Hi {{Name}},

This is a quick reminder about {{Course}} section {{Section}}, due {{DueDate}}.
```

The `Email` or `email` column is used as the main To recipient. Optional CC and BCC fields can use fixed email addresses or variables from the CSV.

## CC And BCC

The optional CC and BCC fields can be left blank, filled with fixed email addresses, or filled with CSV variables.

Examples:

```text
CC: ta@example.com
BCC: archive@example.com
```

```text
CC: {{CcEmail}}
BCC: {{BccEmail}}
```

Outlook Web compose deeplinks may ignore CC and BCC query parameters. Because of that, this extension uses `mailto:` whenever a generated draft has CC or BCC recipients. The draft will open in whichever mail app or webmail service Chrome is configured to use for `mailto:` links.

## Setting Chrome Mailto Handling

Use these steps if CC/BCC drafts open in the wrong app, or if nothing opens when using CC/BCC.

1. Open Chrome.
2. Go to `chrome://settings/handlers`.
3. Turn on `Sites can ask to handle protocols`.
4. Open the mail service you want to use for `mailto:` links, such as Gmail or Outlook Web, in Chrome.
5. If Chrome shows a protocol-handler icon in the address bar, usually a double-diamond icon, click it.
6. Choose `Allow`, then confirm with `Done`.
7. Return to `chrome://settings/handlers` and confirm the site is listed as the handler for email or `mailto:`.

If the handler icon does not appear, remove old blocked/default email handlers from `chrome://settings/handlers`, refresh the mail site, and try again. On macOS or Windows, you may also need to set Chrome or your preferred mail app as the system default email app.

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
- Outlook Web deeplinks may ignore CC/BCC, so messages with CC or BCC use `mailto:` and depend on the user's configured mail handler.
- Progress is saved before opening each Outlook tab, because extension popups can close automatically when browser focus changes.
- The extension opens drafts only; it does not automatically send email.

## Current Limitations

- CSV rows without an `Email` or `email` value are ignored.
- There is no attachment support.
- There is no rich-text email editor.
- There is no duplicate-recipient detection.
- Opening a large range at once may trigger browser popup/tab throttling if the range is very large.
