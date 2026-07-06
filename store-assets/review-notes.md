# Chrome Web Store Review Notes

Mail Merge Draft Helper is a Manifest V3 extension that helps users generate personalized Gmail or Outlook Web email drafts from CSV data.

## Core User Flow

1. Open the extension popup.
2. Paste CSV text or upload a CSV file with an `Email` column.
3. Select Gmail Mode or Outlook Mode.
4. Edit the subject/body template and optional CC/BCC fields.
5. Click `Generate Drafts`.
6. Review each generated message in the popup.
7. Open the current preview draft or a selected range in Gmail or Outlook Web.

## Permission Usage

`storage` is used only for local popup state, including CSV text, templates, generated draft data, and progress.

`tabs` is used only to open compose tabs for Gmail, Outlook Web, or `mailto:`.

Host permissions for Gmail and Outlook are used by the optional Auto-Send content script. The content script only acts when the compose URL contains the extension's `mailMergeAutoSend=true` marker. It attempts to identify the compose window opened by the extension before clicking Send.

## User Data

The extension does not run a backend service and does not transmit CSV data, draft content, or recipient information to a third-party server.

Data is stored locally in Chrome extension storage and can be deleted from the popup with `Clear Saved Data`.

## Auto-Send Safeguards

Auto-Send is off by default and labeled experimental. When enabled, action buttons switch from draft-opening language to explicit `Open & Send...` language. Outlook Auto-Send is blocked when CC or BCC is present because Outlook Web may ignore copy recipients in deeplink compose URLs.

## Test CSV For Review

```csv
Name,Email,Course,Section,DueDate,CcEmail,BccEmail
John,john@example.com,ISOM 210,A,Friday,ta@example.com,archive@example.com
Jane,jane@example.com,ISOM 340,B,Monday,ta@example.com,archive@example.com
```
