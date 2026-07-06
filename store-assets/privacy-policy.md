# Privacy Policy

Effective date: July 5, 2026

Mail Merge Draft Helper is designed to create personalized Gmail or Outlook Web email drafts from CSV data in the user's browser.

## Data Collection

Mail Merge Draft Helper does not collect, sell, transmit, or share user data with a third-party server.

The extension processes CSV text, email templates, generated draft content, and progress locally in the browser.

## Local Storage

The extension uses Chrome extension storage to save:

- CSV text pasted or loaded by the user
- Subject, body, CC, and BCC templates
- Generated draft preview data
- Draft-opening progress
- User-selected compose mode and Auto-Send setting

This local storage is used only to let users close and reopen the popup without losing their work.

Users can remove this saved state at any time by clicking `Clear Saved Data` in the extension popup.

## Email Services

Mail Merge Draft Helper opens Gmail, Outlook Web, or `mailto:` compose tabs using the user's selected draft data.

The extension does not access the user's mailbox, read existing emails, or use Gmail API or Microsoft Graph API.

## Auto-Send

Auto-Send is an optional experimental feature. When enabled, the extension content script attempts to identify the compose window opened by the extension and click the matching Send button.

Auto-Send does not transmit data to any third-party server. Users should review generated messages carefully before enabling Auto-Send.

## Permissions

The extension requests:

- `storage`: to save draft-generation state locally.
- `tabs`: to open compose tabs.
- Gmail and Outlook host permissions: to support optional Auto-Send behavior on compose pages.

## Changes

This privacy policy may be updated if the extension's behavior changes.

## Contact

Add your contact email or support URL here before publishing.
