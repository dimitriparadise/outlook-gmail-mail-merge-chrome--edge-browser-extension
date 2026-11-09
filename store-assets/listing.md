# Chrome Web Store Listing Draft

## Name

Mail Merge Draft Helper

## Short Description

Create personalized Gmail or Outlook drafts from CSV lists with previews and local-only state.

## Detailed Description

Mail Merge Draft Helper helps you prepare personalized email drafts from a CSV list directly in Chrome. It is designed for instructors, administrators, coordinators, and small teams who need to send many similar messages while still reviewing each email before it goes out.

Load an Excel or CSV list, write one reusable subject and body template, and use variables such as `{{Name}}`, `{{Course}}`, or any other CSV column. The extension generates personalized drafts, lets you preview each message, and opens the current preview or a selected range in Gmail or Outlook Web.

Key features:

- Load an Excel (.xlsx) or CSV list and choose which columns to use; everything else is left out.
- Generate personalized Gmail or Outlook Web drafts from each row.
- Use any column header as a template variable.
- Review notes flag names that may not match the email address, unusual email domains, and empty fields, with inline fixes.
- Send yourself a test copy before mailing the list.
- Send in batches (50 per batch by default) with remembered progress.
- Save or load the whole setup as an email plan file.
- Interface in English, Chinese, and Spanish.
- Preview every generated email before opening it.
- Open the current preview draft or a selected range of drafts.
- Add optional CC and BCC fields, including values from CSV columns.
- Catch common issues with preflight warnings for missing variables, invalid email addresses, duplicate recipients, and To/CC/BCC overlap.
- Save progress locally in Chrome extension storage.
- Clear saved local data at any time.
- Optional experimental Auto-Send mode with explicit Open & Send button labels.

Privacy-friendly by design:

Mail Merge Draft Helper does not run a backend service and does not send your CSV data to a third-party server. CSV text, templates, generated draft data, and progress are stored locally in Chrome extension storage so you can resume your work after closing the popup.

Current limitations:

- Attachments are not supported in the current URL-based compose flow.
- Rich-text editing is not supported.
- Very large ranges may still be limited by browser tab throttling.
- Auto-Send is experimental and should be used only after careful review.

## Category

Productivity

## Language

English

## Support / Website

https://github.com/dimitriparadise/outlook-gmail-mail-merge-chrome--edge-browser-extension/issues

## Suggested Keywords

mail merge, Gmail drafts, Outlook drafts, CSV email, email templates, productivity, teacher tools, bulk email drafts

## Permission Justifications

### storage

Used to save CSV text, templates, generated draft data, and progress locally in Chrome extension storage so users can close and reopen the popup without losing their work.

### Host permissions for Gmail and Outlook

Used by the optional Auto-Send content script to identify the compose window opened by the extension and click the matching Send button only when Auto-Send is enabled.

## Single Purpose Statement

Mail Merge Draft Helper creates personalized Gmail or Outlook Web email drafts from CSV data and helps users review, open, and optionally send those generated messages.
