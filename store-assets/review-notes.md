# Chrome Web Store Review Notes (0.2.2)

Paste the "Notes for the reviewer" block into the Developer Dashboard when submitting.

## Notes for the reviewer

```text
Mail Merge Draft Helper is a Manifest V3 extension that turns an Excel or CSV list into personalized Outlook on the web or Gmail email drafts. It has no backend, no sign-in, and no remote code. All processing happens in the popup; data is kept in chrome.storage.local.

How to test (about 2 minutes, no account or file needed):
1. Click the toolbar icon to open the popup.
2. Click "New here? Try it with sample data". This loads 7 made-up recipients (example.edu / example.net addresses) and a template, builds the drafts, and opens step 3 (Review).
3. Review notes list recipients worth a look (a name that may not match the address, a repeated address, an unusual domain). Use Previous / Next to read the personalized drafts.
4. Click "Next: send". Enter your own email address under "Send yourself a test first" and click "Open Test Draft to Me". A compose tab opens in Outlook on the web (or Gmail, if selected under Service Mode) addressed only to you, with "[TEST]" in the subject. Nothing is sent automatically.
5. "Open Batch 1 of 3" opens one compose tab for each of the first 3 sample recipients (the sample uses batches of 3). Close them without sending; the sample addresses are reserved example domains.

Permissions:
- storage: saves the popup state locally (list, templates, drafts, progress, settings).
- Host permissions for mail.google.com, outlook.office.com, and outlook.live.com: used only by the optional, experimental Auto-Send feature (off by default). When enabled, content.js acts only on compose tabs opened by the extension (marked with mailMergeAutoSend=true in the URL) and clicks that window's Send button. It never reads the mailbox.

Excel files are read locally by xlsx.js using the browser's built-in DecompressionStream; no third-party library is bundled.
```

## Auto-Send safeguards

- Off by default and labeled experimental; importing a plan or loading sample data always turns it off.
- When it is on, every action button says "Send" ("Send Test to Me", "Send Batch 1 of N") instead of "Open".
- In Outlook mode it is blocked when CC or BCC is present, because those drafts open through `mailto:`.
- Users are asked to send themselves a test first; the content script gives up after 15 seconds if it cannot find the matching Send button.

## What changed since the published version (0.1.1)

- Excel (.xlsx) import with a column picker, sample data, a step-by-step flow, review notes, a test send to yourself, batch sending with saved progress, email plan import/export, variable defaults and suggestions, and an English / Chinese / Spanish interface.
- Fix: Auto-Send now finds Outlook on the web's full-page compose window.
- Permissions: `storage` plus the Gmail and Outlook host permissions only. The `tabs` permission is not requested.
