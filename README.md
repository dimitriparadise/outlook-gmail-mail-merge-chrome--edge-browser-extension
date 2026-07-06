# Mail Merge Draft Helper

Mail Merge Draft Helper is a lightweight Chrome extension for creating personalized Gmail or Outlook Web email drafts from a CSV file. It is designed for instructors, administrators, and small teams who need to prepare many similar messages while still reviewing each email before it goes out.

The extension runs locally in the browser popup. You paste or upload a CSV, write reusable subject and body templates, preview each personalized message, and open the selected draft in Gmail or Outlook Web.

## Highlights

- Create personalized Gmail or Outlook Web drafts from CSV rows.
- Use template variables such as `{{Name}}`, `{{Course}}`, or any other CSV column.
- Preview each generated message before opening it.
- Open the current preview draft or open a selected range of drafts.
- Add optional CC and BCC values, including values from CSV columns.
- Get preflight warnings for missing variables, invalid email addresses, duplicate recipients, and To/CC/BCC overlap.
- Keep progress locally with `chrome.storage.local`.
- Clear saved local data at any time.
- Optional experimental Auto-Send mode with clearer "Open & Send..." button labels.

## How It Works

1. Upload a CSV file or paste CSV text into the popup.
2. Choose Gmail Mode or Outlook Mode.
3. Customize the subject, body, CC, and BCC templates.
4. Click `Generate Drafts`.
5. Review each personalized preview.
6. Click `Open Preview Draft`, or choose a From/To range and click `Open Selected Drafts`.

When Auto-Send is enabled, the single-draft button changes to `Open & Send Preview Email`, and the range button changes to `Open & Send Selected Emails`.

## CSV Format

Your CSV must include a header row and at least one data row. It must include either `Email` or `email`.

```csv
Name,Email,Course,Section,DueDate,CcEmail,BccEmail
John,john@example.com,ISOM 210,A,Friday,ta@example.com,archive@example.com
Jane,jane@example.com,ISOM 340,B,Monday,ta@example.com,archive@example.com
```

Any header can be used as a template variable:

```text
Subject:
Reminder for {{Course}}

Body:
Hi {{Name}},

This is a quick reminder about {{Course}} section {{Section}}, due {{DueDate}}.

CC:
{{CcEmail}}

BCC:
{{BccEmail}}
```

Use a real line break between the header row and the first data row. Visual wrapping inside the text box does not count as a CSV row break.

## Gmail And Outlook Modes

`Gmail Mode` opens Gmail compose URLs directly and supports To, CC, BCC, subject, and body.

`Outlook Mode` opens Outlook Web compose links. If a generated message has CC or BCC recipients, the extension uses `mailto:` because Outlook Web compose links may ignore copy recipients in some tenants.

Auto-Send cannot be used with Outlook Mode when CC or BCC is present. Use Gmail Mode or turn off Auto-Send for those messages.

## Privacy

Mail Merge Draft Helper does not run a backend service and does not send your CSV data to a third-party server.

The extension stores CSV text, templates, generated draft data, and progress locally in Chrome extension storage so you can close and reopen the popup without losing your place. Use `Clear Saved Data` to remove this local saved state.

The extension requests:

- `storage`: saves your local draft-generation state.
- `tabs`: opens Gmail, Outlook, or `mailto:` compose tabs.
- Gmail and Outlook host permissions: allows the optional Auto-Send content script to identify the compose window opened by the extension.

## Limitations

- Attachments are not supported in the current URL-based compose flow.
- There is no rich-text email editor.
- Very large ranges may still be limited by browser tab or popup throttling.
- Auto-Send is experimental. Review your drafts carefully before using it.

## Local Development

1. Open Chrome or Edge.
2. Go to `chrome://extensions` or `edge://extensions`.
3. Enable Developer mode.
4. Click `Load unpacked`.
5. Select this folder.
6. Click the extension icon to open the popup.

Run the test suite with:

```bash
node --test
```

Run a syntax check with:

```bash
node --check popup.js && node --check content.js
```

## Publishing To The Chrome Web Store

Before publishing, prepare the following:

- A production-ready zip package of the extension files.
- Store listing name, summary, detailed description, category, language, and screenshots.
- A 128x128 extension icon and promotional images if you want a stronger listing page.
- A privacy policy URL if required for your data handling disclosures.
- Clear permission justifications for `storage`, `tabs`, and host permissions.
- Completed privacy practices and limited-use certification in the Chrome Developer Dashboard.

Publishing flow:

1. Create or sign in to a Chrome Web Store developer account.
2. Open the [Chrome Developer Dashboard](https://chrome.google.com/webstore/devconsole/).
3. Click `Add new item`.
4. Upload the extension zip file.
5. Complete the store listing fields.
6. Complete the privacy practices fields.
7. Set distribution, visibility, and regions.
8. Submit the item for Chrome Web Store review.

Official references:

- [Publish in the Chrome Web Store](https://developer.chrome.com/docs/webstore/publish)
- [Chrome Web Store Developer Program Policies](https://developer.chrome.com/docs/webstore/program-policies)
- [Privacy disclosure requirements](https://developer.chrome.com/docs/webstore/program-policies/user-data-faq)
