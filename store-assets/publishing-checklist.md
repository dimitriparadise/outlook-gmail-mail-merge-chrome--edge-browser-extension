# Chrome Web Store Publishing Checklist

## Account

- [ ] Create or sign in to a Chrome Web Store developer account.
- [ ] Complete developer account registration and payment if required.
- [ ] Confirm publisher name and contact information.

## Extension Package

- [ ] Run `node --test`.
- [ ] Run `node --check popup.js && node --check content.js`.
- [ ] Generate store assets with `python3 scripts/generate_store_assets.py`.
- [ ] Build the release zip.
- [ ] Upload the zip in the Chrome Developer Dashboard.

## Store Listing

- [ ] Name: `Mail Merge Draft Helper`.
- [ ] Short description from `store-assets/listing.md`.
- [ ] Detailed description from `store-assets/listing.md`.
- [ ] Category: Productivity.
- [ ] Language: English.
- [ ] Upload screenshots from `store-assets/screenshots/`.
- [ ] Upload `store-assets/promo/small-promo-tile-440x280.png` if desired.
- [ ] Add support URL.
- [ ] Add website URL if available.

## Privacy

- [ ] Host or publish the privacy policy in `store-assets/privacy-policy.md`.
- [ ] Add the privacy policy URL in the Developer Dashboard if required.
- [ ] Complete privacy practices.
- [ ] Complete limited-use certification.
- [ ] Confirm that the listing accurately states local-only storage behavior.

## Permissions

- [ ] Explain `storage`: local state for CSV, templates, generated drafts, and progress.
- [ ] Explain `tabs`: opens Gmail, Outlook, or mailto compose tabs.
- [ ] Explain Gmail and Outlook host permissions: optional Auto-Send compose-window detection.

## Review

- [ ] Add reviewer notes from `store-assets/review-notes.md`.
- [ ] Submit for review.
- [ ] Monitor review status and respond to any policy feedback.
