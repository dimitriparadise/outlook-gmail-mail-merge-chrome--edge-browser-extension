# Chrome Web Store Publishing Checklist

## Before uploading

- [ ] `node --test` passes.
- [ ] `node --check popup.js && node --check content.js && node --check i18n.js && node --check xlsx.js`.
- [ ] Version bumped in `manifest.json` and the badge in `popup.html` (and `store-assets/popup-screenshot-harness.html`).
- [ ] `./scripts/build_release.sh` → `dist/mail-merge-draft-helper-<version>.zip`.
- [ ] If the UI changed, regenerate images (see "Images" below).
- [ ] GitHub release created with the same zip.

## Package

- [ ] Developer Dashboard → the item → **Package** → **Upload new package** → the zip.

## Store listing (one per language)

Text and file paths are in [`listing.md`](listing.md).

- [ ] English (default): description, 5 screenshots from `screenshots/en/`, promo tile from `promo/en/`.
- [ ] Add language **Chinese (Simplified)**: description, `screenshots/zh/`, `promo/zh/`.
- [ ] Add language **Spanish**: description, `screenshots/es/`, `promo/es/`.
- [ ] Name and summary come from `_locales/<lang>/messages.json` in the package; check they look right.
- [ ] Category, homepage URL, support URL.

## Privacy practices

- [ ] Single purpose, permission justifications, remote code = No (from `listing.md`).
- [ ] Data usage: no data types ticked; all three certifications ticked.
- [ ] Privacy policy URL points to `store-assets/privacy-policy.md` on GitHub.

## Submit

- [ ] Paste the reviewer notes from [`review-notes.md`](review-notes.md).
- [ ] Submit for review and watch the item status; respond to any policy feedback.
- [ ] After approval, update the README install section if anything changed.

## Images

Both scripts load the built extension (`dist/`) unpacked in Chromium and use the files in `examples/`:

```bash
npm install --no-save playwright && npx playwright install chromium
node scripts/capture_docs_screenshots.mjs   # README screenshots → docs/images/
node scripts/capture_store_assets.mjs       # store screenshots and promo tiles → store-assets/
```

Icons are drawn by `python3 scripts/generate_icons.py` (needs Pillow).
