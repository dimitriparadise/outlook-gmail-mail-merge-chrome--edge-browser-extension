#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DIST_DIR="$ROOT/dist"
PACKAGE_DIR="$DIST_DIR/mail-merge-draft-helper"
ZIP_PATH="$DIST_DIR/mail-merge-draft-helper-0.1.0.zip"

rm -rf "$PACKAGE_DIR"
mkdir -p "$PACKAGE_DIR/icons" "$DIST_DIR"

cp "$ROOT/manifest.json" "$PACKAGE_DIR/"
cp "$ROOT/popup.html" "$PACKAGE_DIR/"
cp "$ROOT/popup.css" "$PACKAGE_DIR/"
cp "$ROOT/popup.js" "$PACKAGE_DIR/"
cp "$ROOT/content.js" "$PACKAGE_DIR/"
cp "$ROOT/icons"/icon-*.png "$PACKAGE_DIR/icons/"

rm -f "$ZIP_PATH"
(cd "$PACKAGE_DIR" && zip -qr "$ZIP_PATH" .)

echo "$ZIP_PATH"
