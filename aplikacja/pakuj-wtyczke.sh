#!/usr/bin/env bash
# Buduje aplikację i pakuje ją jako wtyczkę WordPress (ZIP do wgrania na guliwera4.pl).
set -euo pipefail
cd "$(dirname "$0")"
npm run build
OUT="${1:-../wtyczki}"
TMP="$(mktemp -d)"
cp -r ../wtyczki/guliwera-akademia "$TMP/"
rm -rf "$TMP/guliwera-akademia/app"
cp -r dist "$TMP/guliwera-akademia/app"
(cd "$TMP" && rm -f guliwera-akademia.zip && zip -qr guliwera-akademia.zip guliwera-akademia)
mv "$TMP/guliwera-akademia.zip" "$OUT/guliwera-akademia.zip"
rm -rf "$TMP"
echo "Gotowe: $OUT/guliwera-akademia.zip"
