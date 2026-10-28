#!/usr/bin/env bash
# SPDX-License-Identifier: Apache-2.0
#
# build-site.sh <out-dir> — stellt die Seite für GitHub Pages zusammen:
# site/index.html, das Vorschaubild und das frisch gebaute Beispiel (ohne
# Logo). Kopiert wird nur, was hier steht; check-site.sh meldet tote Verweise.

set -euo pipefail

out="${1:?usage: build-site.sh <out-dir>}"
root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

rm -rf "$out"
mkdir -p "$out"
out="$(cd "$out" && pwd)"
cp "$root/site/index.html" "$root/docs/vorschau.png" "$out/"
node -e 'require(process.argv[1]).build(process.argv[2])' \
  "$root/beispiel/beispiel.js" "$out/beispiel.pptx"

echo "build-site: $(find "$out" -type f | wc -l | tr -d ' ') Dateien in $out"
