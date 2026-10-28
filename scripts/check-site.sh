#!/usr/bin/env bash
# SPDX-License-Identifier: Apache-2.0
#
# check-site.sh [site-dir] — prüft die gebaute Seite (Standard: _site):
#
#   - keine Ressource von fremden Servern (Schriften, Skripte, Styles, Bilder)
#   - lokale Verweise sind relativ (die Seite liegt unter /did-theme-template/)
#     und zeigen auf eine mitgelieferte Datei
#   - kein Logo-Bild in der Ausgabe (das DIDAY-Logo darf nicht verbreitet werden)
#
# Exit 0 = sauber, 1 = Befund (jede Zeile nennt Datei und Verweis).

set -euo pipefail

dir="${1:-_site}"
fails=0
bad() { printf 'check-site: %s\n' "$*" >&2; fails=$((fails + 1)); }

[ -d "$dir" ] || { echo "check-site: $dir fehlt — erst scripts/build-site.sh $dir" >&2; exit 1; }
for f in index.html vorschau.png beispiel.pptx; do
  [ -s "$dir/$f" ] || bad "$f fehlt oder ist leer"
done

page="$dir/index.html"
while IFS= read -r hit; do bad "index.html: lädt von außen: $hit"; done < <(
  grep -noE '<(script|img|iframe|source|audio|video|embed)[[:space:]][^>]*src="(https?:)?//[^"]*"' "$page" || true
  grep -noE '<link[[:space:]][^>]*>' "$page" | grep -E 'href="(https?:)?//' | grep -vE 'rel="canonical"' || true
  grep -noE "url\([[:space:]]*[\"']?(https?:)?//[^)]*\)|@import" "$page" || true
)

while IFS= read -r ref; do
  case "$ref" in
    ''|'#'*|data:*|mailto:*|http://*|https://*|//*) continue ;;
    /*) bad "index.html: $ref ist absolut — auf GitHub Pages liegt die Seite unter /did-theme-template/" ;;
    *) [ -e "$dir/${ref%%[?#]*}" ] || bad "index.html: Verweis auf $ref — Datei fehlt in $dir" ;;
  esac
done < <(grep -oE '[[:space:]](src|href)="[^"]*"' "$page" | sed -E 's/^[[:space:]]*[a-z]+="(.*)"$/\1/')

# Das Beispiel wird ohne Logo gebaut; ein Bild in der PPTX hieße, dass doch eins drin ist.
if ! inhalt="$(unzip -Z1 "$dir/beispiel.pptx" 2>&1)"; then
  bad "beispiel.pptx lässt sich nicht lesen: $inhalt"
elif grep -qE '^ppt/media/.+[^/]$' <<<"$inhalt"; then
  bad "beispiel.pptx enthält Bilder unter ppt/media/ — das Beispiel muss ohne Logo gebaut sein"
fi

if [ "$fails" -gt 0 ]; then
  echo "check-site: $fails Befund(e) in $dir" >&2
  exit 1
fi
echo "check-site: $dir sauber"
