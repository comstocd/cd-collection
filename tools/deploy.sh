#!/bin/sh
# Rebuild collection.js from the workbook and bump BUILD so installed copies show "Update available".
# Usage: tools/deploy.sh "/path/to/CD_Catalogue.xlsx"
set -e
cd "$(dirname "$0")/.."
B=$(( $(sed -n 's/^var BUILD = \([0-9]*\);/\1/p' sw.js) + 1 ))
sed -i.bak "s/^var BUILD = [0-9]*;/var BUILD = $B;/" sw.js && rm -f sw.js.bak
python3 tools/build.py "$1" collection.js "$B"
git add -A && git commit -m "Build $B" && git push
