#!/bin/bash
# Folia deploy. Usage: ./folia/deploy.sh "commit message"  (run from repo root with github credentials)
set -e
cd "$(dirname "$0")/.."
V=$(date -u +%Y%m%d%H%M%S)
sed "s/__VERSION__/$V/" folia/sw.template.js > folia/sw.js
echo "export const BUILD = '$V';" > folia/build.js
git add folia
git -c user.name=shinesdags-bot -c user.email=338882197+shinesdags-bot@users.noreply.github.com commit -q -m "Folia: $1"
git push -q origin main
echo "deployed folia $V $(git log --oneline -1)"
