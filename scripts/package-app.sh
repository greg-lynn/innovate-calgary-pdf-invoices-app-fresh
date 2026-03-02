#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${ROOT_DIR}"

mkdir -p artifacts

ZIP_NAME="rocketlane-invoice-access-manager-app.zip"
ZIP_PATH="artifacts/${ZIP_NAME}"

rm -f "${ZIP_PATH}"

# Minimal payload Rocketlane needs: manifest + entrypoint assets.
zip -r "${ZIP_PATH}" index.js dist package.json README.md >/dev/null

echo "Created ${ZIP_PATH}"
