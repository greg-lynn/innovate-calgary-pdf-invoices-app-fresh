#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${ROOT_DIR}"

mkdir -p artifacts

ZIP_NAME="rocketlane-invoice-access-manager-rli-app.zip"
ZIP_PATH="artifacts/${ZIP_NAME}"
LEGACY_ZIP_PATH="artifacts/rocketlane-invoice-access-manager-app.zip"

STAGING_DIR="$(mktemp -d)"
cleanup() {
  rm -rf "${STAGING_DIR}"
}
trap cleanup EXIT

cp -R dist "${STAGING_DIR}/dist"
cp index.js "${STAGING_DIR}/index.js"
cp package.json "${STAGING_DIR}/package.json"
cp README.md "${STAGING_DIR}/README.md"
if [[ -f package-lock.json ]]; then
  cp package-lock.json "${STAGING_DIR}/package-lock.json"
fi
if [[ -d scripts ]]; then
  cp -R scripts "${STAGING_DIR}/scripts"
fi
if [[ -d server-actions ]]; then
  cp -R server-actions "${STAGING_DIR}/server-actions"
fi

(
  cd "${STAGING_DIR}"
  npx -y @rocketlane/rli@latest build >/dev/null

  # Some installer paths expect deploy.json at zip root.
  if [[ -f rli-dist/deploy.json ]]; then
    cp rli-dist/deploy.json deploy.json
    zip -q -u app.zip deploy.json
  fi
)

rm -f "${ZIP_PATH}" "${LEGACY_ZIP_PATH}"
cp "${STAGING_DIR}/app.zip" "${ZIP_PATH}"
cp "${ZIP_PATH}" "${LEGACY_ZIP_PATH}"

echo "Created ${ZIP_PATH}"
echo "Created ${LEGACY_ZIP_PATH}"
