#!/usr/bin/env bash
# Sideload the Michael Outlook add-in manifest into Outlook for Mac.
#
# 2025 note: recent Outlook for Mac builds removed the "Get Add-ins -> Upload"
# UI path, so manifests must be placed in Outlook's `wef` documents folder.
# See: https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac
#
# Usage:
#   scripts/sideload-mac.sh           # dev:  install local manifest.xml (https://localhost:3000)
#   scripts/sideload-mac.sh prod      # prod: install deployed manifest from alansynn.com/michael
set -euo pipefail

MODE="${1:-dev}"
PROD_MANIFEST_URL="https://alansynn.com/michael/manifest.prod.xml"
REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DEV_MANIFEST="$REPO_ROOT/manifest.xml"
TMP_DIR="$(mktemp -d)"
trap 'rm -rf "$TMP_DIR"' EXIT

# Outlook for Mac scans these folders for sideloaded manifests:
#   classic: com.microsoft.Outlook
#   new UI:  com.Microsoft.OlkSSL
WEF_CANDIDATES=(
  "$HOME/Library/Containers/com.microsoft.Outlook/Data/Documents/wef"
  "$HOME/Library/Containers/com.Microsoft.OlkSSL/data/Documents/wef"
)

manifest=""
label=""
case "$MODE" in
  dev)
    if [[ ! -f "$DEV_MANIFEST" ]]; then
      echo "✗ Local manifest not found: $DEV_MANIFEST" >&2
      exit 1
    fi
    manifest="$DEV_MANIFEST"
    label="dev (https://localhost:3000) — run 'npm run dev-server' too"
    ;;
  prod)
    echo "↓ Downloading prod manifest from $PROD_MANIFEST_URL"
    manifest="$TMP_DIR/michael.xml"
    if ! curl -fsSL "$PROD_MANIFEST_URL" -o "$manifest"; then
      echo "✗ Failed to download prod manifest." >&2
      exit 1
    fi
    label="prod (https://alansynn.com/michael)"
    ;;
  *)
    echo "Usage: $0 [dev|prod]" >&2
    exit 2
    ;;
esac

installed=0
for wef in "${WEF_CANDIDATES[@]}"; do
  docs_dir="$(dirname "$wef")"
  if [[ -d "$docs_dir" ]]; then
    mkdir -p "$wef"
    cp "$manifest" "$wef/michael.xml"
    echo "✓ Installed $label"
    echo "  → $wef/michael.xml"
    installed=$((installed + 1))
  fi
done

if [[ "$installed" -eq 0 ]]; then
  echo "✗ No Outlook for Mac data container found." >&2
  echo "  Open Outlook at least once so the container is created, then re-run." >&2
  exit 1
fi

echo ""
echo "Next: fully quit Outlook (Cmd+Q), reopen it, open an email, then click the Michael button."
