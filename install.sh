#!/usr/bin/env bash
# M365 Bridge installer for macOS / Linux.
# Usage: curl -fsSL <repo>/install.sh | bash     or     ./install.sh from a clone
set -euo pipefail
REPO_URL="${M365_BRIDGE_REPO:-https://github.com/clearfield-labs/m365-bridge.git}"
DEST="${M365_BRIDGE_DIR:-$HOME/.m365-bridge-app}"

if ! command -v node >/dev/null 2>&1; then
  echo "Node.js is not installed. Install the LTS version from https://nodejs.org (or: brew install node), then run this again."
  exit 1
fi
NODE_MAJOR=$(node -p 'process.versions.node.split(".")[0]')
if [ "$NODE_MAJOR" -lt 18 ]; then echo "Node.js 18 or newer is needed (you have $(node -v))."; exit 1; fi

if [ -f "$(dirname "$0")/package.json" ] && grep -q '"m365-bridge"' "$(dirname "$0")/package.json" 2>/dev/null; then
  DEST="$(cd "$(dirname "$0")" && pwd)"
elif [ -d "$DEST/.git" ]; then
  git -C "$DEST" pull --ff-only
else
  git clone --depth 1 "$REPO_URL" "$DEST"
fi

cd "$DEST"
npm install --omit=dev --no-fund --no-audit
node bin/m365-bridge.js install
