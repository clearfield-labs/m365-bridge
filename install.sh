#!/usr/bin/env bash
# M365 Bridge installer for macOS and Linux. Nothing to install first, no admin password:
# it downloads the code as a zip (no Git), and a private copy of Node.js if this computer
# has none (no Homebrew, no system install). Run it again to update.
#   curl -fsSL https://raw.githubusercontent.com/clearfield-labs/m365-bridge/main/install.sh | bash
set -euo pipefail

REPO="${M365_BRIDGE_REPO_SLUG:-clearfield-labs/m365-bridge}"
REF="${M365_BRIDGE_REF:-main}"
DEST="${M365_BRIDGE_DIR:-$HOME/.m365-bridge-app}"
NODE_DIR="${M365_BRIDGE_NODE_DIR:-$HOME/.m365-bridge-node}"

say() { printf '\033[1m%s\033[0m\n' "$*"; }
fail() { printf '\n\033[31m%s\033[0m\n' "$*" >&2; exit 1; }

case "$(uname -s)" in Darwin) OS=darwin ;; Linux) OS=linux ;; *) fail "This installer is for Mac and Linux. On Windows, use the PowerShell command." ;; esac
case "$(uname -m)" in arm64|aarch64) ARCH=arm64 ;; x86_64|amd64) ARCH=x64 ;; *) fail "Unsupported processor: $(uname -m)" ;; esac

say "Installing M365 Bridge…"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

# 1. Node.js: use this computer's if it is 18 or newer, otherwise a private copy.
NODE=""
if [ -z "${M365_FORCE_PRIVATE_NODE:-}" ] && command -v node >/dev/null 2>&1; then
  if [ "$(node -p 'process.versions.node.split(".")[0]' 2>/dev/null || echo 0)" -ge 18 ]; then NODE="$(command -v node)"; fi
fi
if [ -z "$NODE" ]; then
  if [ ! -x "$NODE_DIR/bin/node" ] || [ -n "${M365_FORCE_PRIVATE_NODE:-}" ]; then
    say "Downloading Node.js (only for M365 Bridge)…"
    EXT=tar.gz; [ "$OS" = linux ] && EXT=tar.xz
    FILE="$(curl -fsSL https://nodejs.org/dist/latest-v22.x/SHASUMS256.txt | awk '{print $2}' | grep -E "^node-v[0-9.]+-$OS-$ARCH\.$EXT$" | head -1)"
    [ -n "$FILE" ] || fail "Could not find a Node.js download for $OS-$ARCH."
    curl -fsSL "https://nodejs.org/dist/latest-v22.x/$FILE" -o "$TMP/$FILE"
    curl -fsSL https://nodejs.org/dist/latest-v22.x/SHASUMS256.txt | grep " $FILE\$" > "$TMP/sum"
    (cd "$TMP" && { shasum -a 256 -c sum >/dev/null 2>&1 || sha256sum -c sum >/dev/null 2>&1; }) || fail "The Node.js download did not match its checksum. Try again."
    rm -rf "$NODE_DIR" && mkdir -p "$NODE_DIR"
    tar -xf "$TMP/$FILE" -C "$NODE_DIR" --strip-components=1
  fi
  NODE="$NODE_DIR/bin/node"
fi
export PATH="$(dirname "$NODE"):$PATH"

# 2. The code, as a zip from GitHub (no Git needed).
SCRIPT="${BASH_SOURCE[0]:-}"
if [ -n "$SCRIPT" ] && [ -f "$SCRIPT" ] && grep -q '"m365-bridge"' "$(dirname "$SCRIPT")/package.json" 2>/dev/null; then
  DEST="$(cd "$(dirname "$SCRIPT")" && pwd)"   # running ./install.sh from a checkout
else
  say "Downloading M365 Bridge…"
  curl -fsSL "https://codeload.github.com/$REPO/tar.gz/$REF" -o "$TMP/app.tgz"
  mkdir -p "$TMP/app" && tar -xzf "$TMP/app.tgz" -C "$TMP/app" --strip-components=1
  rm -rf "$DEST" && mkdir -p "$(dirname "$DEST")" && mv "$TMP/app" "$DEST"
fi

# 3. Dependencies, then set up: background service, Claude, the M365 Bridge app.
cd "$DEST"
"$NODE" "$(dirname "$NODE")/../lib/node_modules/npm/bin/npm-cli.js" install --omit=dev --no-fund --no-audit --loglevel=error 2>/dev/null \
  || npm install --omit=dev --no-fund --no-audit --loglevel=error
if [ -n "${M365_INSTALL_SKIP_RUN:-}" ]; then "$NODE" bin/m365-bridge.js status >/dev/null && say "Installed (setup skipped)."; exit 0; fi
"$NODE" bin/m365-bridge.js install
