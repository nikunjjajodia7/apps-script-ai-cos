#!/usr/bin/env bash
set -euo pipefail

# Deploy pipeline for AI COS Apps Script
# 1) Sync canonical .gs sources into this clasp repo (.js mirror)
# 2) Push to Apps Script via clasp

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$REPO_DIR"

echo "==> Syncing .gs -> .js"
bash ./sync-from-gs.sh

echo "==> clasp push"
clasp push

echo "✅ Deploy complete"


