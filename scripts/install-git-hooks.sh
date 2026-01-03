#!/usr/bin/env bash
set -euo pipefail

# Installs a pre-push hook that always syncs .gs -> .js and then runs clasp push.
# This enforces the "always deploy on push" rule locally.

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
HOOK_DIR="$REPO_DIR/.git/hooks"
HOOK_PATH="$HOOK_DIR/pre-push"

mkdir -p "$HOOK_DIR"

cat > "$HOOK_PATH" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

# NOTE: This hook lives in .git/hooks/, so repo root is two levels up.
REPO_DIR="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$REPO_DIR"

echo "==> [pre-push] Syncing .gs -> .js"
bash ./sync-from-gs.sh

echo "==> [pre-push] clasp push"
clasp push

echo "✅ [pre-push] Apps Script up to date"
EOF

chmod +x "$HOOK_PATH"

echo "✅ Installed git hook: $HOOK_PATH"
echo "   From now on, 'git push' will run sync-from-gs.sh and then clasp push."


