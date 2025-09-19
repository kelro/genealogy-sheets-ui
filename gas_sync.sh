#!/usr/bin/env bash
set -euo pipefail

MSG="${1:-Sync Apps Script}"

need() { command -v "$1" >/dev/null 2>&1 || { echo "ERROR: '$1' not found."; exit 1; }; }
need node
need git
need clasp

if [[ ! -f .clasp.json ]]; then
  echo "ERROR: .clasp.json not found in $(pwd)."
  echo "Tip: run 'clasp clone <SCRIPT_ID>' or create it with 'clasp init'."
  exit 1
fi

# Read config
SCRIPTID=$(node -e 'try{console.log(require("./.clasp.json").scriptId || "");}catch(e){console.log("");}')
ROOTDIR=$(node -e 'try{console.log(require("./.clasp.json").rootDir || "");}catch(e){console.log("");}')

if [[ -z "$SCRIPTID" ]]; then
  echo "ERROR: scriptId missing in .clasp.json."
  echo "Make sure this is the Apps Script *Script ID* (Project Settings → Script ID), not the Cloud *Project ID*."
  exit 1
fi

echo "scriptId: $SCRIPTID"
[[ -n "$ROOTDIR" ]] && echo "rootDir : $ROOTDIR" || echo "rootDir : (repo root)"
echo "clasp   : $(clasp --version || true)"

# Determine if 'clasp pull' supports --force (varies by version)
PULL_CMD=(clasp pull)
if clasp pull --help 2>&1 | grep -q -- '--force'; then
  PULL_CMD=(clasp pull --force)
fi

open_editor_url() {
  local url="https://script.google.com/home/projects/${SCRIPTID}/edit"
  if command -v xdg-open >/dev/null 2>&1; then
    xdg-open "$url" >/dev/null 2>&1 || true
  elif command -v open >/dev/null 2>&1; then
    open "$url" >/dev/null 2>&1 || true
  else
    echo "Open in browser if needed: $url"
  fi
}

clear_clasp_tokens() {
  # Safe to remove; clasp will recreate on login
  rm -f ~/.clasprc.json || true
  rm -f ~/.config/configstore/@google-clasp.json || true
}

relogin_and_retry_pull() {
  echo "First pull failed — attempting auth recovery…"
  clasp logout || true
  clear_clasp_tokens
  echo "Re-login required. A browser window will open (or follow terminal prompts)."
  if ! clasp login; then
    echo "Standard login failed; trying '--no-localhost' flow…"
    clasp login --no-localhost
  fi
  echo "Retrying pull…"
  "${PULL_CMD[@]}"
}

# Open the correct project in your browser (helps spot ID mix-ups)
open_editor_url

echo "Pulling latest from Apps Script…"
if ! "${PULL_CMD[@]}"; then
  # Handle common auth cache error: "Error retrieving access token..."
  relogin_and_retry_pull
fi

# Create a sane claspignore if missing (prefer .gs/.html only)
if [[ ! -f .claspignore ]]; then
  cat > .claspignore <<'EOF'
**
!appsscript.json
!*.gs
!*.html
EOF
  echo "Created .claspignore (syncs .gs and .html only)."
fi

# Show where files landed (root or rootDir)
LIST_DIR="${ROOTDIR:-.}"
echo "Listing ${LIST_DIR}:"
ls -la "$LIST_DIR"

# Commit & push to GitHub
git add -A
git commit -m "$MSG" || echo "Nothing to commit."
git push || echo "Git push skipped/failed."

echo "Done."

