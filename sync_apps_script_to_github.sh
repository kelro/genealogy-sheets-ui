#!/usr/bin/env bash
#
# sync_apps_script_to_github.sh
#
# PURPOSE
#   Pull the latest Google Apps Script files (via `clasp pull`) into this folder,
#   then commit and push those changes to GitHub (remote `origin`, branch `main`).
#   Optionally push your local files back to Google with `--push-google`.
#
# QUICK START
#   cd ~/Documents/"Genealogy Sheets"
#   chmod +x sync_apps_script_to_github.sh
#   ./sync_apps_script_to_github.sh
#
# USAGE
#   ./sync_apps_script_to_github.sh [-m "Commit message"] [--push-google]
#
# EXAMPLES
#   ./sync_apps_script_to_github.sh
#   ./sync_apps_script_to_github.sh -m "Sync after editing in Apps Script editor"
#   ./sync_apps_script_to_github.sh --push-google
#   ./sync_apps_script_to_github.sh -m "Daily sync" --push-google
#
# REQUIREMENTS
#   - Node 20+ with `clasp` installed:
#       curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.7/install.sh | bash
#       source ~/.bashrc && nvm install 20 && nvm use 20
#       npm install -g @google/clasp
#   - `clasp login` (run once) and enable Apps Script API:
#       https://script.google.com/home/usersettings  (turn ON)
#   - This folder is a Git repo with a configured `origin` remote.
#
# WHAT IT DOES
#   1) clasp pull                 # downloads latest .gs/.html from Google → local
#   2) git add/commit             # stages and commits any changes
#   3) git pull --rebase          # sync with remote main
#   4) git push                   # push to GitHub (origin main)
#   5) (optional) clasp push      # send local files back to Google (--push-google)
#
# TROUBLESHOOTING
#   - ERROR: 'User has not enabled the Apps Script API'
#       Enable here: https://script.google.com/home/usersettings (wait ~1–2 min)
#   - ERROR: EBADENGINE / Node < 20
#       Use nvm to install Node 20+, then reinstall clasp.
#   - ERROR: EACCES on npm -g install
#       Use nvm; do NOT use sudo for global npm installs.
#   - ERROR: 'src refspec main does not match any'
#       You have no commit on 'main' yet. Run:
#         git add -A && git commit -m "Initial commit" && git push -u origin main
#   - See differences between local and Google:
#         clasp status
#
set -euo pipefail

die() { echo "❌ $*" >&2; exit 1; }
need() { command -v "$1" >/dev/null 2>&1 || die "Missing required command: $1"; }

need git
need clasp

COMMIT_MSG="Sync from Apps Script → GitHub: $(date -Iseconds)"
PUSH_GOOGLE=0

while [[ $# -gt 0 ]]; do
  case "$1" in
    -m|--message)
      shift
      [[ $# -gt 0 ]] || die "Missing value for -m|--message"
      COMMIT_MSG="$1"
      ;;
    --push-google)
      PUSH_GOOGLE=1
      ;;
    -h|--help)
      sed -n '1,200p' "$0"
      exit 0
      ;;
    *)
      die "Unknown argument: $1"
      ;;
  esac
  shift
done

# Ensure repo exists and branch main is current
if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "ℹ️  Initializing new Git repo with branch 'main'..."
  git init -b main
fi

CURRENT_BRANCH=$(git rev-parse --abbrev-ref HEAD)
if [[ "$CURRENT_BRANCH" != "main" ]]; then
  echo "ℹ️  Switching to 'main' branch..."
  if git show-ref --verify --quiet refs/heads/main; then
    git checkout main
  else
    git branch -M main
  fi
fi

if ! git remote get-url origin >/dev/null 2>&1; then
  echo "⚠️  No 'origin' remote configured."
  echo "   Add it (HTTPS example):"
  echo "     git remote add origin https://github.com/<USER>/<REPO>.git"
  echo "   Or create with GitHub CLI:"
  echo "     gh repo create <REPO-NAME> --source=. --public --push"
  die "Add 'origin' remote, then re-run."
fi

echo "⬇️  Pulling latest Apps Script from Google → local (clasp pull)"
clasp pull

echo "➕ Staging any changes..."
git add -A

if git diff --cached --quiet; then
  echo "✅ No changes to commit after clasp pull."
else
  echo "📝 Committing: ${COMMIT_MSG}"
  git commit -m "${COMMIT_MSG}"
fi

echo "🔄 Sync with remote (git pull --rebase origin main)"
git pull --rebase origin main || true

echo "🚀 Pushing to GitHub (origin main)"
git push -u origin main

if [[ $PUSH_GOOGLE -eq 1 ]]; then
  echo "⬆️  Pushing local files back to Google (clasp push)"
  clasp push
fi

echo "🎉 Done."
