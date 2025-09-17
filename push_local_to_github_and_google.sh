#!/usr/bin/env bash
#
# push_local_to_github_and_google.sh
#
# PURPOSE
#   Commit your LOCAL changes, push them to GitHub (origin main),
#   and then push those same local files to Google Apps Script (clasp push).
#
# QUICK START
#   cd ~/Documents/"Genealogy Sheets"
#   chmod +x push_local_to_github_and_google.sh
#   ./push_local_to_github_and_google.sh
#
# USAGE
#   ./push_local_to_github_and_google.sh [-m "Commit message"] [--pull-first]
#
# EXAMPLES
#   ./push_local_to_github_and_google.sh
#   ./push_local_to_github_and_google.sh -m "Implement relink dropdown with IDs"
#   ./push_local_to_github_and_google.sh --pull-first
#   ./push_local_to_github_and_google.sh -m "Refactor date normalization" --pull-first
#
# FLAGS
#   -m|--message   Commit message (default adds an ISO timestamp)
#   --pull-first   Runs `clasp pull` BEFORE committing, so you include any
#                  online edits made in the Apps Script editor.
#
# REQUIREMENTS
#   - Node 20+ with `clasp` installed and logged in
#   - Apps Script API enabled at https://script.google.com/home/usersettings
#   - This folder is a Git repo with `origin` remote configured
#
# WHAT IT DOES
#   1) (optional) clasp pull      # bring down any online edits first
#   2) git add/commit             # include your local changes
#   3) git pull --rebase          # sync with remote main
#   4) git push                   # push to GitHub (origin main)
#   5) clasp push                 # push the final local state to Google
#
# TROUBLESHOOTING
#   - If you edited online and locally at the same time:
#       Use `--pull-first` to pull Google changes before committing.
#   - To inspect differences:
#       clasp status
#   - If Git warns about merge/rebase conflicts:
#       Resolve the files, then:  git add <fixed-file> && git rebase --continue
#
set -euo pipefail

die() { echo "❌ $*" >&2; exit 1; }
need() { command -v "$1" >/dev/null 2>&1 || die "Missing required command: $1"; }

need git
need clasp

COMMIT_MSG="Local → GitHub → Google: $(date -Iseconds)"
PULL_FIRST=0

while [[ $# -gt 0 ]]; do
  case "$1" in
    -m|--message)
      shift
      [[ $# -gt 0 ]] || die "Missing value for -m|--message"
      COMMIT_MSG="$1"
      ;;
    --pull-first)
      PULL_FIRST=1
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

if [[ $PULL_FIRST -eq 1 ]]; then
  echo "⬇️  Pulling latest Apps Script from Google → local (clasp pull)"
  clasp pull
fi

echo "➕ Staging local changes..."
git add -A

if git diff --cached --quiet; then
  echo "✅ No changes to commit."
else
  echo "📝 Committing: ${COMMIT_MSG}"
  git commit -m "${COMMIT_MSG}"
fi

echo "🔄 Sync with remote (git pull --rebase origin main)"
git pull --rebase origin main || true

echo "🚀 Pushing to GitHub (origin main)"
git push -u origin main

echo "⬆️  Pushing local files to Google (clasp push)"
clasp push

echo "🎉 Done."
