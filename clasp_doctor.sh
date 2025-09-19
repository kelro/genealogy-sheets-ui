#!/usr/bin/env bash
set -euo pipefail

echo "== PWD =="
pwd

echo "== Checking .clasp.json exists =="
if [[ ! -f .clasp.json ]]; then
  echo "ERROR: .clasp.json not found here. cd into your Apps Script repo directory."
  exit 1
fi

echo "== .clasp.json =="
cat .clasp.json || true

ROOTDIR=$(node -e 'try{console.log(require("./.clasp.json").rootDir || "")}catch(e){console.log("")}')
SCRIPTID=$(node -e 'try{console.log(require("./.clasp.json").scriptId || "")}catch(e){console.log("")}')

if [[ -z "$SCRIPTID" ]]; then
  echo "ERROR: scriptId missing in .clasp.json"
  exit 1
fi

echo "== clasp login status =="
clasp login --status || true

echo "== Remote files (clasp files) =="
clasp files || true

echo "== Status before pull =="
clasp status || true

echo "== Forcing pull =="
clasp pull --force

echo "== Status after pull =="
clasp status || true

if [[ -n "$ROOTDIR" ]]; then
  echo "== Listing $ROOTDIR after pull =="
  ls -la "$ROOTDIR"
else
  echo "== Listing repo root after pull =="
  ls -la
fi

echo "== Grep for listAllFormulas.gs locally =="
grep -R --line-number --include="*.gs" "function listAllFormulas" "${ROOTDIR:-.}" || echo "Not found locally."

echo "== Next steps: add/commit/push (only if file now exists locally) =="
echo "git add -A && git commit -m \"Sync Apps Script\" || echo \"Nothing to commit.\""
echo "git push"
