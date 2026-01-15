#!/usr/bin/env bash
set -euo pipefail

# Usage:
#   ./scripts/sync.sh "your commit message"
# If you don't pass a message, it uses a default.

MSG="${1:-sync: push to Apps Script and Git}"

# Make sure we're in a clasp project folder
if [ ! -f ".clasp.json" ]; then
  echo "Error: .clasp.json not found. Run this inside the folder that is linked to an Apps Script project."
  exit 1
fi

echo "==> Pushing to Google Apps Script..."
clasp push

echo "==> Committing + pushing to Git..."
git add -A

# Only commit if there are changes
if git diff --cached --quiet; then
  echo "No git changes to commit."
else
  git commit -m "$MSG"
  git push
fi

echo "✅ Done."
