#!/usr/bin/env bash
# Builds the Chrome Web Store upload: dist/teams-excavator-<version>.zip, with
# manifest.json at the root and only the files the extension loads.
# Packs the committed state (HEAD) through git archive, so it refuses to run
# while packaged files have uncommitted changes. Pass --force to pack HEAD anyway.
set -e
cd "$(dirname "$0")"
PATHS="manifest.json service-worker.js src _locales icons"
if [ "$1" != "--force" ] && [ -n "$(git status --porcelain -- $PATHS)" ]; then
    echo "Uncommitted changes in packaged files (commit them, or pass --force to pack HEAD):" >&2
    git status --short -- $PATHS >&2
    exit 1
fi
VERSION=$(git show HEAD:manifest.json | sed -n 's/^ *"version": *"\([^"]*\)".*/\1/p')
[ -n "$VERSION" ] || { echo "No version found in manifest.json" >&2; exit 1; }
OUT="dist/teams-excavator-$VERSION.zip"
mkdir -p dist
git archive --format=zip -9 -o "$OUT" HEAD $PATHS
echo "$OUT"
