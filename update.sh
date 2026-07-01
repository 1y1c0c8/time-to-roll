#!/usr/bin/env bash
set -e; cd "$(dirname "$0")"
git pull --ff-only
clasp push -f
DEP="$(clasp deployments | grep -v '@HEAD' | grep -oE 'AKfyc[A-Za-z0-9_-]+' | head -1 || true)"
[ -n "$DEP" ] && clasp deploy -i "$DEP" || { echo "先在 Apps Script 部署一次網頁 App"; exit 1; }
