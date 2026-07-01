#!/usr/bin/env bash
# 一鍵更新：推程式 → 把正式部署(非 @HEAD)更新到最新版本。/exec 網址不變。
# clasp v3：deploy=create-deployment(建新的)、redeploy=update-deployment(更新既有)
set -e; cd "$(dirname "$0")"
git pull --ff-only 2>/dev/null || true
clasp push -f
DEP="$(clasp list-deployments 2>/dev/null | grep -v '@HEAD' | grep -oE 'AKfyc[A-Za-z0-9_-]+' | head -1 || true)"
[ -n "$DEP" ] && clasp redeploy "$DEP" || { echo "找不到正式部署，先在 Apps Script 部署一次網頁 App"; exit 1; }
