# Time to Roll · 抽菸紀錄

一個個人用的抽菸紀錄 PWA。Google 試算表當後台、Google Apps Script 當前後端、iPhone Safari「加入主畫面」當 App。跟 LEDGER 同一套技術與風格（cream + teal、圓角卡片、bottom-sheet）。

- 記錄：一鍵「記一根」，先選菸品（口味）再選原因，時間自動記。
- 原因：使用者自己新增／刪除，不寫死。
- 菸品：分三種菸種（加熱菸、盒菸、捲菸），可設常用預設。
- 之後：庫存連動（P2）、統計與花費（P3）。

## 首次安裝

```bash
npm i -g @google/clasp
clasp login                      # 用要放資料的那個 Google 帳號登入

# 建一個新的試算表 + 綁定的 Apps Script 專案：
clasp create --type sheets --title "Time to Roll" --rootDir src
# ↑ 會自動產生 .clasp.json（含 scriptId、rootDir:"src"）

clasp push                       # 上傳 src/ 的三個檔
```

（`.clasp.json.example` 只是給別人自建時參考用，你自己跑 `clasp create` 會自動產生真正的 `.clasp.json`，它已被 gitignore。）

接著在 Apps Script 編輯器裡：
1. 執行一次 `setup()`（建立「設定／原因／菸品／菸草包」分頁與當月紀錄分頁）。
2. 部署 → 新增部署 → 網頁應用程式 → 執行身分「我」、存取權「僅限我自己」→ 記下 `/exec` 網址與部署 ID。
3. iPhone Safari 開 `/exec` → 分享 → 加入主畫面。

## 每次改程式（clasp v3）

```bash
clasp push
clasp redeploy <你的部署ID>      # v3：redeploy＝更新既有部署；/exec 網址不變
```

- 部署 ID 用 `clasp list-deployments` 看非 `@HEAD` 那筆（`AKfyc…`）。
- ⚠️ v3 的 `clasp deploy`（＝`create-deployment`）會建**新**部署＋新網址，別用來更新。
- 懶人版：`./update.sh`（自動抓非 @HEAD 的部署 ID 並 `redeploy`）。
- 更新後 iPhone 主畫面 App 完全關掉再重開即新版。`git push` 只是備份、與 App 無關。

## 資料模型

見 `CLAUDE.md`。所有資料存在你自己的試算表，跟別人完全隔離。
