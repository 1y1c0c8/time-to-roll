# Time to Roll — 開發脈絡（給 Claude / 未來的自己）

抽菸紀錄 PWA。Stack：Google Sheet（DB）+ GAS（`src/Code.js`）+ HtmlService 單檔前端（`src/Index.html`）。iPhone 主畫面 PWA。與 LEDGER 同套風格與部署流程。命名「Time to Roll」取自國蛋同名曲；剩最後一支的彩蛋文案「盒仔內賰尾枝」取自阿跨面歌詞。

## 分階段
- **P1（已完成，可用）**：記錄。原因自訂 + 菸品清單（名稱／類別／常用預設）+ 記一根（先選菸品口味再選原因，時間自動）+ 今日紀錄列表（點→編輯／刪除）。設定頁管理原因與菸品。統計／庫存為佔位。
- **P2（已完成，可用）**：庫存。加熱菸/盒菸有「剩餘支數」，買入輸入盒數 → +（盒×每盒支數），記一根自動 −1，顯示「X 盒 Y 支」，某口味整體剩 1 支時卡片跳彩蛋「盒仔內賰尾枝」（襯線斜體 --warm）。捲菸有「未開包數」，買入 +包數；「開新包」建立使用中的菸草包（未開 −1），記一根（捲菸）要選一包 → 該包已捲 +1，「這包用完」歸檔記用完日；歷史顯示「共捲 N 支、每支成本 = 售價/N」。可同時多包使用中。開新包按鈕文案「Time to Roll」。刪除/編輯紀錄會**回沖庫存**（restore/consume）。addSmoke/updateSmoke/deleteSmoke/buyStock/openPouch/finishPouch 都回傳整包狀態 {records?,products,pouches}，前端 applyState() 一次刷新。
- **P3**：統計。頻率趨勢（日／週／月／年）+ 原因分佈 + 時段分佈 + 花費估算。

## 資料模型（試算表分頁）
- `設定`：key-value（時區、幣別、預設每盒支數）。讀 A:B。
- `原因`：名稱 | 排序 | 啟用。刪除＝啟用設 FALSE（軟刪，保留歷史）。
- `菸品`：id | 類別 | 名稱 | 每盒支數 | 售價 | 售價單位 | 剩餘支數 | 常用預設 | 狀態 | 建立時間 | 未開包數。捲菸的每盒支數留空；剩餘支數給加熱菸/盒菸、未開包數給捲菸。刪除＝狀態設「停用」。（未開包數為 P2 新增第 11 欄，`ensureReady_→migrate_()` 會幫既有分頁補表頭。）
- `菸草包`（P2）：id | 菸品id | 口味 | 開封日 | 用完日 | 已捲支數 | 售價 | 狀態。
- `紀錄`：**一個月一分頁**，分頁名 `yyyy-MM`。欄位 id | 時間 | 原因 | 菸品id | 菸品名稱 | 類別 | 菸草包id | 成本 | 備註。原因與菸品名稱都存**快照**，之後改名／刪除不影響歷史。

## 固定欄位索引
`Code.js` 用 `P_*` / `R_*` / `RS_*` 常數定位欄，**絕不用 `HEADERS.length`**。加欄時排在尾端、更新常數即可。菸品／菸草包／紀錄的分頁一開始就建好 P2/P3 全部欄位，P1 只是留白。

## 後端函式（Code.js）
- `doGet()` / `setup()`（首次在編輯器跑一次）
- `getBootData()` → {settings, reasons, products, records(當月), thisMonth}
- 記錄：`addSmoke({productId,reason,pouchId?})`、`updateSmoke({id,month,reason?,productId?,pouchId?,timeMillis?})`、`deleteSmoke({id,month})`、`getMonthRecords(tab)`。捲菸必帶 pouchId；三者都連動庫存並回傳 `state_()`。
- 原因：`addReason(name)`、`deleteReason(name)`（軟）、`renameReason(old,new)`
- 菸品：`addProduct(p)`、`updateProduct(p)`、`deleteProduct(id)`（軟）、`setDefaultProduct(id)`
- 庫存（P2）：`buyStock({id,qty})`、`openPouch({productId})`、`finishPouch({pouchId})`、`getPouches()`→{using,done}。庫存連動 helper：`consumeInventory_/restoreInventory_/adjustProductLeft_/adjustPouchRolled_`。

前端只透過 `google.script.run` 溝通；**不可用 localStorage/sessionStorage**（GAS iframe 沙盒）。

## Gotchas（沿用 LEDGER 血淚）
- 日期樣字串寫進儲存格會被 Sheets 轉成 Date；要存字串旗標時 `setNumberFormat('@')` 並讀取時防禦。
- 建 Date 用當地正午避免時區跨日；輸出用 `Utilities.formatDate(d, TZ, …)`。
- 重用同一個部署 ID（`clasp deploy -i …`）→ `/exec` 網址不變。
- 私有 GAS web app 沒辦法設主畫面自訂圖示（沙盒 iframe），iOS 會用標題首字母字塊。

## 部署 = 雙推
`clasp push` → `clasp deploy -i <id>`（更新 Google web app）**且** `git push`（GitHub）。或 `./update.sh`。
