# Time to Roll — 開發脈絡（給 Claude / 未來的自己）

抽菸紀錄 PWA。Stack：Google Sheet（DB）+ GAS（`src/Code.js`）+ HtmlService 單檔前端（`src/Index.html`）。iPhone 主畫面 PWA。與 LEDGER 同套風格與部署流程。命名「Time to Roll」取自國蛋同名曲；剩最後一支的彩蛋文案「盒仔內賰尾枝」取自阿跨面歌詞。

## 分階段
- **P1（已完成，可用）**：記錄。原因自訂 + 菸品清單（名稱／類別／常用預設）+ 記一根（先選菸品口味再選原因，時間自動）+ 今日紀錄列表（點→編輯／刪除）。設定頁管理原因與菸品。統計／庫存為佔位。
- **P2（已完成，可用）**：庫存。加熱菸/盒菸有「剩餘支數」，買入輸入盒數 → +（盒×每盒支數），記一根自動 −1，顯示「X 盒 Y 支」，某口味整體剩 1 支時卡片跳彩蛋「盒仔內賰尾枝」（襯線斜體 --warm）。捲菸有「未開包數」，買入 +包數；「開新包」建立使用中的菸草包（未開 −1），記一根（捲菸）要選一包 → 該包已捲 +1，「這包用完」歸檔記用完日；歷史顯示「共捲 N 支、每支成本 = 售價/N」。可同時多包使用中。開新包按鈕文案「Time to Roll」。刪除/編輯紀錄會**回沖庫存**（restore/consume）。addSmoke/updateSmoke/deleteSmoke/buyStock/openPouch/finishPouch 都回傳整包狀態 {records?,products,pouches}，前端 applyState() 一次刷新。
- **P3（已完成，可用）**：統計。segmented 日／週／月／年 + **LEDGER 式期間選擇器**（日/週用 `<input type=date>`、月用 `<input type=month>`、年用下拉；週選週內任一天，後端以週一換算整週）。選定期間後，趨勢圖顯示**該期間內的細分長條**：日→24 小時、週→7 天、月→當月每天、年→12 個月；疊移動平均折線（月 MA7、年 MA3；日/週不畫）。原因分佈＝水平排行長條（含 %）、時段分佈＝早午晚深夜 4 組，都只算該期間。metric：期間總支數、平均每天；期間起訖顯示在選擇器下方（`statRange`）。`getStats(gran,key)`，前端 `setGran()`/`doLoadStats()`，切到統計分頁預設 `month`＝本月。**花費估算延後到下一版（P4）**。
- **P4（未做）**：花費估算。

## 資料模型（試算表分頁）
- `設定`：key-value（時區、幣別、預設每盒支數）。讀 A:B。
- `原因`：名稱 | 排序 | 啟用。**setup() 不預設任何原因**（使用者自建）。**刪除＝硬刪整列**（下面上移）；改名（`renameReason`）會**回溯更新所有月分頁的原因文字**，統計才不會分裂；排序由 `reorderReasons(names)` 依前端傳入順序重寫「排序」欄，設定頁用上下箭頭調。歷史紀錄存原因文字快照，硬刪不影響歷史。「常用預設」功能已移除（P_DEFAULT 欄保留但不再使用）。
- `菸品`：id | 類別 | 名稱 | 每盒支數 | 售價 | 售價單位 | 剩餘支數 | 常用預設 | 狀態 | 建立時間 | 未開包數。捲菸的每盒支數留空；剩餘支數給加熱菸/盒菸、未開包數給捲菸。**刪除＝硬刪整列**（並順手刪掉該菸品「使用中」的菸草包，保留已用完歷史）。剩餘支數／未開包數可用 `setStock` 直接設定（處理半盒／修正）。（未開包數為 P2 新增第 11 欄，`ensureReady_→migrate_()` 會幫既有分頁補表頭。）
- `菸草包`（P2）：id | 菸品id | 口味 | 開封日 | 用完日 | 已捲支數 | 售價 | 狀態。
- `紀錄`：**一個月一分頁**，分頁名 `yyyy-MM`。欄位 id | 時間 | 原因 | 菸品id | 菸品名稱 | 類別 | 菸草包id | 成本 | 備註。原因與菸品名稱都存**快照**，之後改名／刪除不影響歷史。

## 固定欄位索引
`Code.js` 用 `P_*` / `R_*` / `RS_*` 常數定位欄，**絕不用 `HEADERS.length`**。加欄時排在尾端、更新常數即可。菸品／菸草包／紀錄的分頁一開始就建好 P2/P3 全部欄位，P1 只是留白。

## 後端函式（Code.js）
- `doGet()` / `setup()`（首次在編輯器跑一次）
- `getBootData()` → {settings, reasons, products, records(當月), thisMonth}
- 記錄：`addSmoke({productId,reason,pouchId?})`（時間＝當下）、`updateSmoke({id,month,reason?,productId?,pouchId?,timeMillis?})`、`deleteSmoke({id,month})`、`getMonthRecords(tab)`。捲菸必帶 pouchId；三者都連動庫存並回傳 `state_()`。**時間只能在「編輯這根」改**（datetime-local；記錄頁不出現）；改時間若跨月，updateSmoke 會把整列搬到正確月分頁。
- 原因：`addReason(name)`、`deleteReason(name)`（硬刪）、`renameReason(old,new)`（回溯歷史）、`reorderReasons(names)`
- 菸品：`addProduct(p)`、`updateProduct(p)`、`deleteProduct(id)`（硬刪＋刪使用中菸草包）、`setDefaultProduct(id)`
- 庫存（P2）：`buyStock({id,qty})`（加）、`setStock({id,value})`（設定絕對值：stick 支/捲菸 未開包數）、`openPouch({productId})`、`finishPouch({pouchId})`、`updatePouch({pouchId,rolled})`、`deletePouch({pouchId})`、`getPouches()`→{using,done}。連動 helper：`consumeInventory_/restoreInventory_/adjustProductLeft_/adjustPouchRolled_`。
- 統計（P3）：`getStats(gran)`→{labels,counts,ma,maWindow,reasons[],hourGroups[],total,perDay}。讀跨月分頁用 `readRecordsSince_/monthKeysBetween_`；桶/日期 helper：`bucketKey_/hourGroup_/movingAvg_/addDays_/addMonths_/mondayOf_/firstOfMonth_/fmt_`。前端用 Chart.js（head CDN 載入）。

前端只透過 `google.script.run` 溝通；**不可用 localStorage/sessionStorage**（GAS iframe 沙盒）。

## Gotchas（沿用 LEDGER 血淚）
- 日期樣字串寫進儲存格會被 Sheets 轉成 Date；要存字串旗標時 `setNumberFormat('@')` 並讀取時防禦。
- 建 Date 用當地正午避免時區跨日；輸出用 `Utilities.formatDate(d, TZ, …)`。
- 重用同一個部署 ID（v3：`clasp redeploy <id>`）→ `/exec` 網址不變。`/dev` 拿來當主畫面 App 不穩（要編輯者登入 session，私密瀏覽會壞），已放棄。
- 私有 GAS web app 沒辦法設主畫面自訂「圖片」圖示（沙盒 iframe）。iOS 會用**標題第一個字**自動生圖 → 所以 `doGet().setTitle('🚬 Time to Roll')`，主畫面圖示就是 🚬（保有私有＋全螢幕）。要換圖示 emoji 就改 setTitle 首字；重加時要先移除舊圖示、關掉 Safari 分頁再重開 /exec 加一次（iOS 會快取圖示）。

## 部署（clasp v3！）
手機主畫面用 **`/exec`（版本化部署）**，不是 `/dev`。更新兩步：`clasp push` → **`clasp redeploy <部署ID>`**（v3 把舊的 `clasp deploy -i` 改名成 `redeploy`／`update-deployment`；`deploy`／`create-deployment` 會建**新**部署＋新網址，別用）。沿用同一個部署 ID → `/exec` 不變，iPhone 關掉重開即新版。部署 ID 用 `clasp list-deployments` 看非 `@HEAD` 那筆。懶人版：`./update.sh`（會自動抓非 @HEAD 的部署 ID）。GitHub `git push` 只是備份。
正式部署 ID 用 `clasp list-deployments` 查（非 @HEAD 那筆）；不寫進 repo（跟 scriptId 一樣屬本機資訊）。
