115年度高雄軟體園區CEO聯誼會
掃碼循環修正版

【問題原因】
原 Code.gs：
SCANNER_URL = https://shun197590.github.io/ceo-checkin/

但目前 GitHub 根目錄 index.html 是「CEO報到 PWA 首頁」，
不是掃碼器。
因此在 Apps Script 控制頁按「啟動連續掃碼器」時，
實際又打開 PWA 首頁，所以看起來像「返回啟動 Apps Script 報到控制頁」。

【本版修正】
1. GitHub 根目錄 index.html：保留 CEO報到 PWA 首頁。
2. 新增 scanner/index.html：專門放連續 QR 掃碼器。
3. Code.gs 的 SCANNER_URL 改為：
   https://shun197590.github.io/ceo-checkin/scanner/
4. Apps Script 控制頁 window.open 改用 _blank，避免重用錯誤分頁。
5. service-worker.js 更新為 v5，scanner 路徑強制 Network First / no-store。
6. PWA 圖示路徑改為目前 GitHub 根目錄的實際檔案位置。

【GitHub 正確結構】
ceo-checkin/
├─ index.html                  ← CEO報到 PWA 首頁
├─ manifest.json
├─ service-worker.js
├─ offline.html
├─ apple-touch-icon.png
├─ icon-192.png
├─ icon-512.png
├─ icon-maskable-512.png
└─ scanner/
   └─ index.html               ← 連續 QR 掃碼器

【Apps Script】
將 Code.gs 整份覆蓋目前 Apps Script 程式碼。

重要設定：
SCANNER_URL:
https://shun197590.github.io/ceo-checkin/scanner/

正式 Apps Script：
https://script.google.com/macros/s/AKfycbw472sz_13s3Sk5Y4oexnY2sSOUwjQuL44axG4blYIzkvGJmx_CPoiu2K5TgLmIM1SK/exec

報名狀況試算表：
https://docs.google.com/spreadsheets/d/1-EFuZGivOEESxtwmNdjy8krUBSCNhrsFWylJjm90CAo/edit?gid=1412860227#gid=1412860227

【完成後測試】
1. 開啟 CEO報到 PWA。
2. 按「啟動 Apps Script 報到控制頁」。
3. 按「測試 Apps Script 後端」。
4. 按「啟動連續掃碼器」。
5. 此時網址應包含：
   /ceo-checkin/scanner/?controlOrigin=...&sessionToken=...
6. 掃碼頁應顯示「正在連線 Apps Script 控制頁…」後變為已連線。
7. 按「啟用相機＋測試報到語音」。

【Apps Script 更新部署】
修改 Code.gs 後，請到：
部署 → 管理部署作業 → 編輯目前 Web App 部署 → 建立新版本 → 部署。
不要只儲存程式碼而沒有更新部署版本。
