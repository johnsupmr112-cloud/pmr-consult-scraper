# PMR 會診清單擷取腳本

在院內會診系統的清單頁面，於瀏覽器 Console 執行，自動抓取 PMR 會診病人與明細頁「病史」，直接下載排版好、可直接列印的 Excel。

- 純 JavaScript，無任何外部相依（不需 CDN、Python 或安裝套件），適用院內網路環境
- 自帶 ZIP / OOXML 產生器，直接輸出 `.xlsx`

## 輸出

`PMR_YYYYMMDD.xlsx`

| 病房床號 | 病歷號 | 病人姓名 | 會診內容 | 病人科別 | 會診時限 | 被會醫師 |
|---|---|---|---|---|---|---|

- 新細明體 12pt，表頭粗體置中、淺綠底色，細黑框線
- 會診內容靠左、靠上、自動換行，列高預先計算
- A4 橫式、縮放至一頁寬、凍結首列、每頁重複表頭 → 開啟後直接 Ctrl+P

## 用法

1. 在會診系統查詢出清單頁面
2. `F12` → Console → 貼上 `pmr_consult_scraper.js` 全文 → Enter
3. 等進度跑完，瀏覽器自動下載 xlsx

> 第一次使用請先把 `CONFIG.diagnose` 設為 `true`，確認「病史」欄位抓對位置。

## 主要設定（`CONFIG`）

| 參數 | 預設 | 說明 |
|---|---|---|
| `onlyPMR` | `true` | 只保留「被會科別」含 PMR 的列 |
| `statusFilter` | `null` | `null` / `'待會'` / `'V'` |
| `sortBy` / `sortDesc` | `'會診時限'` / `true` | 排序欄位與方向 |
| `dedupe` | `true` | 以「病歷號 + 會診時限」去重 |
| `detailLabels` | `['病史','會診原因','會診目的']` | 明細頁依序嘗試的欄位標題 |
| `delayMs` | `400` | 每筆明細頁請求間隔（毫秒），避免造成伺服器負擔 |
| `maxRows` | `0` | 0 = 不限；測試時可設 3 |
| `diagnose` | `false` | 只抓第一筆並印出明細頁結構與原始病史 |
| `preserveLineBreaks` | `true` | 保留段落結構 |
| `unwrapHardBreaks` | `true` | 還原固定寬度硬切造成的斷行 |
| `joinBrokenParagraphs` | `true` | 合併把句子切成兩半的假空行 |
| `headerFill` | `'E2EFDA'` | 表頭底色；`null` = 無底色 |
| `borderColor` | `'000000'` | 框線顏色 |
| `freezeHeader` / `repeatHeaderOnEveryPage` | `true` | 凍結首列 / 每頁重複表頭 |

## 疑難排解

- **找不到 table**：表格可能在 iframe 內，先在 Console 左上角切換到該 frame 再執行。
- **完全沒抓到會診內容**：設 `diagnose: true` 跑一次，看明細頁的「標題 → 內容」對照表，調整 `detailLabels`。
- **沒有符合條件的資料**：試著把 `onlyPMR` 改成 `false`。

## 注意

輸出的 Excel 含病人個資，請依院內規定保存與銷毀，切勿上傳至本 repo。
