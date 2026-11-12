# UPG Pallet 倉存管理

倉庫同事用嘅倉存網站：開 link 即用，唔使登入。

- **網址**：https://techyanlam.github.io/UPGL/（GitHub Pages 開通後生效）
- **在倉板數**：約 130 板（2026-11-11 核實）

## 功能

- **頁頂統計**：在倉板數、總體積 (CBM)、今日出倉數、待確認數
- **每日更新**頁：
  - 出倉登記：打板號任何一部分即時模糊搜尋（唔分大細楷、唔理連字號同空格，只列在倉板）；可同時入新高度；確認出倉
  - 待確認新貨：Lark 同步嚟嘅新板先入待確認清單，同事逐項按「確認加入」先正式入倉，唔啱嘅按「唔加入」剔除（可選填原因）；確認／剔除都會記入活動記錄
  - 在倉表：搜尋 Batch#／位置，每行可改高度（體積按 ROUNDUP 到 1 位小數重計）
  - 新板入倉：直接入倉，唔經待確認
  - 已出倉清單：可搜尋 Batch#／位置、按出倉日期篩選，可還原返在倉
- **活動記錄**頁：出倉／改高度／入倉／還原／待確認／確認加入／唔加入
- **Excel 版面**頁：欄位次序跟原檔（Daily Measurement Date、Timestamp、Dock Receipt Date、Business Unit、Reference、Batch#、Date of Measurement、Length、Width、Height、Volume、Ship Date、Remark）；可篩全部／在倉／已出倉＋搜尋 Batch#；底部有總板數同總體積；已出倉嘅板 Ship Date 顯示出倉日期、Remark 打「最後」，在倉留空
- **CSV 匯出**：按當時篩選匯出目前顯示嘅表（UTF-8 BOM，Excel 開到），檔名 `UPGL-YYYY-MM-DD.csv`

## 數據流

```
前端（GitHub Pages 靜態頁）→ Cloudflare Worker（upg-pallet-stock）→ KV
Lark Base 新貨 → GitHub Actions 同步 → Worker 待確認清單 → 同事確認 → 正式入倉
```

- 前端唔寫死任何板數據，全部經 Worker API 讀寫。
- Worker URL 喺 `config.js` 設定：`window.UPG_WORKER_URL = "…"`。
- API：
  - `GET {WORKER_URL}/state` → `{ inStock, departed, pending, activity, inStockCount, inStockVolumeCbm, todayDepartures }`
  - `POST {WORKER_URL}/` with `{ action, ... }`：`depart`、`updateHeight`、`add`、`restore`、`confirm`、`reject`、`sync`（單板入待確認）、`syncBatch`（多板一次入待確認，供 workflow 用）

## 同步範圍

- 只同步 **Labware** BU 嘅 **P-type** 板。
- 箱（C-type）唔入；OM Digital 唔入。

## 檔案

- `index.html` — 頁面結構
- `app.js` — 全部前端邏輯（搜尋、POST 操作、Excel 版面、CSV 匯出）
- `styles.css` — 樣式（手機都睇到）
- `config.js` — Worker URL 設定（得一行）
- `lark_pallet_sync.py` — Lark 同步腳本（讀 Dock Receipt → 寫 Worker 待確認清單）
- `docs/lark-pallet-sync-workflow.yml` — GitHub Actions workflow 檔（預備好，未放入 `.github/workflows/`）

## 注意

- GitHub App 唔可以直接 push `.github/workflows/` 檔案，`docs/lark-pallet-sync-workflow.yml` 要用網頁編輯器手動加入 `.github/workflows/lark-pallet-sync.yml`（repo → Add file → Create new file → 貼上內容 → Commit）。
- GitHub Pages 要喺 repo Settings → Pages 開通（Deploy from a branch，Branch `main`、路徑 `/ (root)`）。
