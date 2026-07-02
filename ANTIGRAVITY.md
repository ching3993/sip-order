# sip-order 專案設定與規則 (ANTIGRAVITY.md)

## 專案中繼資料
- **專案名稱**：sip-order (飲料訂購程式)
- **主要工作目錄 (Cwd)**：`c:\Users\ching\projects\anti_new_070201`
- **技術棧**：Google Apps Script (GAS) Web App, Google 試算表 (Google Sheets), 單一 HTML (HTML5 + RWD CSS + Vanilla JS)
- **GitHub 倉庫**：[ching3993/sip-order](https://github.com/ching3993/sip-order) (Public)
- **專案駕駛艙**：`g:\我的雲端硬碟\secondbrain\sip-order 專案駕駛艙.md`

## 助理設定
- **助理名稱**：阿芝
- **對答語言**：繁體中文（Taiwan）
- **環境設定**：Windows PowerShell, 命令前綴使用 `npm.cmd` 或 `npx.cmd`。

## 檔案結構
- `Code.gs`：Google Apps Script 後端代碼 (doGet 與試算表寫入邏輯)
- `Index.html`：前端飲料訂購頁面 (單一 HTML 響應式 UI)
- `README.md`：專案說明與部署指南
- `ANTIGRAVITY.md`：本專案規則與中繼設定
- `.gitignore`：Git 忽略清單

## 專案工作流規則
1. **開工/收工流程**：遵守 `antigravity_workflow` 技能規範。
   - 開工時讀取 `ANTIGRAVITY.md` 與 Obsidian 專案駕駛艙。
   - 收工時掃描敏感資訊、更新駕駛艙，並以 Git 提交更新到 GitHub。
2. **安全規範**：
   - 嚴禁提交 any Google 帳號憑證、API 金鑰。
   - 試算表 ID 與連結在程式碼中請使用佔位符或容器試算表取得方法（`SpreadsheetApp.getActiveSpreadsheet()`），不要直接寫死私人試算表連結。
